"""Link Excel purchase references to Odoo purchase orders.

Reads N°FACTURE from EXCEL FACTURE ACHATS LOCAL.xlsx and N COMMANDE from
TRACKING ACHATS ETRANGER (1).xlsx, searches for matching records in Odoo
purchase.order using the partner_ref field (Référence fournisseur), and adds
hyperlinks back to the Excel cells.

By default runs in dry-run mode; pass --apply to update the workbooks.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import getpass
import os
import re
import shutil
import sys
import threading
import time
import xmlrpc.client
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import urlencode

try:
    import pythoncom  # type: ignore[import-not-found]
except ImportError:
    pythoncom = None

try:
    import win32com.client  # type: ignore[import-not-found]
except ImportError:
    win32com = None

try:
    import xlwings as xw
except ImportError:
    xw = None

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DEFAULT_ODOO_URL = "https://sphe.cloudoo.ma"
DEFAULT_LOCAL_WORKBOOK = Path(__file__).parent / "excel" / "EXCEL FACTURE ACHATS LOCAL.xlsx"
DEFAULT_ETRANGER_WORKBOOK = Path(__file__).parent / "excel" / "TRACKING ACHATS ETRANGER (1).xlsx"

# LOCAL file: header is row 1, search column = "N°FACTURE"
LOCAL_HEADER_VARIANTS = {"n°facture", "n° facture", "nfacture", "n facture"}
# ETRANGER file: header is row 8, search column = "N COMMANDE"
ETRANGER_HEADER_VARIANTS = {"n commande", "n° commande", "ncommande"}

PURCHASE_ORDER_FIELDS = [
    "name", "partner_ref", "partner_id", "state", "date_order",
    "amount_total", "origin",
]

REPORT_COLUMNS = (
    "source_file", "sheet", "cell", "reference", "status",
    "source_model", "matched_field", "record_id", "record_name",
    "ref_value", "state", "vendor", "amount", "url", "note",
)

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class PurchaseRefCell:
    source_file: str  # "local" or "etranger"
    sheet: str
    row: int
    column: int
    address: str
    ref_value: str


@dataclass(frozen=True)
class PurchaseLinkResult:
    status: str
    source_model: str = ""
    matched_field: str = ""
    record_id: int | None = None
    record_name: str = ""
    ref_value: str = ""
    state: str = ""
    vendor: str = ""
    amount: float = 0.0
    url: str = ""
    note: str = ""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def cell_address(row: int, col: int) -> str:
    letters = ""
    n = col
    while n:
        n, remainder = divmod(n - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def many2one_name(value: Any) -> str:
    if isinstance(value, (list, tuple)) and len(value) >= 2:
        return str(value[1] or "")
    return str(value or "")


def make_timestamp() -> str:
    return dt.datetime.now().strftime("%Y%m%d-%H%M%S")


@contextmanager
def com_scope():
    if pythoncom is None:
        raise RuntimeError("pywin32 is required. Install with: python -m pip install pywin32")
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def open_excel(visible: bool) -> Any:
    if win32com is None:
        raise RuntimeError("pywin32 is required. Install with: python -m pip install pywin32")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = bool(visible)
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    return excel


# ---------------------------------------------------------------------------
# Odoo Client for Purchases
# ---------------------------------------------------------------------------

class OdooPurchaseClient:
    def __init__(self, url: str, db: str, login: str, api_key: str) -> None:
        self.url = url.rstrip("/")
        self.db = db
        self.login = login
        self.api_key = api_key
        self.uid: int | None = None
        transport = (xmlrpc.client.SafeTransport() if self.url.startswith("https")
                     else xmlrpc.client.Transport())
        transport._extra_headers = [("Connection", "keep-alive")]
        self.common = xmlrpc.client.ServerProxy(
            f"{self.url}/xmlrpc/2/common", transport=transport)
        self.models = xmlrpc.client.ServerProxy(
            f"{self.url}/xmlrpc/2/object", transport=transport)

    def authenticate(self) -> int:
        uid = self.common.authenticate(self.db, self.login, self.api_key, {})
        if not uid:
            raise RuntimeError("Odoo authentication failed.")
        self.uid = int(uid)
        return self.uid

    def _call(self, model: str, method: str, args: list, kwargs: dict | None = None) -> Any:
        if self.uid is None:
            self.authenticate()
        return self.models.execute_kw(
            self.db, self.uid, self.api_key, model, method, args, kwargs or {})

    def search_purchase_orders_by_partner_ref(self, ref: str, operator: str) -> list[dict]:
        domain: list[Any] = [
            ("partner_ref", operator, ref),
        ]
        return self._call(
            "purchase.order",
            "search_read",
            [domain],
            {"fields": PURCHASE_ORDER_FIELDS, "limit": 10, "order": "date_order desc, id desc"},
        )

    def resolve_purchase_ref(self, ref: str) -> PurchaseLinkResult:
        """Search for a reference in purchase orders using partner_ref."""
        term = ref.strip()
        if not term:
            return PurchaseLinkResult(status="empty_ref", note="Reference was empty.")

        stages: list[tuple[str, str]] = [
            ("=", "exact:partner_ref"),
            ("ilike", "contains:partner_ref"),
        ]
        for operator, default_match in stages:
            try:
                orders = self.search_purchase_orders_by_partner_ref(term, operator)
            except Exception:
                orders = []
            if not orders:
                continue
            best = self._pick_best_order(term, orders, default_match=default_match)
            if best is not None:
                return best

        return PurchaseLinkResult(
            status="not_found",
            note="Reference not found in Odoo purchase orders (partner_ref).")

    def _pick_best_order(
        self,
        term: str,
        orders: list[dict],
        *,
        default_match: str,
    ) -> PurchaseLinkResult | None:
        folded = term.casefold()
        exact = [
            o
            for o in orders
            if str(o.get("partner_ref") or "").strip().casefold() == folded
        ]
        candidates = exact if exact else orders
        matched = "exact:partner_ref" if exact else default_match
        if not candidates:
            return None
        best = candidates[0]
        po_id = int(best["id"])
        return PurchaseLinkResult(
            status="linked",
            source_model="purchase.order",
            matched_field=matched,
            record_id=po_id,
            record_name=str(best.get("name") or ""),
            ref_value=str(best.get("partner_ref") or ""),
            state=str(best.get("state") or ""),
            vendor=many2one_name(best.get("partner_id")),
            amount=float(best.get("amount_total") or 0),
            url=f"{self.url}/odoo/purchase/{po_id}",
            note=f"Matched purchase order via {matched}.",
        )


# ---------------------------------------------------------------------------
# Excel scanning
# ---------------------------------------------------------------------------

def _find_header_col(all_values: tuple, row_count: int, col_count: int,
                     header_variants: set[str]) -> tuple[int | None, int | None]:
    """Return (header_row_0indexed, header_col_0indexed) or (None, None)."""
    for ri in range(min(row_count, 80)):  # scan first 80 rows for headers
        for ci in range(col_count):
            norm = normalize_header(all_values[ri][ci])
            if norm in header_variants:
                return ri, ci
    return None, None


def collect_refs_from_workbook(workbook: Any, source_file: str,
                               header_variants: set[str],
                               target_sheets: list[str] | None = None,
                               ) -> list[PurchaseRefCell]:
    """Extract reference values from a workbook using COM batch read."""
    cells: list[PurchaseRefCell] = []
    sheet_map = {sheet.Name: sheet for sheet in workbook.Worksheets}
    sheets_to_scan = target_sheets or list(sheet_map.keys())

    for sheet_name in sheets_to_scan:
        sheet = sheet_map.get(sheet_name)
        if sheet is None:
            continue
        used_range = sheet.UsedRange
        all_values = used_range.Value
        if all_values is None:
            continue
        if not isinstance(all_values, tuple):
            all_values = ((all_values,),)
        row_count = len(all_values)
        if row_count == 0:
            continue
        if not isinstance(all_values[0], tuple):
            all_values = tuple((v,) for v in all_values)
            row_count = len(all_values)
        col_count = len(all_values[0])
        first_row = int(used_range.Row)
        first_col = int(used_range.Column)

        header_ri, header_ci = _find_header_col(
            all_values, row_count, col_count, header_variants)
        if header_ri is None or header_ci is None:
            continue

        for ri in range(header_ri + 1, row_count):
            raw = all_values[ri][header_ci]
            val = str(raw).strip() if raw is not None else ""
            if val and val.lower() not in ("none", "", "nan"):
                actual_row = first_row + ri
                actual_col = first_col + header_ci
                cells.append(PurchaseRefCell(
                    source_file=source_file,
                    sheet=sheet_name,
                    row=actual_row,
                    column=actual_col,
                    address=cell_address(actual_row, actual_col),
                    ref_value=val,
                ))
    return cells


def scan_workbook_refs(workbook_path: Path, source_file: str,
                       header_variants: set[str],
                       visible: bool = False) -> list[PurchaseRefCell]:
    """Open a workbook, scan for references, and close."""
    excel = open_excel(visible)
    wb = None
    try:
        wb = excel.Workbooks.Open(str(workbook_path), ReadOnly=True,
                                  UpdateLinks=0, IgnoreReadOnlyRecommended=True)
        return collect_refs_from_workbook(wb, source_file, header_variants)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()


# ---------------------------------------------------------------------------
# Batch resolution
# ---------------------------------------------------------------------------

def resolve_refs(ref_values: list[str],
                 client: OdooPurchaseClient) -> dict[str, PurchaseLinkResult]:
    """Resolve all unique references against Odoo in parallel."""
    unique = list(dict.fromkeys(ref_values))
    results: dict[str, PurchaseLinkResult] = {}

    def _resolve_one(ref: str) -> tuple[str, PurchaseLinkResult]:
        return ref, client.resolve_purchase_ref(ref)

    max_workers = min(8, max(1, len(unique)))
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_resolve_one, ref): ref for ref in unique}
        for future in as_completed(futures):
            ref, result = future.result()
            results[ref] = result

    return results


# ---------------------------------------------------------------------------
# Hyperlink writer
# ---------------------------------------------------------------------------

def apply_links(workbook: Any, cells: list[PurchaseRefCell],
                results: dict[str, PurchaseLinkResult]) -> int:
    """Write hyperlinks into the workbook for linked results."""
    app = workbook.Application
    prev_screen = prev_calc = prev_events = None
    try:
        try:
            prev_screen = app.ScreenUpdating
            prev_calc = app.Calculation
            prev_events = app.EnableEvents
            app.ScreenUpdating = False
            app.Calculation = -4135  # xlCalculationManual
            app.EnableEvents = False
        except Exception:
            pass

        count = 0
        sheet_map = {s.Name: s for s in workbook.Worksheets}
        for cell_info in cells:
            result = results.get(cell_info.ref_value)
            if result is None or result.status != "linked" or not result.url:
                continue
            sheet = sheet_map.get(cell_info.sheet)
            if sheet is None:
                continue
            cell = sheet.Cells(cell_info.row, cell_info.column)
            try:
                cell.Hyperlinks.Delete()
            except Exception:
                pass
            tip = f"Open {result.source_model} in Odoo"
            sheet.Hyperlinks.Add(
                Anchor=cell, Address=result.url,
                TextToDisplay=cell_info.ref_value, ScreenTip=tip)
            count += 1
        return count
    finally:
        try:
            if prev_events is not None:
                app.EnableEvents = prev_events
            if prev_calc is not None:
                app.Calculation = prev_calc
            if prev_screen is not None:
                app.ScreenUpdating = prev_screen
        except Exception:
            pass


def write_links_to_file(workbook_path: Path, cells: list[PurchaseRefCell],
                        results: dict[str, PurchaseLinkResult],
                        visible: bool = False) -> int:
    excel = open_excel(visible)
    wb = None
    try:
        wb = excel.Workbooks.Open(str(workbook_path), ReadOnly=False,
                                  UpdateLinks=0, IgnoreReadOnlyRecommended=True)
        if bool(wb.ReadOnly):
            raise RuntimeError("Workbook opened as read-only.")
        count = apply_links(wb, cells, results)
        wb.Save()
        return count
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def write_report(report_path: Path, cells: list[PurchaseRefCell],
                 results: dict[str, PurchaseLinkResult]) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        for cell in cells:
            r = results.get(cell.ref_value, PurchaseLinkResult(status="unknown"))
            writer.writerow({
                "source_file": cell.source_file,
                "sheet": cell.sheet, "cell": cell.address,
                "reference": cell.ref_value, "status": r.status,
                "source_model": r.source_model, "matched_field": r.matched_field,
                "record_id": r.record_id or "", "record_name": r.record_name,
                "ref_value": r.ref_value, "state": r.state,
                "vendor": r.vendor, "amount": r.amount,
                "url": r.url, "note": r.note,
            })


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--local-workbook", default=str(DEFAULT_LOCAL_WORKBOOK),
                   help="Path to EXCEL FACTURE ACHATS LOCAL.xlsx")
    p.add_argument("--etranger-workbook", default=str(DEFAULT_ETRANGER_WORKBOOK),
                   help="Path to TRACKING ACHATS ETRANGER (1).xlsx")
    p.add_argument("--skip-local", action="store_true",
                   help="Skip scanning the local purchases file.")
    p.add_argument("--skip-etranger", action="store_true",
                   help="Skip scanning the foreign purchases file.")
    p.add_argument("--odoo-url", default=os.getenv("ODOO_URL", DEFAULT_ODOO_URL))
    p.add_argument("--odoo-db", default=os.getenv("ODOO_DB"))
    p.add_argument("--odoo-login", default=os.getenv("ODOO_LOGIN"))
    p.add_argument("--odoo-api-key", default=os.getenv("ODOO_API_KEY"))
    p.add_argument("--report", default=None, help="CSV report path.")
    p.add_argument("--test-orders", nargs="*", default=None,
                   help="Only resolve these references (no Excel scan).")
    p.add_argument("--apply", action="store_true",
                   help="Create backups and update the workbooks with hyperlinks.")
    p.add_argument("--visible-excel", action="store_true")
    p.add_argument("--prompt-secret", action="store_true",
                   help="Prompt for API key if not set.")
    return p.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)

    # Validate config
    missing = []
    if not args.odoo_db:
        missing.append("ODOO_DB / --odoo-db")
    if not args.odoo_login:
        missing.append("ODOO_LOGIN / --odoo-login")
    if not args.odoo_api_key:
        if args.prompt_secret:
            args.odoo_api_key = getpass.getpass("Odoo API key: ")
        if not args.odoo_api_key:
            missing.append("ODOO_API_KEY / --odoo-api-key")
    if missing:
        raise SystemExit("Missing: " + ", ".join(missing))

    client = OdooPurchaseClient(args.odoo_url, args.odoo_db,
                                args.odoo_login, args.odoo_api_key)
    client.authenticate()
    print(f"Authenticated to {args.odoo_url} as {args.odoo_login}")

    # --- Test mode ---
    if args.test_orders is not None:
        test_refs = [r.strip() for r in args.test_orders if r.strip()]
        if not test_refs:
            raise SystemExit("--test-orders requires at least one reference.")
        results = resolve_refs(test_refs, client)
        for ref, result in results.items():
            print(f"\n  {ref}: {result.status}")
            if result.record_name:
                print(f"    Record: {result.record_name} ({result.source_model})")
            if result.url:
                print(f"    URL: {result.url}")
            if result.note:
                print(f"    Note: {result.note}")
        if args.report:
            # Build minimal cells for report
            test_cells = [PurchaseRefCell("test", "", 0, 0, "", r) for r in test_refs]
            write_report(Path(args.report), test_cells, results)
            print(f"\nReport: {args.report}")
        return 0

    # --- Full scan mode ---
    all_cells: list[PurchaseRefCell] = []
    local_cells: list[PurchaseRefCell] = []
    etranger_cells: list[PurchaseRefCell] = []

    with com_scope():
        if not args.skip_local:
            local_path = Path(args.local_workbook).expanduser().resolve()
            if local_path.exists():
                print(f"Scanning LOCAL: {local_path.name} ...")
                local_cells = scan_workbook_refs(
                    local_path, "local", LOCAL_HEADER_VARIANTS, args.visible_excel)
                print(f"  Found {len(local_cells)} N°FACTURE references")
                all_cells.extend(local_cells)
            else:
                print(f"WARNING: Local workbook not found: {local_path}")

        if not args.skip_etranger:
            etr_path = Path(args.etranger_workbook).expanduser().resolve()
            if etr_path.exists():
                print(f"Scanning ETRANGER: {etr_path.name} ...")
                etranger_cells = scan_workbook_refs(
                    etr_path, "etranger", ETRANGER_HEADER_VARIANTS, args.visible_excel)
                print(f"  Found {len(etranger_cells)} N COMMANDE references")
                all_cells.extend(etranger_cells)
            else:
                print(f"WARNING: Etranger workbook not found: {etr_path}")

    if not all_cells:
        print("No references found in any workbook.")
        return 0

    unique_refs = list(dict.fromkeys(c.ref_value for c in all_cells))
    print(f"\nTotal: {len(all_cells)} cells, {len(unique_refs)} unique references")
    print("Resolving against Odoo ...")

    results = resolve_refs(unique_refs, client)

    # Summarize
    counts: dict[str, int] = {}
    for r in results.values():
        counts[r.status] = counts.get(r.status, 0) + 1
    summary = ", ".join(f"{k}={v}" for k, v in sorted(counts.items()))
    print(f"Results: {summary}")

    # Write report
    report_path = (Path(args.report) if args.report
                   else Path(__file__).parent / "excel"
                   / f"purchase-link-report-{make_timestamp()}.csv")
    write_report(report_path, all_cells, results)
    print(f"Report: {report_path}")

    # Apply hyperlinks
    if not args.apply:
        print("\nDry-run. Re-run with --apply to update workbooks.")
        return 0

    linked_total = 0
    with com_scope():
        if local_cells:
            local_path = Path(args.local_workbook).expanduser().resolve()
            backup = local_path.parent / f"{local_path.stem}.backup-{make_timestamp()}{local_path.suffix}"
            shutil.copy2(local_path, backup)
            print(f"Backup LOCAL: {backup.name}")
            n = write_links_to_file(local_path, local_cells, results, args.visible_excel)
            print(f"  Linked {n} cells in LOCAL workbook")
            linked_total += n

        if etranger_cells:
            etr_path = Path(args.etranger_workbook).expanduser().resolve()
            backup = etr_path.parent / f"{etr_path.stem}.backup-{make_timestamp()}{etr_path.suffix}"
            shutil.copy2(etr_path, backup)
            print(f"Backup ETRANGER: {backup.name}")
            n = write_links_to_file(etr_path, etranger_cells, results, args.visible_excel)
            print(f"  Linked {n} cells in ETRANGER workbook")
            linked_total += n

    print(f"\nTotal hyperlinks added: {linked_total}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
