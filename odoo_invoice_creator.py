"""Create Odoo vendor bills from Excel order data.

Reads rows from the target workbook sheets, groups them by supplier,
and creates ``account.move`` records (vendor bills) in Odoo via XML-RPC.
Bills are created in **draft** state by default so they can be reviewed
before posting.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Any

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column header normalisation (matches the style in link_odoo_vendor_bills)
# ---------------------------------------------------------------------------

_HEADER_ALIASES: dict[str, str] = {
    "client": "client",
    "référence commande": "order_ref",
    "reference commande": "order_ref",
    "n commandes": "order_number",
    "achteur": "buyer",
    "acheteur": "buyer",
    "responsaple": "responsible",
    "responsable": "responsible",
    "date de commande": "order_date",
    "date complet de commande": "order_date",
    "observation": "observation",
    "montant total h.t": "amount_ht",
    "montant total ht": "amount_ht",
    "référence pas encore livrer": "part_reference",
    "reference pas encore livrer": "part_reference",
    "qte": "quantity",
    "fournisseur": "supplier",
    "fournisseur /num facture": "supplier",
    "fournisseur /num facture ": "supplier",
    "fournisseur / num facture": "supplier",
    "colonne1": "_skip",
    "status": "status",
    "status ": "status",
    "status 2": "status2",
    "status 2 ": "status2",
}


def _norm_header(val: Any) -> str:
    """Lowercase, strip, collapse whitespace."""
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).strip().lower())


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class ExcelInvoiceRow:
    """One row of invoice-relevant data extracted from Excel."""
    sheet: str
    row: int
    client: str = ""
    order_ref: str = ""
    order_number: str = ""
    buyer: str = ""
    responsible: str = ""
    order_date: str = ""
    observation: str = ""
    amount_ht: float = 0.0
    part_reference: str = ""
    quantity: float = 1.0
    supplier: str = ""
    status: str = ""
    status2: str = ""


@dataclass
class VendorBillResult:
    """Result of creating a single vendor bill."""
    move_id: int | None = None
    bill_name: str = ""
    vendor_name: str = ""
    vendor_id: int | None = None
    line_count: int = 0
    total_amount: float = 0.0
    url: str = ""
    status: str = "pending"  # pending | created | error | skipped
    error: str = ""
    source_rows: list[ExcelInvoiceRow] = field(default_factory=list)


@dataclass
class InvoiceCreationSummary:
    """Summary of a batch bill-creation run."""
    total_rows: int = 0
    bills_created: int = 0
    bills_skipped: int = 0
    bills_errored: int = 0
    results: list[VendorBillResult] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Excel data extraction
# ---------------------------------------------------------------------------

def _parse_amount(raw: Any) -> float:
    """Parse a monetary amount from Excel cell value."""
    if raw is None:
        return 0.0
    if isinstance(raw, (int, float)):
        return float(raw)
    text = str(raw).replace("\xa0", "").replace(" ", "").strip()
    # Handle European format: 15.200,00 or 15 200,00
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except (ValueError, TypeError):
        return 0.0


def _parse_quantity(raw: Any) -> float:
    """Parse quantity from Excel, handling formats like '1 /1' or '2.0'."""
    if raw is None:
        return 1.0
    if isinstance(raw, (int, float)):
        return float(raw)
    text = str(raw).strip()
    # Handle "1 /1" or "1/1" — take the first number
    if "/" in text:
        text = text.split("/")[0].strip()
    try:
        return float(text)
    except (ValueError, TypeError):
        return 1.0


def _parse_date(raw: Any) -> str:
    """Parse a date into YYYY-MM-DD string."""
    if raw is None:
        return ""
    text = str(raw).strip()
    # Handle "2026-02-27 00:00:00+00:00"
    if " " in text:
        text = text.split(" ")[0]
    # Basic validation
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text):
        return text
    return text


def collect_invoice_data(workbook: Any, target_sheets: tuple[str, ...] | None = None) -> list[ExcelInvoiceRow]:
    """Read all invoice-relevant data from workbook sheets in a single batch.

    Uses the same ``UsedRange.Value`` batch-read technique for speed.
    """
    if target_sheets is None:
        from link_odoo_vendor_bills import TARGET_SHEETS
        target_sheets = TARGET_SHEETS

    rows: list[ExcelInvoiceRow] = []
    sheet_map = {sheet.Name: sheet for sheet in workbook.Worksheets}

    for sheet_name in target_sheets:
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

        # Map header columns
        col_map: dict[int, str] = {}
        for ci in range(col_count):
            norm = _norm_header(all_values[0][ci])
            field_name = _HEADER_ALIASES.get(norm)
            if field_name and field_name != "_skip":
                col_map[ci] = field_name

        if not col_map:
            continue

        # Read data rows
        for ri in range(1, row_count):
            row_data: dict[str, Any] = {}
            all_empty = True
            for ci, field_name in col_map.items():
                val = all_values[ri][ci]
                if val is not None:
                    all_empty = False
                row_data[field_name] = val
            if all_empty:
                continue

            invoice_row = ExcelInvoiceRow(
                sheet=sheet_name,
                row=first_row + ri,
                client=str(row_data.get("client") or "").strip(),
                order_ref=str(row_data.get("order_ref") or "").strip(),
                order_number=str(row_data.get("order_number") or "").strip(),
                buyer=str(row_data.get("buyer") or "").strip(),
                responsible=str(row_data.get("responsible") or "").strip(),
                order_date=_parse_date(row_data.get("order_date")),
                observation=str(row_data.get("observation") or "").strip(),
                amount_ht=_parse_amount(row_data.get("amount_ht")),
                part_reference=str(row_data.get("part_reference") or "").strip(),
                quantity=_parse_quantity(row_data.get("quantity")),
                supplier=str(row_data.get("supplier") or "").strip(),
                status=str(row_data.get("status") or "").strip(),
                status2=str(row_data.get("status2") or "").strip(),
            )
            rows.append(invoice_row)

    return rows


# ---------------------------------------------------------------------------
# Grouping logic
# ---------------------------------------------------------------------------

def group_rows_into_bills(rows: list[ExcelInvoiceRow]) -> list[list[ExcelInvoiceRow]]:
    """Group Excel rows into vendor bill groups.

    Rows with the same (supplier, order_number) are grouped into one bill.
    Rows without a supplier are placed in individual groups.
    """
    groups: dict[tuple[str, str], list[ExcelInvoiceRow]] = {}
    ungrouped: list[list[ExcelInvoiceRow]] = []

    for row in rows:
        supplier = row.supplier.strip().upper()
        if not supplier or supplier in ("", "STANDBY", "-"):
            # No supplier — skip or create individual
            ungrouped.append([row])
            continue
        key = (supplier, row.order_number or row.order_ref or f"row-{row.row}")
        groups.setdefault(key, []).append(row)

    result = list(groups.values()) + ungrouped
    return result


# ---------------------------------------------------------------------------
# Odoo API operations
# ---------------------------------------------------------------------------

class OdooInvoiceCreator:
    """Creates vendor bills in Odoo from Excel data."""

    def __init__(self, client: Any):
        """*client* is an ``OdooClient`` instance from link_odoo_vendor_bills."""
        self.client = client
        self._vendor_cache: dict[str, int] = {}
        self._product_cache: dict[str, int | None] = {}

    # -- Vendor lookup / creation ------------------------------------------

    def find_vendor(self, name: str) -> int | None:
        """Search for a vendor (res.partner) by name. Returns ID or None."""
        if not name:
            return None
        folded = name.strip().upper()
        if folded in self._vendor_cache:
            return self._vendor_cache[folded]

        try:
            results = self.client.execute_kw(
                "res.partner", "search_read",
                [[["name", "ilike", name], ["supplier_rank", ">", 0]]],
                {"fields": ["id", "name"], "limit": 5},
            )
        except Exception:
            # supplier_rank may not exist in older Odoo versions
            try:
                results = self.client.execute_kw(
                    "res.partner", "search_read",
                    [[["name", "ilike", name]]],
                    {"fields": ["id", "name"], "limit": 5},
                )
            except Exception:
                return None

        if not results:
            return None

        # Prefer exact match
        for r in results:
            if str(r.get("name", "")).strip().upper() == folded:
                self._vendor_cache[folded] = int(r["id"])
                return int(r["id"])

        # Use first result
        partner_id = int(results[0]["id"])
        self._vendor_cache[folded] = partner_id
        return partner_id

    def create_vendor(self, name: str) -> int:
        """Create a new vendor partner in Odoo."""
        vals = {
            "name": name.strip(),
            "supplier_rank": 1,
            "company_type": "company",
        }
        try:
            partner_id = self.client.execute_kw(
                "res.partner", "create", [vals],
            )
        except Exception:
            # supplier_rank may fail — try without
            vals.pop("supplier_rank", None)
            partner_id = self.client.execute_kw(
                "res.partner", "create", [vals],
            )
        self._vendor_cache[name.strip().upper()] = int(partner_id)
        return int(partner_id)

    def find_or_create_vendor(self, name: str) -> int | None:
        """Find vendor by name, create if not found."""
        if not name or name.strip().upper() in ("", "STANDBY", "-"):
            return None
        vendor_id = self.find_vendor(name)
        if vendor_id is not None:
            return vendor_id
        return self.create_vendor(name)

    # -- Bill creation -----------------------------------------------------

    def create_vendor_bill(
        self,
        rows: list[ExcelInvoiceRow],
        *,
        auto_post: bool = False,
    ) -> VendorBillResult:
        """Create a single vendor bill from grouped Excel rows.

        Returns a ``VendorBillResult`` with the created bill info.
        """
        result = VendorBillResult(source_rows=list(rows))

        if not rows:
            result.status = "skipped"
            result.error = "No rows provided."
            return result

        first_row = rows[0]
        supplier_name = first_row.supplier.strip()

        if not supplier_name or supplier_name.upper() in ("STANDBY", "-"):
            result.status = "skipped"
            result.vendor_name = supplier_name
            result.error = "No valid supplier name."
            return result

        # Find or create vendor
        try:
            vendor_id = self.find_or_create_vendor(supplier_name)
        except Exception as exc:
            result.status = "error"
            result.vendor_name = supplier_name
            result.error = f"Failed to find/create vendor: {exc}"
            return result

        if vendor_id is None:
            result.status = "skipped"
            result.vendor_name = supplier_name
            result.error = "Could not resolve vendor."
            return result

        result.vendor_id = vendor_id
        result.vendor_name = supplier_name

        # Build invoice lines
        invoice_lines: list[tuple[int, int, dict[str, Any]]] = []
        total = 0.0
        for row in rows:
            line_name = row.part_reference or row.order_ref or row.order_number or "Service"
            qty = row.quantity if row.quantity > 0 else 1.0
            price = row.amount_ht
            if len(rows) > 1 and row.amount_ht > 0 and row.quantity > 0:
                # If multiple lines, price_unit = amount / qty
                price = row.amount_ht / qty

            line_vals: dict[str, Any] = {
                "name": line_name,
                "quantity": qty,
                "price_unit": price,
            }
            invoice_lines.append((0, 0, line_vals))
            total += row.amount_ht

        if not invoice_lines:
            result.status = "skipped"
            result.error = "No valid invoice lines."
            return result

        # Build bill values
        bill_vals: dict[str, Any] = {
            "move_type": "in_invoice",
            "partner_id": vendor_id,
            "invoice_line_ids": invoice_lines,
        }

        # Optional fields
        ref_parts = []
        if first_row.order_number:
            ref_parts.append(first_row.order_number)
        if first_row.order_ref:
            ref_parts.append(first_row.order_ref)
        if ref_parts:
            bill_vals["ref"] = " / ".join(ref_parts)

        if first_row.order_date:
            bill_vals["invoice_date"] = first_row.order_date

        observations = [r.observation for r in rows if r.observation]
        if observations:
            bill_vals["narration"] = "\n".join(observations)

        # Create the bill
        try:
            move_id = self.client.execute_kw(
                "account.move", "create", [bill_vals],
            )
            move_id = int(move_id)
        except Exception as exc:
            result.status = "error"
            result.error = f"Failed to create bill: {exc}"
            log.error("Failed to create vendor bill for %s: %s", supplier_name, exc)
            return result

        result.move_id = move_id
        result.line_count = len(invoice_lines)
        result.total_amount = total
        result.status = "created"

        # Read back the bill name
        try:
            bill_data = self.client.execute_kw(
                "account.move", "read", [[move_id]],
                {"fields": ["name"]},
            )
            if bill_data:
                result.bill_name = str(bill_data[0].get("name", ""))
        except Exception:
            pass

        # Auto-post if requested
        if auto_post:
            try:
                self.client.execute_kw(
                    "account.move", "action_post", [[move_id]],
                )
            except Exception as exc:
                log.warning("Failed to post bill %s: %s", move_id, exc)

        log.info(
            "Created vendor bill %s (ID=%s) for %s — %d lines, total=%.2f",
            result.bill_name, move_id, supplier_name, len(invoice_lines), total,
        )
        return result

    # -- Batch creation ----------------------------------------------------

    def create_bills_from_excel(
        self,
        rows: list[ExcelInvoiceRow],
        *,
        auto_post: bool = False,
    ) -> InvoiceCreationSummary:
        """Create vendor bills for all grouped rows.

        Returns a summary with results for each bill.
        """
        summary = InvoiceCreationSummary(total_rows=len(rows))
        groups = group_rows_into_bills(rows)

        for group in groups:
            result = self.create_vendor_bill(group, auto_post=auto_post)
            summary.results.append(result)
            if result.status == "created":
                summary.bills_created += 1
            elif result.status == "error":
                summary.bills_errored += 1
            else:
                summary.bills_skipped += 1

        return summary


# ---------------------------------------------------------------------------
# URL builder for vendor bills
# ---------------------------------------------------------------------------

def build_vendor_bill_url(base_url: str, move_id: int) -> str:
    """Build the Odoo URL for a vendor bill."""
    base = base_url.rstrip("/")
    return f"{base}/odoo/accounting/vendor-bills/{move_id}"


def build_vendor_bill_url_from_example(example_url: str, move_id: int) -> str:
    """Build vendor bill URL from an example record URL pattern."""
    if not example_url:
        return ""
    # Try to replace the ID in the example URL
    # Example: https://sphe.cloudoo.ma/odoo/sales/123 -> replace 123 with move_id
    import re as _re
    # Replace last numeric segment
    pattern = r"/(\d+)(?:\?|$|#)"
    match = _re.search(pattern, example_url)
    if match:
        return example_url[:match.start(1)] + str(move_id) + example_url[match.end(1):]
    # Fallback: append
    base = example_url.rstrip("/")
    # Change path to vendor-bills
    if "/sales/" in base:
        base = base.rsplit("/sales/", 1)[0]
    return f"{base}/odoo/accounting/vendor-bills/{move_id}"
