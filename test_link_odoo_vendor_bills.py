import unittest
import xmlrpc.client
from pathlib import Path

try:
    from openpyxl import Workbook
except ImportError:  # pragma: no cover
    Workbook = None

from link_odoo_vendor_bills import (
    ALL_HEADER_VARIANTS,
    ETRANGER_TOTAL_HEADER_VARIANTS,
    LOCAL_COMMAND_HEADER_VARIANTS,
    LOCAL_HEADER_VARIANTS,
    LOOKUP_MODE_COMMAND_REF,
    LOOKUP_MODE_PARTNER_REF,
    LOOKUP_MODE_TOTAL_AMOUNT,
    OdooClient,
    PurchaseLinkResult,
    WORKBOOK_SLOT_ACHATS_ETRANGER,
    WorkbookOrderCell,
    apply_links_to_workbook,
    explain_odoo_exception,
    scan_workbook_orders_from_file,
    select_cells_for_results,
    validate_odoo_settings,
    write_links_with_openpyxl,
    workbook_rule_for_path,
    workbook_rule_for_slot,
)


def make_order(
    order_id: int,
    *,
    name: str,
    partner_ref: str,
    amount_total: float,
) -> dict:
    return {
        "id": order_id,
        "name": name,
        "partner_ref": partner_ref,
        "partner_id": [11, "Vendor X"],
        "state": "purchase",
        "amount_total": amount_total,
    }


class RuleSelectionTests(unittest.TestCase):
    def test_local_workbook_uses_name_lookup_with_row_fallback_headers(self) -> None:
        rule = workbook_rule_for_path(Path(r"C:\tmp\EXCEL FACTURE ACHATS LOCAL.xlsx"))
        self.assertEqual(rule.lookup_mode, LOOKUP_MODE_COMMAND_REF)
        self.assertEqual(rule.headers, LOCAL_HEADER_VARIANTS | LOCAL_COMMAND_HEADER_VARIANTS)
        self.assertTrue(rule.row_fallback_on_not_found)

    def test_etranger_workbook_uses_total_lookup(self) -> None:
        rule = workbook_rule_for_path(Path(r"C:\tmp\TRACKING ACHATS ETRANGER (1).xlsx"))
        self.assertEqual(rule.lookup_mode, LOOKUP_MODE_TOTAL_AMOUNT)
        self.assertEqual(rule.headers, ETRANGER_TOTAL_HEADER_VARIANTS)

    def test_etranger_slot_forces_total_lookup_for_renamed_workbook(self) -> None:
        rule = workbook_rule_for_slot(WORKBOOK_SLOT_ACHATS_ETRANGER, Path(r"C:\tmp\Renamed Copy.xlsx"))
        self.assertEqual(rule.lookup_mode, LOOKUP_MODE_TOTAL_AMOUNT)
        self.assertEqual(rule.headers, ETRANGER_TOTAL_HEADER_VARIANTS)
        self.assertFalse(rule.row_fallback_on_not_found)

    def test_other_workbooks_keep_legacy_lookup(self) -> None:
        rule = workbook_rule_for_path(Path(r"C:\tmp\Anything Else.xlsx"))
        self.assertEqual(rule.lookup_mode, LOOKUP_MODE_PARTNER_REF)
        self.assertEqual(rule.headers, ALL_HEADER_VARIANTS)


class CommandRefClient(OdooClient):
    def __init__(self) -> None:
        super().__init__("https://sphe.cloudoo.ma", "db", "login", "key")
        self.name_calls: list[tuple[str, str]] = []
        self.partner_ref_calls: list[tuple[str, str]] = []

    def search_purchase_orders_by_name(self, ref: str, operator: str) -> list[dict]:
        self.name_calls.append((ref, operator))
        if operator == "=":
            return []
        return [make_order(66, name="PO00066", partner_ref="FA356409", amount_total=120.0)]

    def search_purchase_orders_by_partner_ref(self, ref: str, operator: str) -> list[dict]:
        self.partner_ref_calls.append((ref, operator))
        return []


class CommandRefPartnerFallbackClient(OdooClient):
    def __init__(self) -> None:
        super().__init__("https://sphe.cloudoo.ma", "db", "login", "key")
        self.name_calls: list[tuple[str, str]] = []
        self.partner_ref_calls: list[tuple[str, str]] = []

    def search_purchase_orders_by_name(self, ref: str, operator: str) -> list[dict]:
        self.name_calls.append((ref, operator))
        return []

    def search_purchase_orders_by_partner_ref(self, ref: str, operator: str) -> list[dict]:
        self.partner_ref_calls.append((ref, operator))
        if operator == "=":
            return []
        return [make_order(77, name="PO00077", partner_ref="FA202603527", amount_total=460.01)]


class TotalClient(OdooClient):
    def __init__(self) -> None:
        super().__init__("https://sphe.cloudoo.ma", "db", "login", "key")
        self.exact_calls: list[float] = []
        self.range_calls: list[tuple[float, float]] = []

    def search_purchase_orders_by_amount_exact(self, amount: float) -> list[dict]:
        self.exact_calls.append(amount)
        return []

    def search_purchase_orders_by_amount_range(self, minimum: float, maximum: float) -> list[dict]:
        self.range_calls.append((minimum, maximum))
        return [
            make_order(310, name="PO00310", partner_ref="REF-A", amount_total=8753.76),
            make_order(309, name="PO00309", partner_ref="REF-B", amount_total=8753.76),
        ]


class TotalToleranceFallbackClient(OdooClient):
    def __init__(self) -> None:
        super().__init__("https://sphe.cloudoo.ma", "db", "login", "key")
        self.exact_calls: list[float] = []
        self.range_calls: list[tuple[float, float]] = []

    def search_purchase_orders_by_amount_exact(self, amount: float) -> list[dict]:
        self.exact_calls.append(amount)
        return []

    def search_purchase_orders_by_amount_range(self, minimum: float, maximum: float) -> list[dict]:
        self.range_calls.append((minimum, maximum))
        if round(maximum - minimum, 2) <= 0.02:
            return []
        return [make_order(401, name="PO00401", partner_ref="REF-401", amount_total=100.04)]


class ResolveModeTests(unittest.TestCase):
    def test_command_ref_lookup_uses_name(self) -> None:
        client = CommandRefClient()
        result = client.resolve_purchase_ref("PO000", lookup_mode=LOOKUP_MODE_COMMAND_REF)
        self.assertEqual(client.name_calls, [("PO000", "="), ("PO000", "ilike")])
        self.assertEqual(client.partner_ref_calls, [("PO000", "=")])
        self.assertEqual(result.status, "linked")
        self.assertEqual(result.matched_field, "contains:name")
        self.assertEqual(result.record_id, 66)
        self.assertEqual(result.url, "https://sphe.cloudoo.ma/odoo/purchase/66")
        self.assertEqual(result.ref_value, "PO00066")

    def test_command_ref_lookup_falls_back_to_partner_ref(self) -> None:
        client = CommandRefPartnerFallbackClient()
        result = client.resolve_purchase_ref("FA202603527", lookup_mode=LOOKUP_MODE_COMMAND_REF)
        self.assertEqual(result.status, "linked")
        self.assertEqual(result.matched_field, "exact:partner_ref")
        self.assertEqual(result.record_id, 77)
        self.assertEqual(result.url, "https://sphe.cloudoo.ma/odoo/purchase/77")
        self.assertEqual(client.name_calls, [("FA202603527", "="), ("FA202603527", "ilike")])
        self.assertEqual(
            client.partner_ref_calls,
            [("FA202603527", "="), ("FA202603527", "ilike")],
        )

    def test_total_lookup_uses_range_and_selects_latest(self) -> None:
        client = TotalClient()
        result = client.resolve_purchase_ref("8753.76", lookup_mode=LOOKUP_MODE_TOTAL_AMOUNT)
        self.assertEqual(result.status, "linked")
        self.assertEqual(result.matched_field, "range:amount_total+-0.01")
        self.assertEqual(result.record_id, 310)
        self.assertEqual(result.url, "https://sphe.cloudoo.ma/odoo/purchase/310")
        self.assertEqual(len(client.exact_calls), 1)
        self.assertEqual(len(client.range_calls), 1)
        self.assertAlmostEqual(client.exact_calls[0], 8753.76, places=2)
        self.assertAlmostEqual(client.range_calls[0][0], 8753.75, places=2)
        self.assertAlmostEqual(client.range_calls[0][1], 8753.77, places=2)

    def test_total_lookup_tries_wider_tolerance_when_needed(self) -> None:
        client = TotalToleranceFallbackClient()
        result = client.resolve_purchase_ref("100.00", lookup_mode=LOOKUP_MODE_TOTAL_AMOUNT)
        self.assertEqual(result.status, "linked")
        self.assertEqual(result.matched_field, "range:amount_total+-0.05")
        self.assertEqual(result.record_id, 401)
        self.assertEqual(len(client.exact_calls), 1)
        self.assertGreaterEqual(len(client.range_calls), 2)

    def test_total_lookup_invalid_amount(self) -> None:
        client = TotalClient()
        result = client.resolve_purchase_ref("NOT-A-NUMBER", lookup_mode=LOOKUP_MODE_TOTAL_AMOUNT)
        self.assertEqual(result.status, "invalid_amount")
        self.assertEqual(client.exact_calls, [])
        self.assertEqual(client.range_calls, [])


class OdooSettingsValidationTests(unittest.TestCase):
    def test_validate_odoo_settings_rejects_url_in_db_field(self) -> None:
        with self.assertRaisesRegex(ValueError, "database must be the database name only"):
            validate_odoo_settings(
                "https://sphe.cloudoo.ma",
                "https://limewire.com/d/9nluz#2VJ4biiLRw",
                "user@example.com",
            )

    def test_validate_odoo_settings_rejects_mismatched_record_url_host(self) -> None:
        with self.assertRaisesRegex(ValueError, "same host"):
            validate_odoo_settings(
                "https://sphe.cloudoo.ma",
                "sphe.cloudoo.ma",
                "user@example.com",
                "https://other.example.com/odoo/purchase/1",
            )

    def test_explain_odoo_exception_for_url_like_database(self) -> None:
        exc = xmlrpc.client.Fault(1, 'database "https://limewire.com/d/9nluz#2VJ4biiLRw" does not exist')
        message = str(
            explain_odoo_exception(
                exc,
                odoo_url="https://sphe.cloudoo.ma",
                odoo_db="https://limewire.com/d/9nluz#2VJ4biiLRw",
            )
        )
        self.assertIn("database field contains a URL", message)

    def test_explain_odoo_exception_for_connection_slots(self) -> None:
        exc = xmlrpc.client.Fault(
            1,
            "psycopg2.OperationalError: remaining connection slots are reserved for roles with the SUPERUSER attribute",
        )
        message = str(explain_odoo_exception(exc, odoo_url="https://sphe.cloudoo.ma", odoo_db="sphe.cloudoo.ma"))
        self.assertIn("temporarily overloaded", message)


class FakeApplication:
    def __init__(self) -> None:
        self.ScreenUpdating = True
        self.Calculation = -4105
        self.EnableEvents = True


class FakeCellHyperlinks:
    def __init__(self) -> None:
        self.deleted = False

    def Delete(self) -> None:
        self.deleted = True


class FakeCell:
    def __init__(self) -> None:
        self.Hyperlinks = FakeCellHyperlinks()


class FakeSheetHyperlinks:
    def __init__(self) -> None:
        self.added: list[dict] = []

    def Add(self, **kwargs) -> None:
        self.added.append(kwargs)


class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name
        self.Hyperlinks = FakeSheetHyperlinks()
        self._cells: dict[tuple[int, int], FakeCell] = {}

    def Cells(self, row: int, column: int) -> FakeCell:
        key = (row, column)
        if key not in self._cells:
            self._cells[key] = FakeCell()
        return self._cells[key]


class FakeWorkbook:
    def __init__(self, sheet: FakeSheet) -> None:
        self.Application = FakeApplication()
        self.Worksheets = [sheet]


class HyperlinkWriteTests(unittest.TestCase):
    def test_apply_links_keeps_original_cell_text(self) -> None:
        sheet = FakeSheet("Feuil1")
        workbook = FakeWorkbook(sheet)
        cells = [WorkbookOrderCell(sheet="Feuil1", row=9, column=3, address="C9", order_name="FA356409")]
        results = {
            "FA356409": PurchaseLinkResult(
                status="linked",
                source_model="purchase.order",
                record_id=66,
                url="https://sphe.cloudoo.ma/odoo/purchase/66",
            )
        }
        linked = apply_links_to_workbook(workbook, cells, results)
        self.assertEqual(linked, 1)
        self.assertEqual(len(sheet.Hyperlinks.added), 1)
        self.assertEqual(sheet.Hyperlinks.added[0]["TextToDisplay"], "FA356409")
        self.assertEqual(sheet.Hyperlinks.added[0]["Address"], "https://sphe.cloudoo.ma/odoo/purchase/66")

    @unittest.skipIf(Workbook is None, "openpyxl is not installed")
    def test_openpyxl_writer_keeps_original_cell_text(self) -> None:
        import tempfile
        from openpyxl import load_workbook

        with tempfile.TemporaryDirectory() as tmpdir:
            workbook_path = Path(tmpdir) / "book.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Feuil1"
            sheet["C9"] = "FA356409"
            workbook.save(workbook_path)

            cells = [WorkbookOrderCell(sheet="Feuil1", row=9, column=3, address="C9", order_name="FA356409")]
            results = {
                "FA356409": PurchaseLinkResult(
                    status="linked",
                    source_model="purchase.order",
                    record_id=66,
                    url="https://sphe.cloudoo.ma/odoo/purchase/66",
                )
            }
            linked = write_links_with_openpyxl(workbook_path, cells, results)

            updated = load_workbook(workbook_path, read_only=False)
            try:
                updated_cell = updated["Feuil1"]["C9"]
                self.assertEqual(linked, 1)
                self.assertEqual(updated_cell.value, "FA356409")
                self.assertEqual(updated_cell.hyperlink.target, "https://sphe.cloudoo.ma/odoo/purchase/66")
            finally:
                updated.close()


@unittest.skipIf(Workbook is None, "openpyxl is not installed")
class LocalFallbackScanTests(unittest.TestCase):
    def test_etranger_slot_scan_ignores_n_commande_and_reads_only_mtt_de_facture(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory() as tmpdir:
            workbook_path = Path(tmpdir) / "Renamed ACHATS ETRANGER.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Feuil1"
            sheet["C1"] = "N COMMANDE"
            sheet["H1"] = "MTT DE FACTURE"
            sheet["C2"] = "SA2026000002942"
            sheet["H2"] = 3542.87
            workbook.save(workbook_path)

            rule = workbook_rule_for_slot(WORKBOOK_SLOT_ACHATS_ETRANGER, workbook_path)
            scan_result = scan_workbook_orders_from_file(workbook_path, visible_excel=False, workbook_rule=rule)

        self.assertEqual(scan_result.issue_code, "")
        self.assertEqual(len(scan_result.cells), 1)
        self.assertEqual(scan_result.cells[0].order_name, "3542.87")
        self.assertEqual(scan_result.cells[0].address, "H2")
        self.assertEqual(scan_result.cells[0].header_name, "mtt de facture")

    def test_local_scan_collects_both_headers_and_uses_secondary_when_primary_not_found(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory() as tmpdir:
            workbook_path = Path(tmpdir) / "EXCEL FACTURE ACHATS LOCAL.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Feuil1"
            sheet["C1"] = "N°FACTURE"
            sheet["D1"] = "N commandes"
            sheet["C2"] = "FA-NOT-FOUND"
            sheet["D2"] = "FA202603527"
            workbook.save(workbook_path)

            rule = workbook_rule_for_path(workbook_path)
            scan_result = scan_workbook_orders_from_file(workbook_path, visible_excel=False, workbook_rule=rule)

        self.assertEqual(len(scan_result.cells), 2)
        self.assertEqual([cell.order_name for cell in scan_result.cells], ["FA-NOT-FOUND", "FA202603527"])
        results = {
            "FA-NOT-FOUND": PurchaseLinkResult(status="not_found"),
            "FA202603527": PurchaseLinkResult(status="linked", url="https://sphe.cloudoo.ma/odoo/purchase/77"),
        }
        selected = select_cells_for_results(scan_result.cells, results, rule)
        self.assertEqual(len(selected), 1)
        self.assertEqual(selected[0].order_name, "FA202603527")
        self.assertTrue(selected[0].fallback_used)
        self.assertEqual(selected[0].fallback_from, "FA-NOT-FOUND")

    def test_missing_required_header_is_reported(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory() as tmpdir:
            workbook_path = Path(tmpdir) / "EXCEL FACTURE ACHATS LOCAL.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet["A1"] = "Something else"
            workbook.save(workbook_path)

            rule = workbook_rule_for_path(workbook_path)
            scan_result = scan_workbook_orders_from_file(workbook_path, visible_excel=False, workbook_rule=rule)

        self.assertEqual(scan_result.cells, [])
        self.assertEqual(scan_result.issue_code, "missing_required_header")


if __name__ == "__main__":
    unittest.main()

