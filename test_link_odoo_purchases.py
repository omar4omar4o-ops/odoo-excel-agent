import unittest

from link_odoo_purchases import OdooPurchaseClient


class TestPickBestOrder(unittest.TestCase):
    def setUp(self) -> None:
        self.client = OdooPurchaseClient("https://sphe.cloudoo.ma", "db", "login", "key")

    def test_partner_ref_prefers_exact(self) -> None:
        orders = [
            {"id": 2, "name": "PO0002", "partner_ref": "ABC-01"},
            {"id": 1, "name": "PO0001", "partner_ref": "REF-123"},
        ]
        result = self.client._pick_best_order(
            "REF-123",
            orders,
            default_match="contains:partner_ref",
        )
        assert result is not None
        self.assertEqual(result.matched_field, "exact:partner_ref")
        self.assertEqual(result.record_id, 1)

    def test_partner_ref_contains_when_no_exact(self) -> None:
        orders = [
            {"id": 9, "name": "PO0009", "partner_ref": "REF-123-A"},
            {"id": 8, "name": "PO0008", "partner_ref": "REF-123-B"},
        ]
        result = self.client._pick_best_order(
            "REF-123",
            orders,
            default_match="contains:partner_ref",
        )
        assert result is not None
        self.assertEqual(result.matched_field, "contains:partner_ref")
        self.assertEqual(result.record_id, 9)
        self.assertEqual(result.url, "https://sphe.cloudoo.ma/odoo/purchase/9")

    def test_empty_ref_returns_empty_ref_status(self) -> None:
        result = self.client.resolve_purchase_ref("  ")
        self.assertEqual(result.status, "empty_ref")


class FakeClient(OdooPurchaseClient):
    def __init__(self) -> None:
        super().__init__("https://sphe.cloudoo.ma", "db", "login", "key")
        self.calls: list[tuple[str, str]] = []

    def search_purchase_orders_by_partner_ref(self, ref: str, operator: str) -> list[dict]:
        self.calls.append((ref, operator))
        if operator == "=":
            return []
        if operator == "ilike":
            return [{"id": 44, "name": "PO0044", "partner_ref": "CMD-44"}]
        return []


class TestResolvePurchaseRef(unittest.TestCase):
    def test_fallback_from_exact_to_contains_partner_ref(self) -> None:
        client = FakeClient()
        result = client.resolve_purchase_ref("CMD")
        self.assertEqual(client.calls, [("CMD", "="), ("CMD", "ilike")])
        self.assertEqual(result.status, "linked")
        self.assertEqual(result.source_model, "purchase.order")
        self.assertEqual(result.matched_field, "contains:partner_ref")
        self.assertEqual(result.record_id, 44)
        self.assertEqual(result.url, "https://sphe.cloudoo.ma/odoo/purchase/44")

    def test_not_found_when_partner_ref_missing(self) -> None:
        class NotFoundClient(FakeClient):
            def search_purchase_orders_by_partner_ref(self, ref: str, operator: str) -> list[dict]:
                self.calls.append((ref, operator))
                return []

        client = NotFoundClient()
        result = client.resolve_purchase_ref("NOT-THERE")
        self.assertEqual(client.calls, [("NOT-THERE", "="), ("NOT-THERE", "ilike")])
        self.assertEqual(result.status, "not_found")


if __name__ == "__main__":
    unittest.main()
