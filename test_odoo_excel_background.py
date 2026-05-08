import tempfile
import unittest
import json
from pathlib import Path
from types import SimpleNamespace

from link_odoo_vendor_bills import WORKBOOK_SLOT_ACHATS_ETRANGER
from odoo_excel_agent_support import WATCH_MODE_SELECTED_WORKBOOKS
import odoo_excel_background
from odoo_excel_background import OdooExcelAgent, load_agent_config


class BackgroundWatchFilteringTests(unittest.TestCase):
    def _make_agent(self, targets: tuple[Path, ...], report_dir: Path, backup_dir: Path) -> OdooExcelAgent:
        agent = OdooExcelAgent.__new__(OdooExcelAgent)
        agent.config = SimpleNamespace(
            processing=SimpleNamespace(
                watch_mode=WATCH_MODE_SELECTED_WORKBOOKS,
                watch_file=None,
                achats_local_file=targets[0] if len(targets) > 0 else None,
                achats_etranger_file=targets[1] if len(targets) > 1 else None,
                seller_previous_file=targets[2] if len(targets) > 2 else None,
                watch_targets=targets,
                report_dir=report_dir,
                backup_dir=backup_dir,
            )
        )
        return agent

    def test_should_ignore_unselected_workbooks_in_selected_mode(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            targets = (
                root / "EXCEL FACTURE ACHATS LOCAL.xlsx",
                root / "TRACKING ACHATS ETRANGER (1).xlsx",
                root / "L'ETAT DES COMMANDES.xlsx",
            )
            report_dir = root / "reports"
            backup_dir = root / "backups"
            agent = self._make_agent(targets, report_dir, backup_dir)

            self.assertFalse(agent._should_ignore(targets[0]))
            self.assertFalse(agent._should_ignore(targets[1]))
            self.assertFalse(agent._should_ignore(targets[2]))
            self.assertTrue(agent._should_ignore(root / "Other Workbook.xlsx"))
            self.assertTrue(agent._should_ignore(root / "~$TEMP.xlsx"))
            self.assertTrue(agent._should_ignore(report_dir / "report.xlsx"))
            self.assertTrue(agent._should_ignore(backup_dir / "backup.xlsx"))

    def test_workbook_slot_for_path_detects_configured_achats_etranger(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            targets = (
                root / "Local Any Name.xlsx",
                root / "Renamed ACHATS ETRANGER.xlsx",
                root / "Seller.xlsx",
            )
            agent = self._make_agent(targets, root / "reports", root / "backups")

            self.assertEqual(agent._workbook_slot_for_path(targets[1]), WORKBOOK_SLOT_ACHATS_ETRANGER)

    def test_load_agent_config_rejects_url_in_database_field(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            config_path = root / "config.json"
            payload = {
                "version": 4,
                "odoo": {
                    "url": "https://sphe.cloudoo.ma",
                    "db": "https://limewire.com/d/9nluz#2VJ4biiLRw",
                    "login": "user@example.com",
                    "api_key": "legacy-api-key",
                    "credential_target": "",
                    "record_url_example": "",
                },
                "manual": {"last_file": ""},
                "background": {
                    "watch_mode": WATCH_MODE_SELECTED_WORKBOOKS,
                    "achats_local_file": "",
                    "achats_etranger_file": "",
                    "seller_previous_file": "",
                    "watch_file": "",
                    "watch_folder": str(root),
                    "recursive": False,
                    "process_existing_on_start": False,
                    "update_open_workbook": True,
                    "excel_event_monitoring": True,
                    "excel_session_backend": "pywin32",
                    "excel_save_debounce_seconds": 1,
                    "allow_live_update_with_autosave": False,
                    "visible_excel": False,
                    "write_report_file": False,
                    "stable_backup_name": True,
                    "settle_seconds": 3,
                    "retry_delay_seconds": 15,
                },
                "paths": {
                    "backup_dir": str(root / "backups"),
                    "report_dir": str(root / "reports"),
                    "state_file": str(root / "state.json"),
                    "log_file": str(root / "agent.log"),
                    "runtime_status_file": str(root / "runtime_status.json"),
                },
            }
            config_path.write_text(json.dumps(payload), encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "database must be the database name only"):
                load_agent_config(config_path)

    def test_background_main_returns_cleanly_when_api_key_is_missing(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            config_path = root / "config.json"
            payload = {
                "version": 4,
                "odoo": {
                    "url": "https://sphe.cloudoo.ma",
                    "db": "sphe.cloudoo.ma",
                    "login": "user@example.com",
                    "credential_target": "",
                    "record_url_example": "",
                },
                "background": {
                    "watch_mode": WATCH_MODE_SELECTED_WORKBOOKS,
                    "achats_local_file": "",
                    "achats_etranger_file": "",
                    "seller_previous_file": "",
                },
                "paths": {
                    "backup_dir": str(root / "backups"),
                    "report_dir": str(root / "reports"),
                    "state_file": str(root / "state.json"),
                    "log_file": str(root / "agent.log"),
                    "runtime_status_file": str(root / "runtime_status.json"),
                },
            }
            config_path.write_text(json.dumps(payload), encoding="utf-8")
            original_open_setup_ui = odoo_excel_background.open_setup_ui
            opened: list[Path] = []
            odoo_excel_background.open_setup_ui = lambda path: opened.append(path)
            try:
                exit_code = odoo_excel_background.main(["--config", str(config_path)])
            finally:
                odoo_excel_background.open_setup_ui = original_open_setup_ui

            status = json.loads((root / "runtime_status.json").read_text(encoding="utf-8"))

        self.assertEqual(exit_code, 2)
        self.assertEqual(status["state"], "startup_failed")
        self.assertEqual(status["last_issue_code"], "missing_api_key")
        self.assertEqual(opened, [config_path.resolve()])


if __name__ == "__main__":
    unittest.main()
