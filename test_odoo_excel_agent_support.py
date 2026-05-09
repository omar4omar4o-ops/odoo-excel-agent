import json
import tempfile
import unittest
from pathlib import Path

from link_odoo_vendor_bills import (
    ETRANGER_WORKBOOK_FILE_NAME,
    LOCAL_WORKBOOK_FILE_NAME,
    PERFORMANCE_MODE_LIVE,
    PERFORMANCE_MODE_SILENT,
)
from odoo_excel_agent_support import (
    DEFAULT_UPDATE_URL,
    WATCH_MODE_SELECTED_WORKBOOKS,
    WATCH_MODE_FILE,
    default_config,
    get_background_watch_targets,
    load_normalized_config,
)


class OdooExcelAgentSupportTests(unittest.TestCase):
    def test_default_config_uses_selected_workbooks_mode(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            config = default_config(Path(tmpdir))
        self.assertEqual(config["background"]["watch_mode"], WATCH_MODE_SELECTED_WORKBOOKS)
        self.assertEqual(config["background"]["achats_local_file"], "")
        self.assertEqual(config["background"]["achats_etranger_file"], "")
        self.assertEqual(config["background"]["seller_previous_file"], "")
        self.assertEqual(config["background"]["performance_mode"], PERFORMANCE_MODE_SILENT)
        self.assertFalse(config["background"]["update_open_workbook"])
        self.assertFalse(config["background"]["excel_event_monitoring"])
        self.assertFalse(config["background"]["allow_live_update_with_autosave"])
        self.assertEqual(config["background"]["excel_session_backend"], "pywin32")
        self.assertEqual(config["updates"]["manifest_url"], DEFAULT_UPDATE_URL)

    def test_silent_mode_forces_live_options_off_on_load(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_path = Path(tmpdir)
            raw = {
                "background": {
                    "performance_mode": PERFORMANCE_MODE_SILENT,
                    "update_open_workbook": True,
                    "excel_event_monitoring": True,
                    "allow_live_update_with_autosave": True,
                    "excel_session_backend": "xlwings",
                }
            }
            config_path = temp_path / "config.json"
            config_path.write_text(json.dumps(raw), encoding="utf-8")

            normalized, messages = load_normalized_config(config_path)

        background = normalized["background"]
        self.assertEqual(background["performance_mode"], PERFORMANCE_MODE_SILENT)
        self.assertFalse(background["update_open_workbook"])
        self.assertFalse(background["excel_event_monitoring"])
        self.assertFalse(background["allow_live_update_with_autosave"])
        self.assertEqual(background["excel_session_backend"], "pywin32")
        self.assertTrue(any("xlwings" in message for message in messages))

    def test_live_mode_preserves_live_options_on_load(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_path = Path(tmpdir)
            raw = {
                "background": {
                    "performance_mode": PERFORMANCE_MODE_LIVE,
                    "update_open_workbook": True,
                    "excel_event_monitoring": True,
                    "allow_live_update_with_autosave": True,
                }
            }
            config_path = temp_path / "config.json"
            config_path.write_text(json.dumps(raw), encoding="utf-8")

            normalized, _ = load_normalized_config(config_path)

        background = normalized["background"]
        self.assertEqual(background["performance_mode"], PERFORMANCE_MODE_LIVE)
        self.assertTrue(background["update_open_workbook"])
        self.assertTrue(background["excel_event_monitoring"])
        self.assertTrue(background["allow_live_update_with_autosave"])
        self.assertEqual(background["excel_session_backend"], "pywin32")

    def test_legacy_xlwings_backend_never_survives_normalization(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_path = Path(tmpdir)
            raw = {
                "background": {
                    "performance_mode": PERFORMANCE_MODE_LIVE,
                    "update_open_workbook": True,
                    "excel_event_monitoring": True,
                    "allow_live_update_with_autosave": True,
                    "excel_session_backend": "xlwings",
                }
            }
            config_path = temp_path / "config.json"
            config_path.write_text(json.dumps(raw), encoding="utf-8")

            normalized, messages = load_normalized_config(config_path)

        background = normalized["background"]
        self.assertEqual(background["performance_mode"], PERFORMANCE_MODE_LIVE)
        self.assertTrue(background["update_open_workbook"])
        self.assertTrue(background["excel_event_monitoring"])
        self.assertTrue(background["allow_live_update_with_autosave"])
        self.assertEqual(background["excel_session_backend"], "pywin32")
        self.assertTrue(any("xlwings" in message for message in messages))

    def test_load_normalized_config_migrates_legacy_local_watch_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_path = Path(tmpdir)
            watch_file = temp_path / LOCAL_WORKBOOK_FILE_NAME
            raw = {
                "background": {
                    "watch_mode": WATCH_MODE_FILE,
                    "watch_file": str(watch_file),
                }
            }
            config_path = temp_path / "config.json"
            config_path.write_text(json.dumps(raw), encoding="utf-8")

            normalized, messages = load_normalized_config(config_path)

        background = normalized["background"]
        self.assertEqual(background["watch_mode"], WATCH_MODE_SELECTED_WORKBOOKS)
        self.assertEqual(background["achats_local_file"], str(watch_file.resolve()))
        self.assertEqual(background["achats_etranger_file"], "")
        self.assertEqual(background["seller_previous_file"], "")
        self.assertTrue(any("ACHATS LOCAL" in message for message in messages))

    def test_load_normalized_config_recovers_from_corrupt_json(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "config.json"
            config_path.write_text("{not valid json", encoding="utf-8")

            normalized, messages = load_normalized_config(config_path)

        self.assertEqual(normalized["background"]["watch_mode"], WATCH_MODE_SELECTED_WORKBOOKS)
        self.assertTrue(any("unreadable" in message for message in messages))

    def test_load_normalized_config_keeps_unknown_legacy_file_mode(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_path = Path(tmpdir)
            watch_file = temp_path / "custom.xlsx"
            raw = {
                "background": {
                    "watch_mode": WATCH_MODE_FILE,
                    "watch_file": str(watch_file),
                }
            }
            config_path = temp_path / "config.json"
            config_path.write_text(json.dumps(raw), encoding="utf-8")

            normalized, messages = load_normalized_config(config_path)

        background = normalized["background"]
        self.assertEqual(background["watch_mode"], WATCH_MODE_FILE)
        self.assertEqual(background["watch_file"], str(watch_file.resolve()))
        self.assertEqual(background["achats_local_file"], "")
        self.assertEqual(background["achats_etranger_file"], "")
        self.assertEqual(messages, [])

    def test_get_background_watch_targets_returns_configured_selected_workbooks(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_path = Path(tmpdir)
            local_path = temp_path / LOCAL_WORKBOOK_FILE_NAME
            etranger_path = temp_path / ETRANGER_WORKBOOK_FILE_NAME
            seller_path = temp_path / "L'ETAT DES COMMANDES.xlsx"
            config = default_config(temp_path)
            config["background"]["achats_local_file"] = str(local_path)
            config["background"]["achats_etranger_file"] = str(etranger_path)
            config["background"]["seller_previous_file"] = str(seller_path)

            targets = get_background_watch_targets(config)

        self.assertEqual(targets, [local_path.resolve(), etranger_path.resolve(), seller_path.resolve()])


if __name__ == "__main__":
    unittest.main()
