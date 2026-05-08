import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from odoo_excel_updater import (
    check_for_update,
    compare_versions,
    parse_update_manifest,
    prepare_update_payload,
    sha256_file,
)


class UpdaterTests(unittest.TestCase):
    def test_compare_versions(self) -> None:
        self.assertGreater(compare_versions("2026.05.08.6", "2026.05.08.5"), 0)
        self.assertEqual(compare_versions("v2026.05.08.5", "2026.05.08.5"), 0)
        self.assertLess(compare_versions("2026.05.08.4", "2026.05.08.5"), 0)

    def test_parse_custom_manifest(self) -> None:
        manifest = {
            "version": "2026.05.08.6",
            "windows_zip": {
                "kind": "zip",
                "url": "https://example.com/OdooExcelAgent-Windows.zip",
                "sha256": "abc123",
            },
            "notes": "Test release",
        }

        info = parse_update_manifest(manifest, "2026.05.08.5")

        self.assertTrue(info.has_update)
        self.assertEqual(info.version, "2026.05.08.6")
        self.assertEqual(info.asset_kind, "zip")
        self.assertEqual(info.sha256, "abc123")

    def test_parse_github_release_manifest(self) -> None:
        manifest = {
            "tag_name": "v2026.05.08.6",
            "body": "Release notes",
            "html_url": "https://github.com/example/repo/releases/tag/v2026.05.08.6",
            "assets": [
                {
                    "name": "OdooExcelAgent-Windows.zip",
                    "browser_download_url": "https://github.com/example/repo/releases/download/v2026.05.08.6/OdooExcelAgent-Windows.zip",
                    "digest": "sha256:def456",
                }
            ],
        }

        info = parse_update_manifest(manifest, "2026.05.08.5")

        self.assertTrue(info.has_update)
        self.assertEqual(info.version, "2026.05.08.6")
        self.assertEqual(info.sha256, "def456")

    def test_check_for_update_from_file_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            manifest_path = Path(tmpdir) / "manifest.json"
            manifest_path.write_text(
                json.dumps(
                    {
                        "version": "2026.05.08.6",
                        "windows_zip": {"url": "https://example.com/OdooExcelAgent-Windows.zip"},
                    }
                ),
                encoding="utf-8",
            )

            info = check_for_update(manifest_path.as_uri(), "2026.05.08.5")

        self.assertTrue(info.has_update)

    def test_prepare_update_payload_from_zip(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            payload_dir = root / "payload"
            payload_dir.mkdir()
            exe = payload_dir / "OdooExcelAgent.exe"
            exe.write_bytes(b"new-exe")
            archive = root / "OdooExcelAgent-Windows.zip"
            with zipfile.ZipFile(archive, "w") as zip_file:
                zip_file.write(exe, "OdooExcelAgent-Windows/OdooExcelAgent.exe")

            prepared = prepare_update_payload(archive, root / "staging")

            self.assertEqual(prepared.new_exe.name, "OdooExcelAgent.exe")
            self.assertEqual(sha256_file(prepared.new_exe), sha256_file(exe))


if __name__ == "__main__":
    unittest.main()
