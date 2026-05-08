"""Self-update helpers for the packaged Odoo Excel Agent."""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


USER_AGENT = "OdooExcelAgent-Updater/1.0"
UPDATE_SCRIPT_NAME = "install-odoo-excel-agent-update.ps1"


@dataclass(frozen=True)
class UpdateInfo:
    current_version: str
    version: str
    has_update: bool
    download_url: str
    sha256: str = ""
    asset_kind: str = "zip"
    notes: str = ""
    release_url: str = ""


@dataclass(frozen=True)
class PreparedUpdate:
    payload_dir: Path
    new_exe: Path


def normalize_version(value: Any) -> str:
    return str(value or "").strip().lstrip("vV")


def _version_parts(version: str) -> list[Any]:
    parts: list[Any] = []
    for token in re.split(r"[^0-9A-Za-z]+", normalize_version(version)):
        if not token:
            continue
        if token.isdigit():
            parts.append(int(token))
        else:
            parts.append(token.casefold())
    return parts


def compare_versions(left: str, right: str) -> int:
    left_parts = _version_parts(left)
    right_parts = _version_parts(right)
    max_len = max(len(left_parts), len(right_parts))
    for index in range(max_len):
        left_item = left_parts[index] if index < len(left_parts) else 0
        right_item = right_parts[index] if index < len(right_parts) else 0
        if left_item == right_item:
            continue
        if isinstance(left_item, int) and isinstance(right_item, str):
            return 1
        if isinstance(left_item, str) and isinstance(right_item, int):
            return -1
        return 1 if left_item > right_item else -1
    return 0


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fetch_update_manifest(manifest_url: str) -> dict[str, Any]:
    value = str(manifest_url or "").strip()
    if not value:
        raise ValueError("Update manifest URL is empty.")
    parsed = urllib.parse.urlparse(value)
    if parsed.scheme not in {"http", "https", "file"}:
        raise ValueError("Update manifest URL must be http(s) or file://.")
    request = urllib.request.Request(value, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    with urllib.request.urlopen(request, timeout=20) as response:
        payload = response.read()
    loaded = json.loads(payload.decode("utf-8-sig"))
    if not isinstance(loaded, dict):
        raise ValueError("Update manifest must be a JSON object.")
    return loaded


def parse_update_manifest(manifest: dict[str, Any], current_version: str) -> UpdateInfo:
    if "assets" in manifest and "tag_name" in manifest:
        return _parse_github_release_manifest(manifest, current_version)
    version = normalize_version(manifest.get("version") or manifest.get("tag_name"))
    if not version:
        raise ValueError("Update manifest is missing version.")
    asset = _pick_manifest_asset(manifest)
    download_url = str(asset.get("url") or asset.get("download_url") or "").strip()
    if not download_url:
        raise ValueError("Update manifest is missing a download URL.")
    sha256 = str(asset.get("sha256") or asset.get("digest") or "").strip()
    if sha256.casefold().startswith("sha256:"):
        sha256 = sha256.split(":", 1)[1].strip()
    kind = str(asset.get("kind") or _kind_from_url(download_url)).strip().casefold()
    return UpdateInfo(
        current_version=normalize_version(current_version),
        version=version,
        has_update=compare_versions(version, current_version) > 0,
        download_url=download_url,
        sha256=sha256,
        asset_kind=kind,
        notes=str(manifest.get("notes") or manifest.get("body") or ""),
        release_url=str(manifest.get("release_url") or manifest.get("html_url") or ""),
    )


def check_for_update(manifest_url: str, current_version: str) -> UpdateInfo:
    return parse_update_manifest(fetch_update_manifest(manifest_url), current_version)


def _pick_manifest_asset(manifest: dict[str, Any]) -> dict[str, Any]:
    for key in ("windows_zip", "windows", "asset", "windows_exe"):
        value = manifest.get(key)
        if isinstance(value, dict):
            return value
    assets = manifest.get("assets")
    if isinstance(assets, list):
        return _pick_asset_from_list(assets)
    raise ValueError("Update manifest is missing a Windows asset.")


def _parse_github_release_manifest(manifest: dict[str, Any], current_version: str) -> UpdateInfo:
    version = normalize_version(manifest.get("tag_name") or manifest.get("name"))
    if not version:
        raise ValueError("GitHub release response is missing tag_name.")
    assets = manifest.get("assets")
    if not isinstance(assets, list):
        raise ValueError("GitHub release response is missing assets.")
    asset = _pick_asset_from_list(assets)
    download_url = str(asset.get("browser_download_url") or asset.get("url") or "").strip()
    if not download_url:
        raise ValueError("Selected GitHub release asset is missing browser_download_url.")
    digest = str(asset.get("digest") or "").strip()
    sha256 = digest.split(":", 1)[1].strip() if digest.casefold().startswith("sha256:") else ""
    return UpdateInfo(
        current_version=normalize_version(current_version),
        version=version,
        has_update=compare_versions(version, current_version) > 0,
        download_url=download_url,
        sha256=sha256,
        asset_kind=_kind_from_url(download_url),
        notes=str(manifest.get("body") or ""),
        release_url=str(manifest.get("html_url") or ""),
    )


def _pick_asset_from_list(assets: list[Any]) -> dict[str, Any]:
    candidates = [asset for asset in assets if isinstance(asset, dict)]
    if not candidates:
        raise ValueError("No usable update assets were found.")

    def score(asset: dict[str, Any]) -> tuple[int, int]:
        name = str(asset.get("name") or asset.get("url") or asset.get("browser_download_url") or "").casefold()
        if "odoexcelagent-windows.zip" in name or "odooexcelagent-windows.zip" in name:
            return (0, 0)
        if "windows" in name and name.endswith(".zip"):
            return (1, 0)
        if name.endswith(".exe"):
            return (2, 0)
        if name.endswith(".zip"):
            return (3, 0)
        return (9, 0)

    selected = sorted(candidates, key=score)[0]
    if score(selected)[0] >= 9:
        raise ValueError("No .zip or .exe Windows update asset was found.")
    return selected


def _kind_from_url(url: str) -> str:
    path = urllib.parse.urlparse(url).path.casefold()
    if path.endswith(".exe"):
        return "exe"
    return "zip"


def download_update_asset(info: UpdateInfo, download_dir: Path | None = None) -> Path:
    if not info.download_url:
        raise ValueError("No update download URL is available.")
    download_root = download_dir or Path(tempfile.mkdtemp(prefix="OdooExcelAgentUpdate-"))
    download_root.mkdir(parents=True, exist_ok=True)
    parsed = urllib.parse.urlparse(info.download_url)
    filename = Path(urllib.parse.unquote(parsed.path)).name or f"OdooExcelAgent-update.{info.asset_kind}"
    destination = download_root / filename
    request = urllib.request.Request(info.download_url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=120) as response, destination.open("wb") as output:
        shutil.copyfileobj(response, output)
    if info.sha256:
        actual = sha256_file(destination)
        if actual.casefold() != info.sha256.casefold():
            destination.unlink(missing_ok=True)
            raise RuntimeError("Downloaded update failed SHA-256 verification.")
    return destination


def prepare_update_payload(asset_path: Path, staging_dir: Path | None = None) -> PreparedUpdate:
    staging_root = staging_dir or Path(tempfile.mkdtemp(prefix="OdooExcelAgentUpdatePayload-"))
    staging_root.mkdir(parents=True, exist_ok=True)
    suffix = asset_path.suffix.casefold()
    if suffix == ".zip":
        with zipfile.ZipFile(asset_path) as archive:
            archive.extractall(staging_root)
        exe_candidates = sorted(staging_root.rglob("OdooExcelAgent.exe"))
        if not exe_candidates:
            raise RuntimeError("Update ZIP does not contain OdooExcelAgent.exe.")
        new_exe = exe_candidates[0]
        return PreparedUpdate(payload_dir=new_exe.parent, new_exe=new_exe)
    if suffix == ".exe":
        new_exe = staging_root / "OdooExcelAgent.exe"
        shutil.copy2(asset_path, new_exe)
        return PreparedUpdate(payload_dir=staging_root, new_exe=new_exe)
    raise RuntimeError("Update asset must be .zip or .exe.")


def schedule_update_install(
    *,
    current_exe: Path,
    payload_dir: Path,
    install_dir: Path,
    config_path: Path,
    restart: bool = True,
) -> Path:
    script_path = Path(tempfile.mkdtemp(prefix="OdooExcelAgentUpdateScript-")) / UPDATE_SCRIPT_NAME
    script_path.write_text(_powershell_update_script(), encoding="utf-8")
    args = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        str(script_path),
        "-CurrentExe",
        str(current_exe),
        "-PayloadDir",
        str(payload_dir),
        "-InstallDir",
        str(install_dir),
        "-ConfigPath",
        str(config_path),
        "-OldPid",
        str(os.getpid()),
        "-Restart",
        "1" if restart else "0",
    ]
    subprocess.Popen(args, cwd=str(install_dir), close_fds=True)
    return script_path


def _powershell_update_script() -> str:
    return r'''param(
    [Parameter(Mandatory=$true)][string]$CurrentExe,
    [Parameter(Mandatory=$true)][string]$PayloadDir,
    [Parameter(Mandatory=$true)][string]$InstallDir,
    [Parameter(Mandatory=$true)][string]$ConfigPath,
    [Parameter(Mandatory=$true)][int]$OldPid,
    [string]$Restart = "1"
)

$ErrorActionPreference = "Stop"
Start-Sleep -Milliseconds 800
while (Get-Process -Id $OldPid -ErrorAction SilentlyContinue) {
    Start-Sleep -Milliseconds 500
}

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
$backupPath = Join-Path $InstallDir ("OdooExcelAgent.exe.bak-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
if (Test-Path -LiteralPath $CurrentExe) {
    Copy-Item -LiteralPath $CurrentExe -Destination $backupPath -Force
}

Copy-Item -Path (Join-Path $PayloadDir "*") -Destination $InstallDir -Recurse -Force

if ($Restart -eq "1") {
    Start-Process -FilePath $CurrentExe -ArgumentList @("--config", $ConfigPath) -WorkingDirectory $InstallDir
}
'''
