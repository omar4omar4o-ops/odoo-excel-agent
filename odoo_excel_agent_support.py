"""Shared helpers for the Odoo Excel agent UI and background watcher."""

from __future__ import annotations

import hashlib
import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pywintypes
import win32cred

from link_odoo_vendor_bills import (
    DEFAULT_ODOO_URL,
    ETRANGER_WORKBOOK_FILE_NAME,
    LOCAL_WORKBOOK_FILE_NAME,
    PERFORMANCE_MODE_LIVE,
    PERFORMANCE_MODE_SILENT,
)


APP_NAME = "Odoo Excel Agent"
APP_DIR_NAME = "OdooExcelAgent"
APP_VERSION = "2026.05.09.1"
DEFAULT_UPDATE_URL = "https://api.github.com/repos/omar4omar4o-ops/odoo-excel-agent/releases/latest"
AGENT_SCRIPT = "odoo_excel_background.py"
UI_SCRIPT = "odoo_excel_agent_ui.py"
STARTUP_SHORTCUT = Path(os.getenv("APPDATA", str(Path.home()))) / "Microsoft/Windows/Start Menu/Programs/Startup/Odoo Excel Agent.lnk"
DEFAULT_INSTALL_DIR = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / APP_DIR_NAME
DEFAULT_SETTLE_SECONDS = 15
DEFAULT_RETRY_DELAY_SECONDS = 45
DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS = 1
CONFIG_VERSION = 4
WATCH_MODE_SELECTED_WORKBOOKS = "selected_workbooks"
WATCH_MODE_ACHATS_PAIR = "achats_pair"
WATCH_MODE_FILE = "file"
WATCH_MODE_FOLDER = "folder"
PERFORMANCE_MODES = {PERFORMANCE_MODE_SILENT, PERFORMANCE_MODE_LIVE}
SUPPORTED_EXCEL_SESSION_BACKEND = "pywin32"
SELECTED_WORKBOOK_KEYS = (
    "achats_local_file",
    "achats_etranger_file",
    "seller_previous_file",
)


class AgentCredentialError(RuntimeError):
    """Raised when the stored credential is missing or invalid."""


@dataclass(frozen=True)
class RuntimeStatus:
    state: str
    message: str
    updated_at: str = ""
    last_issue_code: str = ""
    last_issue_message: str = ""


def expand_path(raw_path: str | Path) -> Path:
    return Path(os.path.expandvars(os.path.expanduser(str(raw_path)))).resolve()


def default_watch_folder() -> str:
    return str(Path.home() / "Downloads")


def default_runtime_paths(install_dir: Path) -> dict[str, Path]:
    return {
        "report_dir": install_dir / "reports",
        "backup_dir": install_dir / "backups",
        "state_file": install_dir / "state.json",
        "log_file": install_dir / "agent.log",
        "runtime_status_file": install_dir / "runtime_status.json",
    }


def make_credential_target(config_path: Path, db: str, login: str) -> str:
    payload = f"{expand_path(config_path)}|{db.strip()}|{login.strip()}".encode("utf-8")
    digest = hashlib.sha1(payload).hexdigest()[:12]
    return f"{APP_DIR_NAME}:{digest}"


def credential_exists(target: str) -> bool:
    try:
        read_secret(target)
        return True
    except AgentCredentialError:
        return False


def store_secret(target: str, username: str, secret: str) -> None:
    credential = {
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": target,
        "UserName": username or target,
        "CredentialBlob": secret,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        "Comment": APP_NAME,
        "Attributes": [],
    }
    win32cred.CredWrite(credential, 0)


def read_secret(target: str) -> str:
    if not target:
        raise AgentCredentialError("No credential target is configured.")
    try:
        credential = win32cred.CredRead(target, win32cred.CRED_TYPE_GENERIC)
    except pywintypes.error as exc:
        raise AgentCredentialError(f"Credential '{target}' was not found.") from exc
    blob = credential.get("CredentialBlob") or b""
    if isinstance(blob, bytes):
        return blob.decode("utf-16-le")
    return str(blob)


def delete_secret(target: str) -> None:
    if not target:
        return
    try:
        win32cred.CredDelete(target, win32cred.CRED_TYPE_GENERIC, 0)
    except pywintypes.error:
        pass


def normalize_watch_folders(raw_value: Any) -> list[str]:
    folders = raw_value if isinstance(raw_value, list) else [raw_value]
    normalized: list[str] = []
    for item in folders:
        value = str(item or "").strip()
        if not value:
            continue
        if value.casefold() == "default":
            value = default_watch_folder()
        normalized.append(str(expand_path(value)))
    return normalized


def normalize_watch_mode(raw_value: Any) -> str:
    value = str(raw_value or "").strip().casefold()
    if value == WATCH_MODE_FILE:
        return WATCH_MODE_FILE
    if value in {WATCH_MODE_SELECTED_WORKBOOKS, WATCH_MODE_ACHATS_PAIR}:
        return WATCH_MODE_SELECTED_WORKBOOKS
    return WATCH_MODE_FOLDER


def normalize_optional_path(raw_value: Any) -> str:
    value = str(raw_value or "").strip()
    if not value:
        return ""
    if value.casefold() == "default":
        value = default_watch_folder()
    return str(expand_path(value))


def normalize_performance_mode(raw_value: Any) -> str:
    value = str(raw_value or "").strip().casefold()
    if value == PERFORMANCE_MODE_LIVE:
        return PERFORMANCE_MODE_LIVE
    return PERFORMANCE_MODE_SILENT


def normalize_excel_session_backend(raw_value: Any, messages: list[str] | None = None) -> str:
    value = str(raw_value or "").strip().casefold()
    if value == "xlwings":
        if messages is not None:
            messages.append("Migrated legacy xlwings Excel backend to pywin32.")
        return SUPPORTED_EXCEL_SESSION_BACKEND
    if value and value != SUPPORTED_EXCEL_SESSION_BACKEND and messages is not None:
        messages.append(f"Unsupported Excel backend '{value}' was reset to pywin32.")
    return SUPPORTED_EXCEL_SESSION_BACKEND


def achats_slot_for_path(raw_value: Any) -> str:
    value = str(raw_value or "").strip()
    if not value:
        return ""
    name = Path(value).name.casefold()
    if name == LOCAL_WORKBOOK_FILE_NAME.casefold():
        return "achats_local_file"
    if name == ETRANGER_WORKBOOK_FILE_NAME.casefold():
        return "achats_etranger_file"
    return ""


def default_config(install_dir: Path) -> dict[str, Any]:
    paths = default_runtime_paths(install_dir)
    return {
        "version": CONFIG_VERSION,
        "odoo": {
            "url": DEFAULT_ODOO_URL,
            "db": "",
            "login": "",
            "credential_target": "",
            "record_url_example": "",
        },
        "manual": {
            "last_file": "",
        },
        "background": {
            "watch_mode": WATCH_MODE_SELECTED_WORKBOOKS,
            "performance_mode": PERFORMANCE_MODE_SILENT,
            "achats_local_file": "",
            "achats_etranger_file": "",
            "seller_previous_file": "",
            "watch_file": "",
            "watch_folder": default_watch_folder(),
            "recursive": False,
            "process_existing_on_start": False,
            "update_open_workbook": False,
            "excel_event_monitoring": False,
            "excel_session_backend": SUPPORTED_EXCEL_SESSION_BACKEND,
            "excel_save_debounce_seconds": DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS,
            "allow_live_update_with_autosave": False,
            "visible_excel": False,
            "write_report_file": False,
            "stable_backup_name": True,
            "settle_seconds": DEFAULT_SETTLE_SECONDS,
            "retry_delay_seconds": DEFAULT_RETRY_DELAY_SECONDS,
        },
        "updates": {
            "manifest_url": DEFAULT_UPDATE_URL,
            "last_checked_at": "",
            "last_seen_version": "",
        },
        "paths": {key: str(value) for key, value in paths.items()},
    }


def load_normalized_config(config_path: Path) -> tuple[dict[str, Any], list[str]]:
    install_dir = expand_path(config_path).parent
    config = default_config(install_dir)
    messages: list[str] = []
    if not config_path.exists():
        return config, messages

    try:
        raw = json.loads(config_path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        messages.append(f"Config file is unreadable; loaded safe defaults instead: {exc}")
        return config, messages
    if not isinstance(raw, dict):
        messages.append("Config file is not a JSON object; loaded safe defaults instead.")
        return config, messages
    odoo_raw = raw.get("odoo", {}) if isinstance(raw, dict) else {}
    processing_raw = raw.get("processing", {}) if isinstance(raw, dict) else {}
    background_raw = raw.get("background", {}) if isinstance(raw, dict) else {}
    manual_raw = raw.get("manual", {}) if isinstance(raw, dict) else {}
    paths_raw = raw.get("paths", {}) if isinstance(raw, dict) else {}
    updates_raw = raw.get("updates", {}) if isinstance(raw, dict) else {}

    watch_folders = normalize_watch_folders(background_raw.get("watch_folders"))
    if not watch_folders:
        watch_folders = normalize_watch_folders(raw.get("watch_folders") if isinstance(raw, dict) else [])
    if not watch_folders:
        watch_folders = [default_watch_folder()]

    achats_local_file = normalize_optional_path(background_raw.get("achats_local_file"))
    achats_etranger_file = normalize_optional_path(background_raw.get("achats_etranger_file"))
    seller_previous_file = normalize_optional_path(background_raw.get("seller_previous_file"))
    has_selected_config = bool(achats_local_file or achats_etranger_file or seller_previous_file)
    excel_session_backend = normalize_excel_session_backend(background_raw.get("excel_session_backend"), messages)
    performance_mode = normalize_performance_mode(background_raw.get("performance_mode"))
    live_mode = performance_mode == PERFORMANCE_MODE_LIVE

    watch_mode_raw = background_raw.get("watch_mode")
    if not watch_mode_raw:
        watch_mode_raw = WATCH_MODE_SELECTED_WORKBOOKS if has_selected_config else (WATCH_MODE_FOLDER if watch_folders else WATCH_MODE_FILE)
    watch_mode = normalize_watch_mode(watch_mode_raw)
    watch_file = normalize_optional_path(
        background_raw.get("watch_file") or (manual_raw.get("last_file") if watch_mode == WATCH_MODE_FILE else "")
    )
    watch_folder = normalize_optional_path(background_raw.get("watch_folder") or watch_folders[0] or default_watch_folder())
    migrated_watch_file_slot = achats_slot_for_path(watch_file)
    if migrated_watch_file_slot == "achats_local_file" and not achats_local_file:
        achats_local_file = watch_file
        messages.append("Migrated legacy ACHATS LOCAL watch file into the dedicated ACHATS setting.")
    elif migrated_watch_file_slot == "achats_etranger_file" and not achats_etranger_file:
        achats_etranger_file = watch_file
        messages.append("Migrated legacy ACHATS ETRANGER watch file into the dedicated ACHATS setting.")
    has_selected_config = bool(achats_local_file or achats_etranger_file or seller_previous_file)
    if watch_mode == WATCH_MODE_FILE and migrated_watch_file_slot:
        watch_mode = WATCH_MODE_SELECTED_WORKBOOKS
    if watch_mode == WATCH_MODE_SELECTED_WORKBOOKS and not has_selected_config and watch_file and watch_folders:
        watch_mode = WATCH_MODE_FOLDER
    if watch_mode == WATCH_MODE_FILE and not watch_file and watch_folders:
        watch_mode = WATCH_MODE_FOLDER
    if watch_mode == WATCH_MODE_FOLDER and not watch_folder:
        watch_folder = default_watch_folder()

    backup_dir = (
        background_raw.get("backup_dir")
        or processing_raw.get("backup_dir")
        or paths_raw.get("backup_dir")
        or config["paths"]["backup_dir"]
    )
    legacy_report_dir = raw.get("report_dir") if isinstance(raw, dict) else None
    legacy_state_file = raw.get("state_file") if isinstance(raw, dict) else None
    legacy_log_file = raw.get("log_file") if isinstance(raw, dict) else None
    legacy_runtime_status_file = raw.get("runtime_status_file") if isinstance(raw, dict) else None

    report_dir = (
        processing_raw.get("report_dir")
        or background_raw.get("report_dir")
        or paths_raw.get("report_dir")
        or legacy_report_dir
    )
    state_file = paths_raw.get("state_file") or legacy_state_file
    log_file = paths_raw.get("log_file") or legacy_log_file
    runtime_status_file = paths_raw.get("runtime_status_file") or legacy_runtime_status_file

    config["version"] = CONFIG_VERSION
    config["odoo"] = {
        "url": str(odoo_raw.get("url") or DEFAULT_ODOO_URL).strip(),
        "db": str(odoo_raw.get("db") or "").strip(),
        "login": str(odoo_raw.get("login") or "").strip(),
        "credential_target": str(odoo_raw.get("credential_target") or "").strip(),
        "record_url_example": str(odoo_raw.get("record_url_example") or "").strip(),
    }
    config["manual"] = {
        "last_file": str(manual_raw.get("last_file") or raw.get("workbook") or "").strip() if isinstance(raw, dict) else "",
    }
    config["background"] = {
        "watch_mode": watch_mode,
        "performance_mode": performance_mode,
        "achats_local_file": achats_local_file,
        "achats_etranger_file": achats_etranger_file,
        "seller_previous_file": seller_previous_file,
        "watch_file": watch_file,
        "watch_folder": watch_folder,
        "recursive": bool(background_raw.get("recursive", processing_raw.get("recursive", False))),
        "process_existing_on_start": bool(
            background_raw.get("process_existing_on_start", processing_raw.get("process_existing_on_start", False))
        ),
        "update_open_workbook": bool(
            background_raw.get("update_open_workbook", processing_raw.get("update_open_workbook", False))
        ) if live_mode else False,
        "excel_event_monitoring": bool(background_raw.get("excel_event_monitoring", False)) if live_mode else False,
        "excel_session_backend": excel_session_backend,
        "excel_save_debounce_seconds": int(
            background_raw.get(
                "excel_save_debounce_seconds",
                processing_raw.get("excel_save_debounce_seconds", DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS),
            )
        ),
        "allow_live_update_with_autosave": bool(background_raw.get("allow_live_update_with_autosave", False)) if live_mode else False,
        "visible_excel": bool(background_raw.get("visible_excel", processing_raw.get("visible_excel", False))),
        "write_report_file": bool(background_raw.get("write_report_file", processing_raw.get("write_report_file", False))),
        "stable_backup_name": bool(background_raw.get("stable_backup_name", processing_raw.get("stable_backup_name", True))),
        "settle_seconds": int(background_raw.get("settle_seconds", processing_raw.get("settle_seconds", DEFAULT_SETTLE_SECONDS))),
        "retry_delay_seconds": int(
            background_raw.get("retry_delay_seconds", processing_raw.get("retry_delay_seconds", DEFAULT_RETRY_DELAY_SECONDS))
        ),
    }
    config["paths"] = {
        "backup_dir": str(expand_path(backup_dir or config["paths"]["backup_dir"])),
        "report_dir": str(expand_path(report_dir or config["paths"]["report_dir"])),
        "state_file": str(expand_path(state_file or config["paths"]["state_file"])),
        "log_file": str(expand_path(log_file or config["paths"]["log_file"])),
        "runtime_status_file": str(expand_path(runtime_status_file or config["paths"]["runtime_status_file"])),
    }
    config["updates"] = {
        "manifest_url": str(updates_raw.get("manifest_url") or DEFAULT_UPDATE_URL).strip(),
        "last_checked_at": str(updates_raw.get("last_checked_at") or "").strip(),
        "last_seen_version": str(updates_raw.get("last_seen_version") or "").strip(),
    }

    if watch_folders and raw.get("watch_folders") and any(str(item).strip().casefold() == "default" for item in raw.get("watch_folders", [])):
        messages.append("Normalized legacy watch folder 'default' to Downloads.")

    legacy_api_key = str(odoo_raw.get("api_key") or "").strip()
    if legacy_api_key and not config["odoo"]["credential_target"]:
        config["odoo"]["credential_target"] = make_credential_target(config_path, config["odoo"]["db"], config["odoo"]["login"])
        messages.append("Legacy plain-text API key detected and ready for Credential Manager migration.")

    return config, messages


def save_normalized_config(config_path: Path, config: dict[str, Any]) -> None:
    config_path.parent.mkdir(parents=True, exist_ok=True)
    data = json.dumps(config, indent=2)
    config_path.write_text(data, encoding="utf-8")


def get_background_watch_targets(config: dict[str, Any]) -> list[Path]:
    background = config.get("background", {})
    mode = normalize_watch_mode(background.get("watch_mode"))
    if mode == WATCH_MODE_SELECTED_WORKBOOKS:
        targets: list[Path] = []
        seen: set[Path] = set()
        for key in SELECTED_WORKBOOK_KEYS:
            raw_target = str(background.get(key) or "").strip()
            if not raw_target:
                continue
            target = expand_path(raw_target)
            if target in seen:
                continue
            seen.add(target)
            targets.append(target)
        return targets
    if mode == WATCH_MODE_FILE:
        raw_file = str(background.get("watch_file") or "").strip()
        if not raw_file:
            return []
        return [expand_path(raw_file)]
    raw_folder = str(background.get("watch_folder") or "").strip()
    if raw_folder:
        return [expand_path(raw_folder)]
    return []


def get_background_watch_folders(config: dict[str, Any]) -> list[Path]:
    return get_background_watch_targets(config)


def get_paths(config: dict[str, Any]) -> dict[str, Path]:
    raw_paths = config.get("paths", {})
    return {key: expand_path(value) for key, value in raw_paths.items()}


def get_runtime_status_path(config_path: Path, config: dict[str, Any] | None = None) -> Path:
    if config is None:
        config, _ = load_normalized_config(config_path)
    return expand_path(config.get("paths", {}).get("runtime_status_file") or default_runtime_paths(config_path.parent)["runtime_status_file"])


def read_runtime_status(status_path: Path) -> RuntimeStatus:
    if not status_path.exists():
        return RuntimeStatus(state="stopped", message="No runtime status file yet.")
    try:
        raw = json.loads(status_path.read_text(encoding="utf-8"))
    except Exception:
        return RuntimeStatus(state="startup_failed", message="Runtime status file is unreadable.")
    return RuntimeStatus(
        state=str(raw.get("state") or "unknown"),
        message=str(raw.get("message") or ""),
        updated_at=str(raw.get("updated_at") or ""),
        last_issue_code=str(raw.get("last_issue_code") or ""),
        last_issue_message=str(raw.get("last_issue_message") or ""),
    )


def write_runtime_status(
    status_path: Path,
    state: str,
    message: str,
    *,
    updated_at: str,
    last_issue_code: str = "",
    last_issue_message: str = "",
) -> None:
    status_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "state": state,
        "message": message,
        "updated_at": updated_at,
        "last_issue_code": last_issue_code,
        "last_issue_message": last_issue_message,
    }
    tmp_path = status_path.with_suffix(status_path.suffix + ".tmp")
    tmp_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    tmp_path.replace(status_path)
