"""Background watcher that auto-links Excel files to Odoo purchase orders."""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import os
import subprocess
import sys
import threading
import time
import traceback
from dataclasses import dataclass
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

import pythoncom
import win32api
import win32con
import win32event
import win32file
import win32gui
import win32com.client
import winerror
from watchdog.events import FileSystemEventHandler, FileSystemMovedEvent
from watchdog.observers import Observer

try:
    import xlwings as xw  # type: ignore[import-not-found]
except ImportError:
    xw = None

from link_odoo_vendor_bills import (
    DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS,
    WorkbookAccessError,
    WorkbookAccessContext,
    WorkbookProcessSummary,
    WORKBOOK_SLOT_ACHATS_ETRANGER,
    WORKBOOK_SLOT_ACHATS_LOCAL,
    WORKBOOK_SLOT_SELLER_PREVIOUS,
    PERFORMANCE_MODE_LIVE,
    PERFORMANCE_MODE_SILENT,
    inspect_workbook_access_state,
    is_expected_excel_save,
    is_managed_workbook,
    is_supported_workbook,
    process_workbook,
    validate_odoo_settings,
)
from odoo_excel_agent_support import (
    AGENT_SCRIPT,
    APP_NAME,
    AgentCredentialError,
    expand_path,
    get_background_watch_targets,
    get_runtime_status_path,
    get_paths,
    load_normalized_config,
    read_secret,
    write_runtime_status,
    WATCH_MODE_ACHATS_PAIR,
    WATCH_MODE_FILE,
    WATCH_MODE_FOLDER,
    WATCH_MODE_SELECTED_WORKBOOKS,
)


MUTEX_NAME = "Local\\OdooExcelAgentSingleton"
ICON_MESSAGE_ID = win32con.WM_USER + 20
MENU_OPEN_UI = 1001
MENU_OPEN_LOG = 1002
MENU_OPEN_BACKUPS = 1003
MENU_SCAN_NOW = 1004
MENU_EXIT = 1005


@dataclass(frozen=True)
class OdooSettings:
    url: str
    db: str
    login: str
    api_key: str
    record_url_example: str


@dataclass(frozen=True)
class ProcessingSettings:
    watch_mode: str
    watch_file: Path | None
    watch_folder: Path | None
    achats_local_file: Path | None
    achats_etranger_file: Path | None
    seller_previous_file: Path | None
    watch_targets: tuple[Path, ...]
    backup_dir: Path
    report_dir: Path
    state_file: Path
    log_file: Path
    runtime_status_file: Path
    settle_seconds: int
    retry_delay_seconds: int
    recursive: bool
    performance_mode: str
    process_existing_on_start: bool
    update_open_workbook: bool
    excel_event_monitoring: bool
    excel_session_backend: str
    excel_save_debounce_seconds: int
    allow_live_update_with_autosave: bool
    visible_excel: bool
    write_report_file: bool
    stable_backup_name: bool
    config_path: Path


@dataclass(frozen=True)
class AgentConfig:
    odoo: OdooSettings
    processing: ProcessingSettings


def timestamp() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def runtime_status(config: AgentConfig, state: str, message: str, *, last_issue_code: str = "", last_issue_message: str = "") -> None:
    write_runtime_status(
        config.processing.runtime_status_file,
        state,
        message,
        updated_at=timestamp(),
        last_issue_code=last_issue_code,
        last_issue_message=last_issue_message,
    )


def configure_logging(log_file: Path) -> logging.Logger:
    logger = logging.getLogger(APP_NAME)
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = RotatingFileHandler(log_file, maxBytes=500_000, backupCount=1, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    logger.propagate = False
    return logger


def current_pythonw() -> str:
    executable = Path(sys.executable)
    if executable.name.lower() == "pythonw.exe":
        return str(executable)
    candidate = executable.with_name("pythonw.exe")
    if candidate.exists():
        return str(candidate)
    return str(executable)


def load_agent_config(config_path: Path) -> AgentConfig:
    raw: dict[str, Any] = {}
    if config_path.exists():
        try:
            import json

            loaded = json.loads(config_path.read_text(encoding="utf-8-sig"))
            if isinstance(loaded, dict):
                raw = loaded
        except Exception:
            raw = {}
    normalized, _ = load_normalized_config(config_path)
    watch_targets = tuple(get_background_watch_targets(normalized))
    paths = get_paths(normalized)
    odoo_raw = normalized["odoo"]
    background = normalized["background"]
    credential_target = str(odoo_raw.get("credential_target") or "").strip()
    legacy_api_key = str(raw.get("odoo", {}).get("api_key") or "").strip() if isinstance(raw.get("odoo"), dict) else ""
    try:
        api_key = read_secret(credential_target) if credential_target else ""
    except AgentCredentialError:
        api_key = legacy_api_key
    if not api_key:
        raise AgentCredentialError("No Odoo API key is available. Open the setup UI and save the Odoo settings again.")
    try:
        validate_odoo_settings(
            str(odoo_raw.get("url") or "").strip(),
            str(odoo_raw.get("db") or "").strip(),
            str(odoo_raw.get("login") or "").strip(),
            str(odoo_raw.get("record_url_example") or "").strip(),
        )
    except ValueError as exc:
        raise ValueError(f"Invalid Odoo settings in config: {exc}") from exc
    watch_mode = str(background.get("watch_mode") or WATCH_MODE_SELECTED_WORKBOOKS).strip().casefold()
    if watch_mode not in {WATCH_MODE_FILE, WATCH_MODE_FOLDER, WATCH_MODE_SELECTED_WORKBOOKS, WATCH_MODE_ACHATS_PAIR}:
        watch_mode = WATCH_MODE_FOLDER
    if watch_mode == WATCH_MODE_ACHATS_PAIR:
        watch_mode = WATCH_MODE_SELECTED_WORKBOOKS
    watch_file = watch_targets[0] if watch_mode == WATCH_MODE_FILE and watch_targets else None
    watch_folder = watch_targets[0] if watch_mode == WATCH_MODE_FOLDER and watch_targets else None
    achats_local_file = expand_path(background["achats_local_file"]) if str(background.get("achats_local_file") or "").strip() else None
    achats_etranger_file = expand_path(background["achats_etranger_file"]) if str(background.get("achats_etranger_file") or "").strip() else None
    seller_previous_file = expand_path(background["seller_previous_file"]) if str(background.get("seller_previous_file") or "").strip() else None
    return AgentConfig(
        odoo=OdooSettings(
            url=str(odoo_raw.get("url") or "").strip(),
            db=str(odoo_raw.get("db") or "").strip(),
            login=str(odoo_raw.get("login") or "").strip(),
            api_key=api_key,
            record_url_example=str(odoo_raw.get("record_url_example") or "").strip(),
        ),
        processing=ProcessingSettings(
            watch_mode=watch_mode,
            watch_file=watch_file,
            watch_folder=watch_folder,
            achats_local_file=achats_local_file,
            achats_etranger_file=achats_etranger_file,
            seller_previous_file=seller_previous_file,
            watch_targets=watch_targets,
            backup_dir=paths["backup_dir"],
            report_dir=paths["report_dir"],
            state_file=paths["state_file"],
            log_file=paths["log_file"],
            runtime_status_file=paths["runtime_status_file"],
            settle_seconds=int(background["settle_seconds"]),
            retry_delay_seconds=int(background["retry_delay_seconds"]),
            recursive=bool(background["recursive"]),
            performance_mode=str(background.get("performance_mode") or PERFORMANCE_MODE_SILENT),
            process_existing_on_start=bool(background["process_existing_on_start"]),
            update_open_workbook=bool(background.get("update_open_workbook", False)),
            excel_event_monitoring=bool(background.get("excel_event_monitoring", False)),
            excel_session_backend=str(background.get("excel_session_backend") or "pywin32"),
            excel_save_debounce_seconds=int(background.get("excel_save_debounce_seconds", DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS)),
            allow_live_update_with_autosave=bool(background.get("allow_live_update_with_autosave", False)),
            visible_excel=bool(background["visible_excel"]),
            write_report_file=bool(background["write_report_file"]),
            stable_backup_name=bool(background["stable_backup_name"]),
            config_path=config_path,
        ),
    )


class StateStore:
    def __init__(self, state_file: Path) -> None:
        self.state_file = state_file
        self._lock = threading.Lock()
        self._data = self._load()

    def _load(self) -> dict[str, dict[str, int]]:
        if not self.state_file.exists():
            return {}
        try:
            import json

            loaded = json.loads(self.state_file.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                return {
                    str(path): {
                        "size": int(value.get("size", 0)),
                        "mtime_ns": int(value.get("mtime_ns", 0)),
                    }
                    for path, value in loaded.items()
                    if isinstance(value, dict)
                }
        except Exception:
            pass
        return {}

    def save(self) -> None:
        with self._lock:
            import json

            tmp_path = self.state_file.with_suffix(self.state_file.suffix + ".tmp")
            tmp_path.parent.mkdir(parents=True, exist_ok=True)
            tmp_path.write_text(json.dumps(self._data, indent=2), encoding="utf-8")
            tmp_path.replace(self.state_file)

    def is_processed(self, path: Path, fingerprint: dict[str, int]) -> bool:
        with self._lock:
            return self._data.get(str(path)) == fingerprint

    def mark_processed(self, path: Path, fingerprint: dict[str, int]) -> None:
        with self._lock:
            self._data[str(path)] = fingerprint


def file_fingerprint(path: Path) -> dict[str, int]:
    stat = path.stat()
    return {"size": int(stat.st_size), "mtime_ns": int(stat.st_mtime_ns)}


def excel_lock_marker_path(path: Path) -> Path:
    return path.with_name(f"~${path.name}")


def has_excel_lock_marker(path: Path) -> bool:
    return excel_lock_marker_path(path).exists()


def can_open_exclusively(path: Path) -> bool:
    handle = None
    try:
        handle = win32file.CreateFile(
            str(path),
            win32con.GENERIC_READ | win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None,
        )
        return True
    except Exception:
        return False
    finally:
        if handle is not None:
            try:
                win32file.CloseHandle(handle)
            except Exception:
                pass


def file_is_stable(path: Path, delay_seconds: float = 1.0) -> bool:
    first = file_fingerprint(path)
    time.sleep(delay_seconds)
    return first == file_fingerprint(path)


class WorkbookEventHandler(FileSystemEventHandler):
    def __init__(self, app: "OdooExcelAgent") -> None:
        self.app = app

    def on_created(self, event: Any) -> None:
        if not event.is_directory:
            self.app.schedule_path(Path(event.src_path), "created")

    def on_modified(self, event: Any) -> None:
        if not event.is_directory:
            self.app.schedule_path(Path(event.src_path), "modified")

    def on_closed(self, event: Any) -> None:
        if not event.is_directory:
            self.app.schedule_path(Path(event.src_path), "closed")

    def on_moved(self, event: FileSystemMovedEvent) -> None:
        if not event.is_directory:
            self.app.schedule_path(Path(event.dest_path), "moved")


class ExcelApplicationEvents:
    monitor: "ExcelEventMonitor"
    source_key: str

    def OnWorkbookAfterSave(self, workbook: Any, success: bool) -> None:
        if success:
            self.monitor.handle_workbook_after_save(self.source_key, workbook)

    def OnWorkbookBeforeClose(self, workbook: Any, cancel: bool) -> None:
        if not cancel:
            self.monitor.handle_workbook_before_close(self.source_key, workbook)

    def OnSheetChange(self, sh: Any, target: Any) -> None:
        try:
            workbook = sh.Parent
        except Exception:
            workbook = None
        if workbook is not None:
            self.monitor.handle_sheet_change(self.source_key, workbook)


class ExcelEventMonitor:
    def __init__(self, app: "OdooExcelAgent") -> None:
        self.app = app
        self.handlers: dict[str, Any] = {}
        self.thread = threading.Thread(target=self._run, name="ExcelEventMonitor", daemon=True)
        self.refresh_interval_seconds = 1.0
        self._last_refresh_at = 0.0
        self._backend_warning_reported = False
        self._xlwings_enumeration_disabled = False
        self._xlwings_warning_reported = False
        self._started = False

    def start(self) -> None:
        self.thread.start()
        self._started = True

    def stop(self) -> None:
        if self._started:
            self.thread.join(timeout=10)

    def handle_workbook_after_save(self, source_key: str, workbook: Any) -> None:
        path = self._path_from_workbook(workbook)
        if path is None or not self._should_track_event(path):
            return
        if is_expected_excel_save(path):
            self.app.logger.info("Ignoring self-triggered Excel AfterSave for %s (%s)", path, source_key)
            return
        self.app.logger.info("Excel AfterSave detected for %s (%s)", path, source_key)
        # Use a fast 1-second delay for after-save events since data is fresh.
        after_save_delay = min(self.app.config.processing.excel_save_debounce_seconds, 1)
        self.app.schedule_path(
            path,
            "excel_after_save",
            delay_override=after_save_delay,
            force=True,
            debounce=True,
        )

    def handle_workbook_before_close(self, source_key: str, workbook: Any) -> None:
        path = self._path_from_workbook(workbook)
        if path is None or not self._should_track_event(path):
            return
        if is_expected_excel_save(path):
            self.app.logger.info("Ignoring self-triggered Excel BeforeClose for %s (%s)", path, source_key)
            return
        self.app.logger.info("Excel BeforeClose detected for %s (%s)", path, source_key)
        self.app.schedule_path(path, "excel_before_close", delay_override=2, force=True)

    def handle_sheet_change(self, source_key: str, workbook: Any) -> None:
        if not self.app.config.processing.update_open_workbook:
            return
        path = self._path_from_workbook(workbook)
        if path is None or not self._should_track_event(path):
            return
        if is_expected_excel_save(path):
            return
        self.app.logger.info("Excel SheetChange detected for %s (%s)", path, source_key)
        self.app.schedule_path(
            path,
            "excel_sheet_change",
            delay_override=self.app.config.processing.excel_save_debounce_seconds,
            force=True,
            debounce=True,
        )

    def _run(self) -> None:
        pythoncom.CoInitialize()
        try:
            while not self.app.stop_event.is_set():
                self._refresh_handlers_if_needed()
                pythoncom.PumpWaitingMessages()
                self.app.stop_event.wait(0.2)
        finally:
            self.handlers.clear()
            pythoncom.CoUninitialize()

    def _refresh_handlers_if_needed(self) -> None:
        now = time.time()
        if now - self._last_refresh_at < self.refresh_interval_seconds:
            return
        self._last_refresh_at = now
        targets = self._list_targets()
        active_keys = {key for key, _ in targets}
        for stale_key in list(self.handlers):
            if stale_key not in active_keys:
                del self.handlers[stale_key]
        for key, application in targets:
            if key in self.handlers:
                continue
            try:
                handler = win32com.client.WithEvents(application, ExcelApplicationEvents)
                handler.monitor = self
                handler.source_key = key
                self.handlers[key] = handler
                self.app.logger.info("Attached Excel event handler: %s", key)
            except Exception as exc:
                self.app.logger.warning("Failed to attach Excel event handler %s: %s", key, exc)

    def _list_targets(self) -> list[tuple[str, Any]]:
        targets: list[tuple[str, Any]] = []
        seen: set[str] = set()
        if xw is not None and not self._xlwings_enumeration_disabled:
            try:
                for pid in xw.apps.keys():
                    application = xw.apps[pid].api
                    try:
                        hwnd = int(getattr(application, "Hwnd", 0) or 0)
                    except Exception:
                        hwnd = 0
                    key = f"excel:{hwnd}" if hwnd else f"xlwings:{int(pid)}"
                    if key in seen:
                        continue
                    targets.append((key, application))
                    seen.add(key)
            except Exception as exc:
                self._xlwings_enumeration_disabled = True
                if not self._xlwings_warning_reported:
                    self._xlwings_warning_reported = True
                    self.app.logger.warning(
                        "Could not enumerate Excel instances via xlwings; falling back to pywin32 only: %s",
                        exc,
                    )
        elif self.app.config.processing.excel_session_backend == "xlwings" and not self._backend_warning_reported:
            self._backend_warning_reported = True
            self.app.logger.warning("Excel event monitoring backend xlwings is unavailable.")
            runtime_status(
                self.app.config,
                "running",
                "xlwings is not installed. Live Excel monitoring is unavailable.",
                last_issue_code="excel_backend_unavailable",
                last_issue_message="xlwings is not installed.",
            )

        try:
            application = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            application = None
        if application is not None:
            try:
                hwnd = int(getattr(application, "Hwnd", 0) or 0)
                source_key = f"excel:{hwnd}" if hwnd else "pywin32:active"
            except Exception:
                source_key = "pywin32:active"
            if source_key not in seen:
                targets.append((source_key, application))
        return targets

    @staticmethod
    def _path_from_workbook(workbook: Any) -> Path | None:
        try:
            raw = str(workbook.FullName or "").strip()
        except Exception:
            raw = ""
        if not raw:
            return None
        try:
            return Path(raw).expanduser().resolve()
        except Exception:
            return None

    def _should_track_event(self, path: Path) -> bool:
        if not is_supported_workbook(path):
            return False
        if is_managed_workbook(path):
            return False
        return not self.app._should_ignore(path)


class TrayIcon:
    def __init__(self, app: "OdooExcelAgent") -> None:
        self.app = app
        self.class_name = "OdooExcelAgentTrayWindow"
        self.hwnd: int | None = None
        self.notify_id: Any = None
        self.restart_message = win32gui.RegisterWindowMessage("TaskbarCreated")
        self._create_window()
        self._refresh_icon()

    def _create_window(self) -> None:
        message_map = {
            self.restart_message: self._on_restart,
            win32con.WM_DESTROY: self._on_destroy,
            win32con.WM_COMMAND: self._on_command,
            ICON_MESSAGE_ID: self._on_icon_event,
        }
        wc = win32gui.WNDCLASS()
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = self.class_name
        wc.lpfnWndProc = message_map
        try:
            class_atom = win32gui.RegisterClass(wc)
        except win32gui.error:
            class_atom = self.class_name
        self.hwnd = win32gui.CreateWindow(
            class_atom,
            APP_NAME,
            0,
            0,
            0,
            win32con.CW_USEDEFAULT,
            win32con.CW_USEDEFAULT,
            0,
            0,
            wc.hInstance,
            None,
        )
        win32gui.UpdateWindow(self.hwnd)

    def _refresh_icon(self) -> None:
        assert self.hwnd is not None
        hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
        tooltip = f"{APP_NAME} - watching Excel files"
        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (self.hwnd, 0, flags, ICON_MESSAGE_ID, hicon, tooltip)
        if self.notify_id:
            win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
        else:
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        self.notify_id = nid

    def show_notification(self, title: str, message: str) -> None:
        if not self.notify_id or self.hwnd is None:
            return
        hicon = win32gui.LoadIcon(0, win32con.IDI_INFORMATION)
        flags = win32gui.NIF_INFO
        info_flag = getattr(win32gui, "NIIF_INFO", getattr(win32con, "NIIF_INFO", 1))
        info = (
            self.hwnd,
            0,
            flags,
            ICON_MESSAGE_ID,
            hicon,
            APP_NAME,
            message[:255],
            5000,
            title[:63],
            info_flag,
        )
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, info)
        except Exception:
            pass

    def _show_menu(self) -> None:
        assert self.hwnd is not None
        menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_OPEN_UI, "Open UI")
        win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_OPEN_LOG, "Open Log")
        win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_OPEN_BACKUPS, "Open Backups")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_SCAN_NOW, "Scan Now")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_EXIT, "Exit")
        pos = win32gui.GetCursorPos()
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    def _on_icon_event(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        if lparam in (win32con.WM_RBUTTONUP, win32con.WM_CONTEXTMENU):
            self._show_menu()
        elif lparam == win32con.WM_LBUTTONDBLCLK:
            self.app.open_ui()
        return 1

    def _on_command(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        command = win32api.LOWORD(wparam)
        if command == MENU_OPEN_UI:
            self.app.open_ui()
        elif command == MENU_OPEN_LOG:
            self.app.open_log()
        elif command == MENU_OPEN_BACKUPS:
            self.app.open_backups()
        elif command == MENU_SCAN_NOW:
            self.app.scan_now()
        elif command == MENU_EXIT:
            self.app.stop()
            win32gui.DestroyWindow(hwnd)
        return 0

    def _on_restart(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        self._refresh_icon()
        return 0

    def _on_destroy(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        if self.notify_id:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, self.notify_id)
        win32gui.PostQuitMessage(0)
        return 0


class OdooExcelAgent:
    def __init__(self, config: AgentConfig, logger: logging.Logger) -> None:
        self.config = config
        self.logger = logger
        self.state = StateStore(config.processing.state_file)
        self.observer = Observer()
        self.stop_event = threading.Event()
        self.work_event = threading.Event()
        self.pending_lock = threading.Lock()
        self.pending: dict[Path, float] = {}
        self.force_pending: set[Path] = set()
        self.last_poll_at = 0.0
        self.worker_thread = threading.Thread(target=self._worker_loop, name="WorkbookWorker", daemon=True)
        self.event_monitor = ExcelEventMonitor(self) if self.config.processing.excel_event_monitoring else None
        self.tray = TrayIcon(self)

    def start(self) -> None:
        event_handler = WorkbookEventHandler(self)
        watch_roots, startup_warnings = self._resolve_watch_roots()
        for warning in startup_warnings:
            self.logger.warning(warning)
        for watch_root, recursive in watch_roots:
            self.observer.schedule(event_handler, str(watch_root), recursive=recursive)
            self.logger.info("Watching %s root: %s (recursive=%s)", self.config.processing.watch_mode, watch_root, recursive)
        self.observer.start()
        if self.event_monitor is not None:
            self.event_monitor.start()
        if self.config.processing.process_existing_on_start:
            self.scan_now()
        else:
            self._record_initial_baseline()
        self.worker_thread.start()
        message = self._runtime_watch_message(startup_warnings)
        runtime_status(self.config, "running", message)
        self.tray.show_notification(APP_NAME, "Background watcher started.")

    def stop(self) -> None:
        if self.stop_event.is_set():
            return
        self.logger.info("Stopping agent.")
        self.stop_event.set()
        if self.observer.is_alive():
            self.observer.stop()
            self.observer.join(timeout=10)
        if self.event_monitor is not None:
            self.event_monitor.stop()
        self.worker_thread.join(timeout=10)
        self.state.save()
        runtime_status(self.config, "stopped", "Background watcher stopped.")

    def open_ui(self) -> None:
        if getattr(sys, "frozen", False):
            args = [sys.executable, "--config", str(self.config.processing.config_path)]
            cwd = str(Path(sys.executable).parent)
        else:
            ui_script = self.config.processing.config_path.parent / "odoo_excel_agent_ui.py"
            args = [current_pythonw(), str(ui_script), "--config", str(self.config.processing.config_path)]
            cwd = self.config.processing.config_path.parent
        subprocess.Popen(
            args,
            cwd=cwd,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

    def open_log(self) -> None:
        os.startfile(str(self.config.processing.log_file))

    def open_backups(self) -> None:
        os.startfile(str(self.config.processing.backup_dir))

    def scan_now(self) -> None:
        self.logger.info("Manual scan requested.")
        if self.config.processing.watch_mode == WATCH_MODE_FOLDER:
            watch_folder = self.config.processing.watch_folder
            if watch_folder is not None:
                iterator = watch_folder.rglob("*") if self.config.processing.recursive else watch_folder.glob("*")
                for path in iterator:
                    if path.is_file():
                        self.schedule_path(path, "scan_now", delay_override=0, force=True)
        else:
            for watch_target in self._exact_watch_targets():
                if watch_target.exists():
                    self.schedule_path(watch_target, "scan_now", delay_override=0, force=True)
        runtime_status(self.config, "running", "Manual scan queued.")
        self.tray.show_notification(APP_NAME, "Scan queued.")

    def schedule_path(
        self,
        path: Path,
        reason: str,
        *,
        delay_override: int | float | None = None,
        force: bool = False,
        debounce: bool = False,
    ) -> None:
        if self._should_ignore(path):
            return
        normalized = path.expanduser().resolve()
        delay_seconds = self.config.processing.settle_seconds if delay_override is None else max(float(delay_override), 0.0)
        due_at = time.time() + delay_seconds
        with self.pending_lock:
            current_due = self.pending.get(normalized, 0)
            if current_due:
                if delay_override is None or debounce:
                    self.pending[normalized] = max(current_due, due_at)
                else:
                    self.pending[normalized] = min(current_due, due_at)
            else:
                self.pending[normalized] = due_at
            if force:
                self.force_pending.add(normalized)
        self.work_event.set()
        self.logger.info("Queued workbook: %s (%s)", normalized, reason)

    def _should_ignore(self, path: Path) -> bool:
        resolved = path.expanduser().resolve()
        exact_targets = self._exact_watch_targets()
        if self.config.processing.watch_mode in {WATCH_MODE_FILE, WATCH_MODE_SELECTED_WORKBOOKS}:
            if resolved not in exact_targets:
                return True
        name = path.name
        if name.startswith("~$"):
            return True
        explicitly_watched = resolved in exact_targets
        if not explicitly_watched and (".backup-" in name or ".odoo-link-report-" in name or ".original." in name):
            return True
        if not is_supported_workbook(path):
            return True
        if resolved.is_relative_to(self.config.processing.report_dir):
            return True
        if not explicitly_watched and resolved.is_relative_to(self.config.processing.backup_dir):
            return True
        return False

    def _worker_loop(self) -> None:
        while not self.stop_event.is_set():
            self._poll_watch_targets_if_needed()
            due_paths: list[tuple[Path, bool]] = []
            now = time.time()
            with self.pending_lock:
                for path, due_at in list(self.pending.items()):
                    if due_at <= now:
                        due_paths.append((path, path in self.force_pending))
                        del self.pending[path]
                        self.force_pending.discard(path)
            for path, force in due_paths:
                try:
                    self._process_path(path, force=force)
                except Exception:
                    self.logger.exception("Unhandled error while processing %s", path)
                    runtime_status(
                        self.config,
                        "running",
                        f"Error while processing {path.name}.",
                        last_issue_code="processing_error",
                        last_issue_message=path.name,
                    )
                    self.tray.show_notification(APP_NAME, f"Error while processing {path.name}. Check the log.")
            self.work_event.wait(timeout=0.3)
            self.work_event.clear()

    def _poll_watch_targets_if_needed(self) -> None:
        if self.config.processing.watch_mode == WATCH_MODE_FOLDER:
            return
        now = time.time()
        poll_seconds = min(max(self.config.processing.settle_seconds / 2, 2), 15)
        if now - self.last_poll_at < poll_seconds:
            return
        self.last_poll_at = now
        for watch_target in self._exact_watch_targets():
            if not watch_target.exists():
                continue
            try:
                fingerprint = file_fingerprint(watch_target)
            except OSError:
                continue
            normalized = watch_target.expanduser().resolve()
            with self.pending_lock:
                already_pending = normalized in self.pending
            if not already_pending and not self.state.is_processed(normalized, fingerprint):
                self.schedule_path(normalized, "poll")

    def _exact_watch_targets(self) -> tuple[Path, ...]:
        if self.config.processing.watch_mode == WATCH_MODE_FILE:
            return (self.config.processing.watch_file,) if self.config.processing.watch_file is not None else ()
        if self.config.processing.watch_mode == WATCH_MODE_SELECTED_WORKBOOKS:
            return self.config.processing.watch_targets
        return ()

    def _workbook_slot_for_path(self, path: Path) -> str:
        resolved = path.expanduser().resolve()
        slot_paths = (
            (WORKBOOK_SLOT_ACHATS_LOCAL, self.config.processing.achats_local_file),
            (WORKBOOK_SLOT_ACHATS_ETRANGER, self.config.processing.achats_etranger_file),
            (WORKBOOK_SLOT_SELLER_PREVIOUS, self.config.processing.seller_previous_file),
        )
        for slot, candidate in slot_paths:
            if candidate is not None and candidate.expanduser().resolve() == resolved:
                return slot
        return ""

    def _resolve_watch_roots(self) -> tuple[list[tuple[Path, bool]], list[str]]:
        if self.config.processing.watch_mode == WATCH_MODE_FILE:
            watch_file = self.config.processing.watch_file
            if watch_file is None:
                raise ValueError("No watch file is configured.")
            if not watch_file.exists() or not watch_file.is_file():
                raise ValueError(f"Invalid watch target: {watch_file}")
            if not is_supported_workbook(watch_file):
                raise ValueError(f"Invalid watch target: {watch_file}")
            return [(watch_file.parent, False)], []

        if self.config.processing.watch_mode == WATCH_MODE_FOLDER:
            watch_folder = self.config.processing.watch_folder
            if watch_folder is None:
                raise ValueError("No watch folder is configured.")
            if not watch_folder.exists() or not watch_folder.is_dir():
                raise ValueError(f"Invalid watch target: {watch_folder}")
            return [(watch_folder, self.config.processing.recursive)], []

        configured_targets = list(self.config.processing.watch_targets)
        if not configured_targets:
            raise ValueError("At least one selected workbook must be configured.")

        warnings: list[str] = []
        roots: list[tuple[Path, bool]] = []
        seen_roots: set[Path] = set()
        valid_targets = 0
        for target in configured_targets:
            if target.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
                raise ValueError(f"Invalid watch target: {target}")
            if target.exists():
                if not target.is_file():
                    raise ValueError(f"Invalid watch target: {target}")
                valid_targets += 1
            else:
                warnings.append(f"Missing selected workbook: {target.name}")
            parent = target.parent
            if not parent.exists() or not parent.is_dir():
                warnings.append(f"Missing watch folder: {parent}")
                continue
            if parent in seen_roots:
                continue
            seen_roots.add(parent)
            roots.append((parent, False))

        if not roots:
            raise ValueError("No valid parent folder is available for the selected workbooks.")

        if valid_targets == 0:
            self.logger.warning("Selected-workbook mode started without an existing target file yet.")
        return roots, warnings

    def _record_initial_baseline(self) -> None:
        for watch_target in self._exact_watch_targets():
            if watch_target is None or not watch_target.exists():
                continue
            try:
                fingerprint = file_fingerprint(watch_target)
                self.state.mark_processed(watch_target.expanduser().resolve(), fingerprint)
                self.logger.info("Baseline recorded for watched file: %s", watch_target)
            except OSError:
                self.logger.info("Could not record initial baseline for watched file yet: %s", watch_target)
        self.state.save()

    def _runtime_watch_message(self, startup_warnings: list[str]) -> str:
        if self.config.processing.watch_mode == WATCH_MODE_FILE and self.config.processing.watch_file is not None:
            message = f"Watching file: {self.config.processing.watch_file.name}"
        elif self.config.processing.watch_mode == WATCH_MODE_FOLDER and self.config.processing.watch_folder is not None:
            message = f"Watching folder: {self.config.processing.watch_folder}"
        else:
            configured = len(self.config.processing.watch_targets)
            existing = sum(1 for path in self.config.processing.watch_targets if path.exists() and path.is_file())
            message = f"Watching {existing}/{configured} selected workbooks."
        if startup_warnings:
            warning_text = "; ".join(startup_warnings[:2])
            message = f"{message} {warning_text}"
        return message

    def _process_path(self, path: Path, *, force: bool = False) -> None:
        if self._should_ignore(path):
            return
        if not path.exists():
            self.logger.info("Skipping missing workbook: %s", path)
            return

        try:
            fingerprint_before = file_fingerprint(path)
        except OSError:
            self._retry_later(path, "Could not stat workbook yet.", "file_unavailable")
            return

        silent_mode = self.config.processing.performance_mode != PERFORMANCE_MODE_LIVE
        if silent_mode:
            workbook_is_locked = has_excel_lock_marker(path) or not can_open_exclusively(path)
            if workbook_is_locked:
                self._retry_later(
                    path,
                    "Workbook is open or locked. Waiting for it to close.",
                    "waiting_for_close",
                )
                return
            access = WorkbookAccessContext(status="closed", workbook_path=path, backend="openpyxl")
        else:
            access = inspect_workbook_access_state(
                path,
                excel_session_backend=self.config.processing.excel_session_backend,
                allow_live_update_with_autosave=self.config.processing.allow_live_update_with_autosave,
            )

        if self.state.is_processed(path, fingerprint_before) and not force:
            self.logger.info("Skipping unchanged workbook: %s", path)
            return

        # --- Handle open workbook states ---
        if access.status == "open_writable":
            if not self.config.processing.update_open_workbook:
                self._retry_later(path, "Workbook is open in Excel. Waiting for it to close.", "excel_waiting_close")
                return
            # Open & writable: skip all stability checks, process immediately via live session.
            self.logger.info("Workbook is open and writable; using live Excel session: %s", path)
        elif access.status == "open_read_only":
            self._retry_later(path, "Workbook is open in Excel as read-only.", "excel_read_only")
            return
        elif access.status == "open_autosave":
            self._retry_later(path, "Workbook is open with AutoSave enabled. Waiting for it to close.", "excel_autosave_deferred")
            return
        elif access.status == "open_ambiguous_instance":
            self._retry_later(path, "Workbook is open in multiple Excel instances. Waiting for it to close.", "excel_ambiguous_instance")
            return
        elif access.status == "unsupported_live_update":
            workbook_is_locked = has_excel_lock_marker(path) or not can_open_exclusively(path)
            if workbook_is_locked:
                reason = access.details or "Live Excel monitoring is unavailable while the workbook is still open."
                self._retry_later(path, reason, "excel_backend_unavailable")
                return
            # Not locked, treat as closed.
            access = WorkbookAccessContext(status="closed", workbook_path=path, backend=access.backend)
        elif access.status == "closed":
            workbook_is_locked = has_excel_lock_marker(path) or not can_open_exclusively(path)
            if workbook_is_locked and not self.config.processing.update_open_workbook:
                self._retry_later(path, "Workbook is still locked. Waiting for it to close.", "excel_waiting_close")
                return

        # --- Stability checks only for closed workbooks ---
        if access.status == "closed":
            age_seconds = time.time() - path.stat().st_mtime
            if age_seconds < self.config.processing.settle_seconds:
                self._retry_later(path, "Workbook is still changing.", "file_changing")
                return
            try:
                if not file_is_stable(path, delay_seconds=0.5):
                    self._retry_later(path, "Workbook write is not stable yet.", "file_changing")
                    return
            except OSError:
                self._retry_later(path, "Workbook is temporarily unavailable.", "file_unavailable")
                return

        self.logger.info("Processing workbook: %s", path)
        try:
            summary = process_workbook(
                workbook_path=path,
                odoo_url=self.config.odoo.url,
                odoo_db=self.config.odoo.db,
                odoo_login=self.config.odoo.login,
                odoo_api_key=self.config.odoo.api_key,
                record_url_example=self.config.odoo.record_url_example,
                report_dir=self.config.processing.report_dir,
                backup_dir=self.config.processing.backup_dir,
                write_report_file=self.config.processing.write_report_file,
                stable_backup_name=self.config.processing.stable_backup_name,
                apply=True,
                visible_excel=self.config.processing.visible_excel,
                allow_open_workbook_update=self.config.processing.update_open_workbook,
                excel_session_backend=self.config.processing.excel_session_backend,
                excel_save_debounce_seconds=self.config.processing.excel_save_debounce_seconds,
                allow_live_update_with_autosave=self.config.processing.allow_live_update_with_autosave,
                workbook_slot=self._workbook_slot_for_path(path),
                performance_mode=self.config.processing.performance_mode,
            )
        except PermissionError:
            self._retry_later(path, "Workbook is locked by another program.", "excel_waiting_close")
            return
        except RuntimeError as exc:
            message = str(exc).strip().lower()
            if "read-only" in message or "readonly" in message:
                self._retry_later(path, "Workbook opened as read-only. Waiting before retry.", "excel_read_only")
                return
            if "locked" in message:
                self._retry_later(path, "Workbook appears locked. Waiting before retry.", "excel_waiting_close")
                return
            raise
        except WorkbookAccessError as exc:
            status = exc.access.status
            if status == "open_writable":
                self._retry_later(path, str(exc), "excel_waiting_close")
                return
            if status == "open_read_only":
                self._retry_later(path, str(exc), "excel_read_only")
                return
            if status == "open_autosave":
                self._retry_later(path, str(exc), "excel_autosave_deferred")
                return
            if status == "open_ambiguous_instance":
                self._retry_later(path, str(exc), "excel_ambiguous_instance")
                return
            if status == "unsupported_live_update":
                self._retry_later(path, str(exc), "excel_backend_unavailable")
                return
            raise

        fingerprint_after = file_fingerprint(path)
        self.state.mark_processed(path, fingerprint_after)
        self.state.save()
        self._log_summary(summary)

        linked = summary.status_counts.get("linked", 0)
        fallback_used = summary.status_counts.get("row_fallback_used", 0)
        issue_code = ""
        if summary.status_counts.get("missing_required_header"):
            message = f"{path.name}: missing required header."
            issue_code = "missing_required_header"
        elif linked:
            message = f"{path.name}: linked {linked} order(s)."
            if fallback_used:
                message = f"{message} Row fallback used for {fallback_used} row(s)."
        elif summary.total_cells == 0:
            message = f"{path.name}: no candidate data found."
        else:
            message = f"{path.name}: processed with no automatic links."
        if not issue_code and summary.live_update_used:
            issue_code = "excel_live_updated"
        runtime_status(self.config, "running", message, last_issue_code=issue_code)
        self.tray.show_notification(APP_NAME, message)

    def _retry_later(self, path: Path, reason: str, code: str) -> None:
        due_at = time.time() + self.config.processing.retry_delay_seconds
        with self.pending_lock:
            self.pending[path] = due_at
        self.logger.info("Retry scheduled for %s: %s", path, reason)
        runtime_status(self.config, "running", reason, last_issue_code=code, last_issue_message=path.name)

    def _log_summary(self, summary: WorkbookProcessSummary) -> None:
        counts = ", ".join(f"{key}={value}" for key, value in sorted(summary.status_counts.items())) or "no statuses"
        self.logger.info(
            "Processed %s | cells=%s unique=%s linked=%s statuses=%s state=%s live_update=%s report=%s backup=%s",
            summary.workbook_path,
            summary.total_cells,
            summary.unique_orders,
            summary.linked_count,
            counts,
            summary.workbook_state,
            summary.live_update_used,
            summary.report_path,
            summary.backup_path,
        )


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", default=str(Path(__file__).with_name("config.json")), help="Path to config.json")
    return parser.parse_args(argv)


def acquire_single_instance() -> Any:
    mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        raise SystemExit(f"{APP_NAME} is already running.")
    return mutex


def status_path_from_config_arg(config_path: Path) -> Path:
    try:
        return get_runtime_status_path(config_path)
    except Exception:
        return expand_path(config_path).parent / "runtime_status.json"


def open_setup_ui(config_path: Path) -> None:
    try:
        if getattr(sys, "frozen", False):
            subprocess.Popen(
                [sys.executable, "--config", str(config_path)],
                cwd=str(Path(sys.executable).parent),
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            return
        subprocess.Popen(
            [current_pythonw(), str(Path(__file__).with_name("odoo_excel_agent_ui.py")), "--config", str(config_path)],
            cwd=str(Path(__file__).parent),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        pass


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    config_path = expand_path(args.config)
    status_path = status_path_from_config_arg(config_path)
    write_runtime_status(status_path, "starting", "Starting background watcher.", updated_at=timestamp())
    try:
        config = load_agent_config(config_path)
    except AgentCredentialError as exc:
        write_runtime_status(status_path, "startup_failed", str(exc), updated_at=timestamp(), last_issue_code="missing_api_key")
        open_setup_ui(config_path)
        return 2
    except Exception as exc:
        lowered = str(exc).lower()
        code = (
            "invalid_watch_target"
            if "watch target" in lowered or "watch file" in lowered or "watch folder" in lowered or "achats file" in lowered
            else "invalid_odoo_settings"
            if "invalid odoo settings" in lowered or "odoo database" in lowered or "odoo url" in lowered or "purchase url example" in lowered
            else "startup_failed"
        )
        write_runtime_status(status_path, code, str(exc), updated_at=timestamp(), last_issue_code=code)
        return 2

    logger = configure_logging(config.processing.log_file)
    logger.info("Starting %s", APP_NAME)
    try:
        mutex = acquire_single_instance()
    except SystemExit as exc:
        runtime_status(config, "running", str(exc))
        return 2

    app = OdooExcelAgent(config, logger)
    try:
        # Validate auth before opening the tray loop so startup failures are explicit.
        from link_odoo_vendor_bills import OdooClient

        OdooClient(config.odoo.url, config.odoo.db, config.odoo.login, config.odoo.api_key).authenticate()
        app.start()
        win32gui.PumpMessages()
    except KeyboardInterrupt:
        logger.info("Interrupted by keyboard.")
    except RuntimeError as exc:
        message = str(exc)
        code = "odoo_auth_failed" if "authentication failed" in message.lower() else "startup_failed"
        runtime_status(config, code, message, last_issue_code=code, last_issue_message=message)
        logger.error("Startup/runtime error: %s", message)
        return 2
    except Exception:
        formatted = traceback.format_exc()
        runtime_status(config, "startup_failed", "Unexpected startup failure.", last_issue_code="startup_failed", last_issue_message=formatted)
        logger.error("Fatal startup error:\n%s", formatted)
        return 2
    finally:
        app.stop()
        del mutex
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
