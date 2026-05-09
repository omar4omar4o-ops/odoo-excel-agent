"""Desktop UI for configuring and controlling the Odoo Excel background agent."""

from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import threading
from contextlib import contextmanager
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, StringVar, Tk, filedialog, messagebox, Canvas
from tkinter import ttk
from typing import Any

import customtkinter as ctk

import pythoncom
import win32com.client  # type: ignore[import-not-found]

from link_odoo_vendor_bills import (
    DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS,
    DEFAULT_EXCEL_SESSION_BACKEND,
    DEFAULT_ODOO_URL,
    OdooClient,
    PERFORMANCE_MODE_LIVE,
    PERFORMANCE_MODE_SILENT,
    WORKBOOK_SLOT_ACHATS_ETRANGER,
    WORKBOOK_SLOT_ACHATS_LOCAL,
    WORKBOOK_SLOT_SELLER_PREVIOUS,
    is_supported_workbook,
    process_workbook,
    validate_odoo_settings,
)
from odoo_excel_agent_support import (
    AGENT_SCRIPT,
    APP_NAME,
    APP_VERSION,
    DEFAULT_UPDATE_URL,
    DEFAULT_INSTALL_DIR,
    STARTUP_SHORTCUT,
    AgentCredentialError,
    credential_exists,
    default_config,
    default_watch_folder,
    delete_secret,
    expand_path,
    get_runtime_status_path,
    load_normalized_config,
    make_credential_target,
    normalize_excel_session_backend,
    read_runtime_status,
    read_secret,
    save_normalized_config,
    store_secret,
    WATCH_MODE_ACHATS_PAIR,
    WATCH_MODE_SELECTED_WORKBOOKS,
)
from odoo_excel_updater import (
    UpdateInfo,
    check_for_update,
    download_update_asset,
    prepare_update_payload,
    schedule_update_install,
)


SUPPORT_FILES = [
    "launcher.py",
    "link_odoo_vendor_bills.py",
    "odoo_excel_background.py",
    "odoo_excel_agent_support.py",
    "odoo_excel_agent_ui.py",
    "odoo_excel_updater.py",
    "install_odoo_excel_agent.ps1",
    "launch_odoo_excel_agent_ui.ps1",
    "odoo_excel_config.example.json",
    "uninstall_odoo_excel_agent.ps1",
]

RUNTIME_DEPENDENCIES = {
    "pywin32": "pythoncom",
    "watchdog": "watchdog",
    "openpyxl": "openpyxl",
}


@contextmanager
def com_scope() -> Any:
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def current_pythonw() -> str:
    executable = Path(sys.executable)
    if executable.name.lower() == "pythonw.exe":
        return str(executable)
    candidate = executable.with_name("pythonw.exe")
    if candidate.exists():
        return str(candidate)
    return str(executable)


def missing_runtime_dependencies() -> list[str]:
    missing: list[str] = []
    for package_name, module_name in RUNTIME_DEPENDENCIES.items():
        if importlib.util.find_spec(module_name) is None:
            missing.append(package_name)
    return missing


class UiRuntime:
    @staticmethod
    def list_agent_processes() -> list[dict[str, str | int]]:
        with com_scope():
            locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            service = locator.ConnectServer(".", "root\\cimv2")
            result: list[dict[str, str | int]] = []
            process = None
            query = service.ExecQuery("SELECT ProcessId, Name, CommandLine FROM Win32_Process")
            for process in query:
                command_line = str(getattr(process, "CommandLine", "") or "")
                is_agent = False
                if getattr(sys, "frozen", False):
                    exe_name = Path(sys.executable).name
                    if exe_name in command_line and "--run-background" in command_line:
                        is_agent = True
                else:
                    if AGENT_SCRIPT in command_line:
                        is_agent = True
                        
                if is_agent:
                    result.append(
                        {
                            "pid": int(process.ProcessId),
                            "name": str(process.Name),
                            "command_line": command_line,
                        }
                    )
            if process is not None:
                del process
            del query
            del service
            del locator
            return result

    @staticmethod
    def stop_agent_processes() -> int:
        processes = UiRuntime.list_agent_processes()
        count = 0
        for process in processes:
            pid = int(process["pid"])
            completed = subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                text=True,
                check=False,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if completed.returncode == 0:
                count += 1
        return count

    @staticmethod
    def create_startup_shortcut(target_pythonw: str, install_dir: Path, config_path: Path) -> None:
        with com_scope():
            install_dir.mkdir(parents=True, exist_ok=True)
            STARTUP_SHORTCUT.parent.mkdir(parents=True, exist_ok=True)
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(str(STARTUP_SHORTCUT))
            if getattr(sys, "frozen", False):
                shortcut.TargetPath = sys.executable
                shortcut.Arguments = f'--run-background --config "{config_path}"'
                shortcut.WorkingDirectory = str(Path(sys.executable).parent)
                shortcut.IconLocation = f"{sys.executable},0"
            else:
                shortcut.TargetPath = target_pythonw
                shortcut.Arguments = f'"{install_dir / AGENT_SCRIPT}" --config "{config_path}"'
                shortcut.WorkingDirectory = str(install_dir)
                shortcut.IconLocation = f"{target_pythonw},0"
            shortcut.Save()
            del shortcut
            del shell

    @staticmethod
    def start_agent_process(target_pythonw: str, install_dir: Path, config_path: Path) -> None:
        if getattr(sys, "frozen", False):
            subprocess.Popen(
                [sys.executable, "--run-background", "--config", str(config_path)],
                cwd=str(Path(sys.executable).parent),
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        else:
            subprocess.Popen(
                [target_pythonw, str(install_dir / AGENT_SCRIPT), "--config", str(config_path)],
                cwd=install_dir,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )


class AgentControlApp:
    def __init__(self, root: ctk.CTk, config_arg: str = "") -> None:
        self.root = root
        self.root.title("Odoo Excel Agent Control Center")

        self.screen_width = root.winfo_screenwidth()
        self.screen_height = root.winfo_screenheight()
        self.compact_layout = self.screen_height < 840
        window_width = min(1260, max(1040, self.screen_width - 30))
        window_height = min(900, max(660, self.screen_height - 40))
        self.tab_view_height = max(300, window_height - (420 if self.compact_layout else 360))

        self.root.geometry(f"{window_width}x{window_height}")
        self.root.minsize(840, 560)
        self.palette = {
            "bg": "#0f1117",
            "card": "#1a1d28",
            "card_alt": "#1e2233",
            "card_soft": "#161a24",
            "border": "#2a2e3d",
            "ink": "#e2e8f0",
            "muted": "#8892a4",
            "accent": "#0ea5e9",
            "accent_deep": "#0284c7",
            "accent_soft": "#0c3449",
            "success": "#22c55e",
            "warning": "#f59e0b",
            "danger": "#ef4444",
            "console": "#0a0c10",
            "console_text": "#a5f3c4",
            "hero_bg": "#111827",
            "sidebar": "#111520",
            "hover": "#252a3a",
        }
        self.root.configure(fg_color=self.palette["bg"])

        self.script_dir = Path(__file__).resolve().parent
        self.config_path = self._initial_config_path(config_arg)
        self.config: dict[str, Any] = default_config(self.config_path.parent)
        self.credential_target = ""
        self.api_key_is_stored = False
        self.activity_count = 0
        self._active_tab_canvas: Canvas | None = None

        self.install_dir_var = StringVar(value=str(DEFAULT_INSTALL_DIR))
        self.odoo_url_var = StringVar(value=DEFAULT_ODOO_URL)
        self.odoo_db_var = StringVar()
        self.odoo_login_var = StringVar()
        self.odoo_api_key_var = StringVar()
        self.record_url_var = StringVar()
        self.credential_state_var = StringVar(value="No API key stored")

        self.manual_file_var = StringVar()
        self.watch_mode_var = StringVar(value=WATCH_MODE_SELECTED_WORKBOOKS)
        self.achats_local_file_var = StringVar()
        self.achats_etranger_file_var = StringVar()
        self.seller_previous_file_var = StringVar()
        self.watch_file_var = StringVar()
        self.watch_folder_var = StringVar(value=default_watch_folder())
        self.backup_dir_var = StringVar()
        self.status_var = StringVar(value="Stopped")
        self.status_detail_var = StringVar(value="Background agent is not running yet.")
        self.watch_summary_var = StringVar(value="No watch target selected")
        self.watch_detail_var = StringVar(value="Choose a workbook or folder to monitor.")
        self.dependency_summary_var = StringVar(value="Checking runtime packages")
        self.dependency_detail_var = StringVar(value="pywin32, watchdog")
        self.config_summary_var = StringVar(value="Configuration file location will appear here.")
        self.last_activity_var = StringVar(value="Activity console is ready.")
        self.activity_count_var = StringVar(value="0 events logged")
        self.update_manifest_url_var = StringVar(value=DEFAULT_UPDATE_URL)
        self.update_state_var = StringVar(value=f"Current version: {APP_VERSION}")
        self.update_detail_var = StringVar(value="Updates are configured automatically from GitHub Releases.")
        self._last_update_info: UpdateInfo | None = None

        self.process_existing_var = StringVar(value="0")
        self.recursive_var = StringVar(value="0")
        self.performance_mode_var = StringVar(value=PERFORMANCE_MODE_SILENT)
        self.update_open_workbook_var = StringVar(value="0")
        self.excel_event_monitoring_var = StringVar(value="0")
        self.excel_session_backend_var = StringVar(value=DEFAULT_EXCEL_SESSION_BACKEND)
        self.excel_save_debounce_var = StringVar(value=str(DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS))
        self.allow_live_update_with_autosave_var = StringVar(value="0")
        self.visible_excel_var = StringVar(value="0")
        self.write_report_var = StringVar(value="0")
        self.stable_backup_var = StringVar(value="1")
        self.settle_seconds_var = StringVar(value="15")
        self.retry_seconds_var = StringVar(value="45")

        self._build_ui()
        self.watch_mode_var.trace_add("write", lambda *_: self._sync_watch_mode_ui())
        self.performance_mode_var.trace_add("write", lambda *_: self._sync_performance_mode_ui())
        self._load_config()
        self._refresh_overview_cards()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._schedule_status_refresh()

    def _initial_config_path(self, config_arg: str) -> Path:
        if config_arg:
            return expand_path(config_arg)
        default_candidate = DEFAULT_INSTALL_DIR / "config.json"
        local_candidate = self.script_dir / "config.json"
        if default_candidate.exists():
            return default_candidate
        if local_candidate.exists():
            return local_candidate
        return default_candidate

    def _build_ui(self) -> None:
        ctk.set_appearance_mode("dark")
        p = self.palette
        bfs = 13 if self.compact_layout else 14

        # Main scrollable frame
        self.main_frame = ctk.CTkScrollableFrame(self.root, fg_color=p["bg"], corner_radius=0)
        self.main_frame.pack(fill=BOTH, expand=True, padx=0, pady=0)
        self.main_frame.columnconfigure(0, weight=1)

        # Hero header
        hero = ctk.CTkFrame(self.main_frame, fg_color=p["hero_bg"], corner_radius=12, border_width=1, border_color=p["border"])
        hero.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        hero.columnconfigure(0, weight=1)

        hero_inner = ctk.CTkFrame(hero, fg_color="transparent")
        hero_inner.pack(fill="x", padx=16, pady=(14, 8))
        hero_inner.columnconfigure(0, weight=1)

        title_frame = ctk.CTkFrame(hero_inner, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(title_frame, text="Odoo Excel Agent", font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"), text_color="#ffffff").pack(anchor="w")
        ctk.CTkLabel(title_frame, text="Silent workbook monitoring - Odoo validation - Background automation", font=ctk.CTkFont(size=12), text_color=p["muted"]).pack(anchor="w", pady=(2, 0))

        ctrl_frame = ctk.CTkFrame(hero_inner, fg_color="transparent")
        ctrl_frame.grid(row=0, column=1, sticky="e")
        self.status_chip_label = ctk.CTkLabel(ctrl_frame, textvariable=self.status_var, font=ctk.CTkFont(size=12, weight="bold"), text_color="#ffffff", fg_color="#4b5563", corner_radius=8, width=100, height=30)
        self.status_chip_label.pack(side=LEFT, padx=(0, 10))
        ctk.CTkButton(ctrl_frame, text="Refresh", width=90, height=32, fg_color="transparent", border_width=1, border_color=p["border"], hover_color=p["hover"], text_color=p["ink"], command=self.refresh_status).pack(side=LEFT, padx=(0, 6))
        ctk.CTkButton(ctrl_frame, text="Save", width=90, height=32, fg_color=p["accent"], hover_color=p["accent_deep"], text_color="#ffffff", command=self.save_config).pack(side=LEFT)

        # Summary cards
        cards_row = ctk.CTkFrame(hero, fg_color="transparent")
        cards_row.pack(fill="x", padx=14, pady=(4, 12))
        for i in range(3):
            cards_row.columnconfigure(i, weight=1)

        self._build_summary_card(cards_row, 0, "Agent State", self.status_var, self.status_detail_var, style_name="SummaryCard.TFrame")
        self._build_summary_card(cards_row, 1, "Watch Target", self.watch_summary_var, self.watch_detail_var, style_name="SummaryCardAlt.TFrame")
        self._build_summary_card(cards_row, 2, "Runtime Packages", self.dependency_summary_var, self.dependency_detail_var, style_name="SummaryCard.TFrame")

        # Config path strip
        config_strip = ctk.CTkFrame(self.main_frame, fg_color="transparent", height=30)
        config_strip.grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 2))
        ctk.CTkLabel(config_strip, text="Config:", font=ctk.CTkFont(size=11), text_color=p["muted"]).pack(side=LEFT)
        ctk.CTkLabel(config_strip, textvariable=self.config_summary_var, font=ctk.CTkFont(size=11), text_color=p["ink"]).pack(side=LEFT, padx=(6, 0))

        # Tabview
        self.tabview = ctk.CTkTabview(self.main_frame, fg_color=p["card"], corner_radius=12, border_width=1, border_color=p["border"], segmented_button_fg_color=p["card_alt"], segmented_button_selected_color=p["accent"], segmented_button_selected_hover_color=p["accent_deep"], segmented_button_unselected_color=p["card_alt"], segmented_button_unselected_hover_color=p["hover"], height=self.tab_view_height)
        self.tabview.grid(row=2, column=0, sticky="nsew", padx=12, pady=(4, 12))
        self.main_frame.rowconfigure(2, weight=1)

        setup_tab = self.tabview.add("Connect")
        process_tab = self.tabview.add("Run Once")
        watch_tab = self.tabview.add("Auto Watch")
        update_tab = self.tabview.add("Update")
        activity_tab = self.tabview.add("Activity")

        for tab in [setup_tab, process_tab, watch_tab, update_tab, activity_tab]:
            tab.columnconfigure(0, weight=1)

        self._build_setup_tab(setup_tab)
        self._build_process_tab(process_tab)
        self._build_watch_tab(watch_tab)
        self._build_update_tab(update_tab)
        self._build_activity_tab(activity_tab)

    def _set_active_tab_canvas(self, canvas: Canvas | None) -> None:
        self._active_tab_canvas = canvas

    def _on_tab_changed(self, event: Any) -> None:
        pass

    def _build_summary_card(self, parent: Any, column: int, title: str, value_var: StringVar, detail_var: StringVar, *, style_name: str) -> None:
        p = self.palette
        bg = p["card_alt"] if "Alt" in style_name else p["card"]
        card = ctk.CTkFrame(parent, fg_color=bg, corner_radius=10, border_width=1, border_color=p["border"])
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 6, 0))
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=10), text_color=p["muted"]).pack(anchor="w", padx=12, pady=(10, 0))
        ctk.CTkLabel(card, textvariable=value_var, font=ctk.CTkFont(size=14, weight="bold"), text_color=p["ink"]).pack(anchor="w", padx=12, pady=(2, 0))
        ctk.CTkLabel(card, textvariable=detail_var, font=ctk.CTkFont(size=10), text_color=p["muted"], wraplength=260, justify="left").pack(anchor="w", padx=12, pady=(0, 10))

    def _create_card(self, parent: Any, *, row: int, column: int, columnspan: int = 1, style_name: str = "Card.TFrame", padding: tuple[int, int, int, int] = (12, 12, 12, 10), padx: tuple[int, int] = (0, 0), pady: tuple[int, int] = (0, 8)) -> ctk.CTkFrame:
        p = self.palette
        bg = p["card_soft"] if "Alt" in style_name else p["card"]
        frame = ctk.CTkFrame(parent, fg_color=bg, corner_radius=10, border_width=1, border_color=p["border"])
        frame.grid(row=row, column=column, columnspan=columnspan, sticky="nsew", padx=padx, pady=pady)
        return frame

    def _add_section_header(self, parent: Any, title: str, subtitle: str, *, alt: bool = False) -> None:
        p = self.palette
        ctk.CTkLabel(parent, text=title, font=ctk.CTkFont(size=15, weight="bold"), text_color=p["ink"]).pack(anchor="w", padx=14, pady=(12, 0))
        ctk.CTkLabel(parent, text=subtitle, font=ctk.CTkFont(size=11), text_color=p["muted"], wraplength=520, justify="left").pack(anchor="w", padx=14, pady=(2, 8))

    def _build_setup_tab(self, parent: Any) -> None:
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=2)
        p = self.palette

        connection = self._create_card(parent, row=0, column=0, padx=(0, 10))
        self._add_section_header(connection, "Odoo Connection", "Secure the Odoo endpoint once. The agent reuses these credentials for live workbook linking in Achats.")
        self._add_entry_row(
            connection,
            "Odoo URL",
            self.odoo_url_var,
            placeholder_text="https://sphe.cloudoo.ma",
        )
        self._add_entry_row(
            connection,
            "Odoo database",
            self.odoo_db_var,
            placeholder_text="sphe.cloudoo.ma",
        )
        self._add_entry_row(
            connection,
            "Odoo login",
            self.odoo_login_var,
            placeholder_text="user@example.com",
        )
        self._add_entry_row(connection, "New Odoo API key", self.odoo_api_key_var, show="*")
        self._add_entry_row(
            connection,
            "Purchase URL example (optional)",
            self.record_url_var,
            placeholder_text="https://sphe.cloudoo.ma/odoo/purchase/1",
        )

        system = self._create_card(parent, row=0, column=1, style_name="CardAlt.TFrame")
        self._add_section_header(system, "Runtime Paths", "Keep installs, backups, and config organized.", alt=True)
        self._add_path_row(system, "Install folder", self.install_dir_var, self.choose_install_dir, alt=True)
        self._add_path_row(system, "Backup folder", self.backup_dir_var, self.choose_backup_dir, alt=True)

        credential = ctk.CTkFrame(system, fg_color="transparent")
        credential.pack(fill="x", padx=14, pady=(10, 10))
        ctk.CTkLabel(credential, text="Stored key", font=ctk.CTkFont(size=12, weight="bold"), text_color=p["ink"]).pack(anchor="w")
        ctk.CTkLabel(credential, textvariable=self.credential_state_var, font=ctk.CTkFont(size=11), text_color=p["muted"], wraplength=360, justify="left").pack(anchor="w", pady=(4, 8))
        ctk.CTkButton(credential, text="Clear Stored Key", width=140, height=30, fg_color=p["danger"], hover_color="#dc2626", text_color="#ffffff", command=self.clear_stored_api_key).pack(anchor="w")

        actions = self._create_card(parent, row=1, column=0, columnspan=2, padding=(18, 16, 18, 16))
        actions.columnconfigure(0, weight=1)
        self._add_section_header(actions, "Deploy", "Validate credentials, save config, then install or refresh the startup agent.")
        action_bar = ctk.CTkFrame(actions, fg_color="transparent")
        action_bar.pack(fill="x", padx=14, pady=(0, 10))
        ctk.CTkButton(action_bar, text="Test Odoo", width=120, height=34, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.test_odoo).pack(side=LEFT)
        ctk.CTkButton(action_bar, text="Save Config", width=120, height=34, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.save_config).pack(side=LEFT, padx=(8, 0))
        ctk.CTkButton(action_bar, text="Install / Update Agent", width=200, height=34, fg_color=p["accent"], hover_color=p["accent_deep"], text_color="#ffffff", command=self.install_agent).pack(side=RIGHT)

    def _build_process_tab(self, parent: Any) -> None:
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=2)
        p = self.palette

        run_card = self._create_card(parent, row=0, column=0, padx=(0, 10))
        self._add_section_header(run_card, "Process A Single Workbook", "One-off pass. Silent mode updates closed workbooks without opening Excel.")
        self._add_file_row(run_card, "Excel file", self.manual_file_var, self.choose_manual_file)
        actions = ctk.CTkFrame(run_card, fg_color="transparent")
        actions.pack(fill="x", padx=14, pady=(10, 14))
        ctk.CTkButton(actions, text="Process File", width=140, height=34, fg_color=p["accent"], hover_color=p["accent_deep"], text_color="#ffffff", command=self.process_selected_file).pack(side=LEFT)
        ctk.CTkButton(actions, text="Open File", width=100, height=34, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.open_selected_file).pack(side=LEFT, padx=(8, 0))
        ctk.CTkButton(actions, text="Open Backups", width=120, height=34, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.open_backups).pack(side=LEFT, padx=(8, 0))

        presets = ctk.CTkFrame(run_card, fg_color="transparent")
        presets.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkLabel(
            presets,
            text="Selected workbook shortcuts",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=p["muted"],
        ).pack(anchor="w", pady=(0, 6))
        self._add_file_row(presets, "ACHATS LOCAL", self.achats_local_file_var, self.choose_achats_local_file)
        local_actions = ctk.CTkFrame(presets, fg_color="transparent")
        local_actions.pack(fill="x", pady=(6, 8))
        ctk.CTkButton(
            local_actions,
            text="Process ACHATS LOCAL",
            width=180,
            height=32,
            fg_color=p["card_alt"],
            hover_color=p["hover"],
            text_color=p["ink"],
            border_width=1,
            border_color=p["border"],
            command=self.process_achats_local_file,
        ).pack(side=LEFT)
        ctk.CTkButton(
            local_actions,
            text="Open File",
            width=100,
            height=32,
            fg_color=p["card_alt"],
            hover_color=p["hover"],
            text_color=p["ink"],
            border_width=1,
            border_color=p["border"],
            command=lambda: self._open_configured_workbook(self.achats_local_file_var),
        ).pack(side=LEFT, padx=(8, 0))
        self._add_file_row(presets, "ACHATS ETRANGER", self.achats_etranger_file_var, self.choose_achats_etranger_file)
        etranger_actions = ctk.CTkFrame(presets, fg_color="transparent")
        etranger_actions.pack(fill="x", pady=(6, 0))
        ctk.CTkButton(
            etranger_actions,
            text="Process ACHATS ETRANGER",
            width=200,
            height=32,
            fg_color=p["card_alt"],
            hover_color=p["hover"],
            text_color=p["ink"],
            border_width=1,
            border_color=p["border"],
            command=self.process_achats_etranger_file,
        ).pack(side=LEFT)
        ctk.CTkButton(
            etranger_actions,
            text="Open File",
            width=100,
            height=32,
            fg_color=p["card_alt"],
            hover_color=p["hover"],
            text_color=p["ink"],
            border_width=1,
            border_color=p["border"],
            command=lambda: self._open_configured_workbook(self.achats_etranger_file_var),
        ).pack(side=LEFT, padx=(8, 0))
        self._add_file_row(presets, "Seller / Previous", self.seller_previous_file_var, self.choose_seller_previous_file)
        seller_actions = ctk.CTkFrame(presets, fg_color="transparent")
        seller_actions.pack(fill="x", pady=(6, 0))
        ctk.CTkButton(
            seller_actions,
            text="Process Seller File",
            width=180,
            height=32,
            fg_color=p["card_alt"],
            hover_color=p["hover"],
            text_color=p["ink"],
            border_width=1,
            border_color=p["border"],
            command=self.process_seller_previous_file,
        ).pack(side=LEFT)
        ctk.CTkButton(
            seller_actions,
            text="Open File",
            width=100,
            height=32,
            fg_color=p["card_alt"],
            hover_color=p["hover"],
            text_color=p["ink"],
            border_width=1,
            border_color=p["border"],
            command=lambda: self._open_configured_workbook(self.seller_previous_file_var),
        ).pack(side=LEFT, padx=(8, 0))

        guide = self._create_card(parent, row=0, column=1, style_name="CardAlt.TFrame")
        self._add_section_header(guide, "What Happens", "The one-click run keeps your workflow safe.", alt=True)
        guide_points = (
            "Reads workbook-specific columns (N\u00b0FACTURE, N COMMANDE, or legacy headers).",
            "ACHATS LOCAL tries N\u00b0FACTURE first, then N commandes if the first value is not found.",
            "If ACHATS LOCAL is not found in purchase orders, it searches accessible Odoo records globally.",
            "ACHATS ETRANGER searches Odoo from N COMMANDE, not the amount.",
            "Uses workbook-specific lookup rules (Reference commande with partner_ref fallback, or legacy partner_ref).",
            "Silent mode waits for open workbooks to close before writing.",
            "Creates a backup before writing hyperlinks.",
        )
        for text in guide_points:
            ctk.CTkLabel(guide, text=text, font=ctk.CTkFont(size=11), text_color=p["muted"], wraplength=380, justify="left").pack(anchor="w", padx=14, pady=3)

    def _build_watch_tab(self, parent: Any) -> None:
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=2)
        p = self.palette

        target = self._create_card(parent, row=0, column=0, padx=(0, 10))
        self._add_section_header(
            target,
            "Selected Background Targets",
            "Choose the three exact workbook targets that the background agent should monitor.",
        )
        self._add_file_row(target, "ACHATS LOCAL", self.achats_local_file_var, self.choose_achats_local_file)
        self._add_file_row(target, "ACHATS ETRANGER", self.achats_etranger_file_var, self.choose_achats_etranger_file)
        self._add_file_row(target, "Seller / Previous", self.seller_previous_file_var, self.choose_seller_previous_file)
        self._add_entry_row(target, "Settle seconds", self.settle_seconds_var)
        self._add_entry_row(target, "Retry delay seconds", self.retry_seconds_var)

        behavior = self._create_card(parent, row=0, column=1, style_name="CardAlt.TFrame")
        self._add_section_header(behavior, "Behavior", "Startup scanning, live updates, visibility, reporting, and backup strategy.", alt=True)
        mode_row = ctk.CTkFrame(behavior, fg_color="transparent")
        mode_row.pack(fill="x", padx=14, pady=6)
        ctk.CTkLabel(mode_row, text="Performance mode", font=ctk.CTkFont(size=12, weight="bold"), text_color=p["ink"]).pack(side=LEFT, padx=(0, 14))
        self.performance_mode_menu = ctk.CTkOptionMenu(
            mode_row,
            variable=self.performance_mode_var,
            values=[PERFORMANCE_MODE_SILENT, PERFORMANCE_MODE_LIVE],
            fg_color=p["card_alt"],
            button_color=p["accent"],
            button_hover_color=p["accent_deep"],
            dropdown_fg_color=p["card"],
            dropdown_hover_color=p["hover"],
            text_color=p["ink"],
        )
        self.performance_mode_menu.pack(side=LEFT, fill="x", expand=True)
        self._add_checkbutton(behavior, "Process existing files on startup", self.process_existing_var, alt=True)
        self.update_open_workbook_check = self._add_checkbutton(behavior, "Update open workbook after save", self.update_open_workbook_var, alt=True)
        self._add_checkbutton(behavior, "Show Excel while processing", self.visible_excel_var, alt=True)
        self._add_checkbutton(behavior, "Write CSV reports", self.write_report_var, alt=True)
        self._add_checkbutton(behavior, "Keep stable original backup", self.stable_backup_var, alt=True)

        monitoring = self._create_card(parent, row=1, column=0, padx=(0, 10))
        self._add_section_header(monitoring, "Advanced Live Excel Monitoring", "Fine-tune how the agent listens to Excel save and close events.")
        self.excel_event_monitoring_check = self._add_checkbutton(monitoring, "Monitor Excel save/close events", self.excel_event_monitoring_var)
        self.allow_live_autosave_check = self._add_checkbutton(monitoring, "Allow live update with AutoSave", self.allow_live_update_with_autosave_var)
        self.excel_save_debounce_entry = self._add_entry_row(monitoring, "Excel save debounce", self.excel_save_debounce_var)
        backend_row = ctk.CTkFrame(monitoring, fg_color="transparent")
        backend_row.pack(fill="x", padx=14, pady=6)
        ctk.CTkLabel(backend_row, text="Live-session backend", font=ctk.CTkFont(size=12, weight="bold"), text_color=p["ink"]).pack(side=LEFT, padx=(0, 14))
        self.excel_backend_menu = ctk.CTkOptionMenu(
            backend_row,
            variable=self.excel_session_backend_var,
            values=[DEFAULT_EXCEL_SESSION_BACKEND],
            fg_color=p["card_alt"],
            button_color=p["accent"],
            button_hover_color=p["accent_deep"],
            dropdown_fg_color=p["card"],
            dropdown_hover_color=p["hover"],
            text_color=p["ink"],
        )
        self.excel_backend_menu.pack(side=LEFT, fill="x", expand=True)

        ctrl_actions = self._create_card(parent, row=1, column=1, style_name="CardAlt.TFrame")
        self._add_section_header(ctrl_actions, "Control Surface", "Start, stop, and inspect the agent.", alt=True)
        action_bar = ctk.CTkFrame(ctrl_actions, fg_color="transparent")
        action_bar.pack(fill="x", padx=14, pady=(0, 10))
        ctk.CTkButton(action_bar, text="Start Agent", width=120, height=34, fg_color=p["success"], hover_color="#16a34a", text_color="#ffffff", command=self.start_agent).pack(side=LEFT)
        ctk.CTkButton(action_bar, text="Stop", width=80, height=34, fg_color=p["danger"], hover_color="#dc2626", text_color="#ffffff", command=self.stop_agent).pack(side=LEFT, padx=(8, 0))
        action_bar2 = ctk.CTkFrame(ctrl_actions, fg_color="transparent")
        action_bar2.pack(fill="x", padx=14, pady=(0, 10))
        ctk.CTkButton(action_bar2, text="Open Watch Folder", width=150, height=30, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.open_watch_target).pack(side=LEFT)
        ctk.CTkButton(action_bar2, text="Open Install Folder", width=140, height=30, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.open_install_folder).pack(side=LEFT, padx=(8, 0))

    def _build_update_tab(self, parent: Any) -> None:
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=2)
        p = self.palette

        update_card = self._create_card(parent, row=0, column=0, padx=(0, 10))
        self._add_section_header(
            update_card,
            "Free Updates",
            "Updates are automatic from the official GitHub Release. Advanced users can override the URL.",
        )
        self._add_entry_row(
            update_card,
            "Update URL",
            self.update_manifest_url_var,
            placeholder_text="https://api.github.com/repos/USER/REPO/releases/latest",
        )
        status_box = ctk.CTkFrame(update_card, fg_color=p["card_alt"], corner_radius=8, border_width=1, border_color=p["border"])
        status_box.pack(fill="x", padx=14, pady=(10, 10))
        ctk.CTkLabel(status_box, textvariable=self.update_state_var, font=ctk.CTkFont(size=13, weight="bold"), text_color=p["ink"]).pack(anchor="w", padx=12, pady=(10, 2))
        ctk.CTkLabel(status_box, textvariable=self.update_detail_var, font=ctk.CTkFont(size=11), text_color=p["muted"], wraplength=660, justify="left").pack(anchor="w", padx=12, pady=(0, 10))

        buttons = ctk.CTkFrame(update_card, fg_color="transparent")
        buttons.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkButton(buttons, text="Check For Updates", width=150, height=34, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.check_for_updates).pack(side=LEFT)
        ctk.CTkButton(buttons, text="Update Now", width=140, height=34, fg_color=p["accent"], hover_color=p["accent_deep"], text_color="#ffffff", command=self.install_update).pack(side=LEFT, padx=(8, 0))

        guide = self._create_card(parent, row=0, column=1, style_name="CardAlt.TFrame")
        self._add_section_header(guide, "Safe Update Rules", "The updater is intentionally simple and verifiable.", alt=True)
        points = (
            f"Current version: {APP_VERSION}",
            "Downloads only a ZIP/EXE from the manifest URL.",
            "Verifies SHA-256 when the manifest provides it.",
            "Stops the background agent before replacing files.",
            "Keeps config and Windows Credential Manager credentials unchanged.",
            "Best free hosting: GitHub Releases with update-manifest.json.",
        )
        for text in points:
            ctk.CTkLabel(guide, text=text, font=ctk.CTkFont(size=11), text_color=p["muted"], wraplength=380, justify="left").pack(anchor="w", padx=14, pady=3)

    def _build_activity_tab(self, parent: Any) -> None:
        parent.columnconfigure(0, weight=1)
        p = self.palette

        top = self._create_card(parent, row=0, column=0, padding=(18, 16, 18, 16))
        top.columnconfigure(0, weight=1)
        self._add_section_header(top, "Runtime Activity", "Review the latest actions and confirm what the background agent is doing.")
        top_row = ctk.CTkFrame(top, fg_color="transparent")
        top_row.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkButton(top_row, text="Open Log", width=100, height=30, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.open_log).pack(side=LEFT)
        ctk.CTkButton(top_row, text="Open Backups", width=110, height=30, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.open_backups).pack(side=LEFT, padx=(6, 0))
        ctk.CTkButton(top_row, text="Open Config", width=110, height=30, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=self.open_config_file).pack(side=LEFT, padx=(6, 0))
        ctk.CTkButton(top_row, text="Clear", width=80, height=30, fg_color=p["danger"], hover_color="#dc2626", text_color="#ffffff", command=self.clear_status).pack(side=RIGHT)

        meta = ctk.CTkFrame(top, fg_color="transparent")
        meta.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkLabel(meta, textvariable=self.last_activity_var, font=ctk.CTkFont(size=12, weight="bold"), text_color=p["ink"]).pack(anchor="w")
        ctk.CTkLabel(meta, textvariable=self.activity_count_var, font=ctk.CTkFont(size=10), text_color=p["muted"]).pack(anchor="w", pady=(2, 0))

        console_card = self._create_card(parent, row=1, column=0, padding=(14, 14, 14, 14))
        self.status_text = ctk.CTkTextbox(
            console_card,
            height=300 if not self.compact_layout else 220,
            fg_color=p["console"],
            text_color=p["console_text"],
            font=ctk.CTkFont(family="Cascadia Code", size=12),
            corner_radius=8,
            border_width=1,
            border_color=p["border"],
        )
        self.status_text.pack(fill=BOTH, expand=True, padx=8, pady=8)
        self.status_text.insert("1.0", "Ready.\n")
        self.status_text.configure(state="disabled")

    def _add_path_row(self, parent: Any, label: str, var: StringVar, browse_command: Any, *, alt: bool = False) -> tuple[ctk.CTkEntry, ctk.CTkButton]:
        p = self.palette
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=12, weight="bold"), text_color=p["ink"], width=120).pack(side=LEFT, padx=(0, 10))
        entry = ctk.CTkEntry(row, textvariable=var, fg_color=p["card_alt"], border_color=p["border"], text_color=p["ink"], placeholder_text_color=p["muted"], height=32)
        entry.pack(side=LEFT, fill="x", expand=True)
        button = ctk.CTkButton(row, text="Browse", width=80, height=32, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=browse_command)
        button.pack(side=LEFT, padx=(8, 0))
        return entry, button

    def _add_file_row(self, parent: Any, label: str, var: StringVar, browse_command: Any, *, alt: bool = False) -> tuple[ctk.CTkEntry, ctk.CTkButton]:
        p = self.palette
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=12, weight="bold"), text_color=p["ink"], width=120).pack(side=LEFT, padx=(0, 10))
        entry = ctk.CTkEntry(row, textvariable=var, fg_color=p["card_alt"], border_color=p["border"], text_color=p["ink"], placeholder_text_color=p["muted"], height=32)
        entry.pack(side=LEFT, fill="x", expand=True)
        button = ctk.CTkButton(row, text="Choose File", width=100, height=32, fg_color=p["card_alt"], hover_color=p["hover"], text_color=p["ink"], border_width=1, border_color=p["border"], command=browse_command)
        button.pack(side=LEFT, padx=(8, 0))
        return entry, button

    def _add_entry_row(
        self,
        parent: Any,
        label: str,
        var: StringVar,
        show: str | None = None,
        *,
        alt: bool = False,
        placeholder_text: str = "",
    ) -> ctk.CTkEntry:
        p = self.palette
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=12, weight="bold"), text_color=p["ink"], width=160).pack(side=LEFT, padx=(0, 10))
        entry = ctk.CTkEntry(
            row,
            textvariable=var,
            fg_color=p["card_alt"],
            border_color=p["border"],
            text_color=p["ink"],
            placeholder_text_color=p["muted"],
            placeholder_text=placeholder_text,
            height=32,
        )
        if show:
            entry.configure(show=show)
        entry.pack(side=LEFT, fill="x", expand=True)
        return entry

    def _add_checkbutton(self, parent: Any, text: str, var: StringVar, *, alt: bool = False) -> ctk.CTkCheckBox:
        p = self.palette
        control = ctk.CTkCheckBox(parent, text=text, variable=var, onvalue="1", offvalue="0", fg_color=p["accent"], hover_color=p["accent_deep"], text_color=p["ink"], border_color=p["border"])
        control.pack(anchor="w", padx=14, pady=3)
        return control

    def append_status(self, message: str) -> None:
        def _append() -> None:
            timestamp = dt.datetime.now().strftime("%H:%M:%S")
            line = f"[{timestamp}] {message.rstrip()}"
            self.activity_count += 1
            self.last_activity_var.set(message.rstrip())
            self.activity_count_var.set(f"{self.activity_count} events logged")
            self.status_text.configure(state="normal")
            self.status_text.insert(END, line + "\n")
            self.status_text.see(END)
            self.status_text.configure(state="disabled")

        self.root.after(0, _append)

    def clear_status(self) -> None:
        self.activity_count = 0
        self.last_activity_var.set("Activity console cleared.")
        self.activity_count_var.set("0 events logged")
        self.status_text.configure(state="normal")
        self.status_text.delete("1.0", END)
        self.status_text.configure(state="disabled")

    def _schedule_status_refresh(self) -> None:
        self.refresh_status()
        self.root.after(5000, self._schedule_status_refresh)

    def choose_install_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.install_dir_var.get() or str(DEFAULT_INSTALL_DIR))
        if selected:
            self.install_dir_var.set(selected)
            self._refresh_overview_cards()

    def choose_backup_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.backup_dir_var.get() or str(self._install_dir() / "backups"))
        if selected:
            self.backup_dir_var.set(selected)
            self._refresh_overview_cards()

    def choose_achats_local_file(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path.home()),
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if selected:
            self.achats_local_file_var.set(selected)
            self._refresh_overview_cards()

    def choose_achats_etranger_file(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path.home()),
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if selected:
            self.achats_etranger_file_var.set(selected)
            self._refresh_overview_cards()

    def choose_seller_previous_file(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path.home()),
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if selected:
            self.seller_previous_file_var.set(selected)
            self._refresh_overview_cards()

    def choose_manual_file(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=str(Path.home()),
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if selected:
            self.manual_file_var.set(selected)

    def _compact_path(self, raw_path: str, *, keep_segments: int = 3) -> str:
        text = str(raw_path or "").strip()
        if not text:
            return ""
        path = Path(text)
        parts = list(path.parts)
        if len(parts) <= keep_segments + 1:
            return text
        return str(Path(parts[0], "...", *parts[-keep_segments:]))

    def _configured_selected_workbooks(self) -> list[str]:
        paths: list[str] = []
        for raw_value in (
            self._normalized_achats_local_file(),
            self._normalized_achats_etranger_file(),
            self._normalized_seller_previous_file(),
        ):
            if raw_value:
                paths.append(raw_value)
        return paths

    def _watch_target_status_lines(self) -> list[str]:
        lines: list[str] = []
        for label, raw_path, required_header in (
            ("ACHATS LOCAL", self._normalized_achats_local_file(), "N\u00b0FACTURE or N commandes"),
            ("ACHATS ETRANGER", self._normalized_achats_etranger_file(), "N COMMANDE"),
            ("Seller / Previous", self._normalized_seller_previous_file(), "legacy seller headers"),
        ):
            if not raw_path:
                lines.append(f"{label}: not configured")
                continue
            path = Path(raw_path)
            if path.exists() and path.is_file():
                lines.append(f"{label}: ready for background watch")
            elif path.parent.exists() and path.parent.is_dir():
                lines.append(f"{label}: missing file ({required_header})")
            else:
                lines.append(f"{label}: invalid folder")
        return lines

    def _refresh_overview_cards(self, runtime: Any | None = None, status_text: str | None = None) -> None:
        configured = self._configured_selected_workbooks()
        target_value = f"{len(configured)}/3 selected workbooks"
        detail_lines = self._watch_target_status_lines()
        if configured:
            detail_lines.extend(self._compact_path(path, keep_segments=4) for path in configured)
            target_detail = "\n".join(detail_lines)
        else:
            target_detail = "Choose ACHATS LOCAL, ACHATS ETRANGER, and/or the seller workbook to monitor."
        self.watch_summary_var.set(target_value)
        self.watch_detail_var.set(target_detail)

        missing = missing_runtime_dependencies()
        if missing:
            self.dependency_summary_var.set("Missing packages")
            self.dependency_detail_var.set(", ".join(sorted(missing)))
        else:
            self.dependency_summary_var.set("All packages ready")
            self.dependency_detail_var.set("pywin32, watchdog, openpyxl detected")

        self.config_summary_var.set(self._compact_path(str(self._config_file_path()), keep_segments=5))
        if status_text is not None:
            self.status_var.set(status_text)
        if runtime is not None:
            detail = str(getattr(runtime, "message", "") or "").strip() or "Background agent is ready."
            updated_at = str(getattr(runtime, "updated_at", "") or "").strip()
            if updated_at:
                detail = f"{detail}  Last update: {updated_at}"
            self.status_detail_var.set(detail)
        self._set_status_chip_style(self.status_var.get())

    def _set_status_chip_style(self, status_text: str) -> None:
        lowered = status_text.strip().casefold()
        p = self.palette
        if "failed" in lowered or "error" in lowered:
            color = p["danger"]
        elif "waiting" in lowered or "locked" in lowered:
            color = p["warning"]
        elif "live updated" in lowered or "running" in lowered:
            color = p["success"]
        elif "unknown" in lowered:
            color = p["accent"]
        else:
            color = "#4b5563"
        if hasattr(self, "status_chip_label"):
            self.status_chip_label.configure(fg_color=color)

    def _sync_watch_mode_ui(self) -> None:
        current = self.watch_mode_var.get().strip()
        if current not in {WATCH_MODE_SELECTED_WORKBOOKS, WATCH_MODE_ACHATS_PAIR}:
            self.watch_mode_var.set(WATCH_MODE_SELECTED_WORKBOOKS)
            return
        if current == WATCH_MODE_ACHATS_PAIR:
            self.watch_mode_var.set(WATCH_MODE_SELECTED_WORKBOOKS)
            return
        self._refresh_overview_cards()

    def _sync_performance_mode_ui(self) -> None:
        mode = self.performance_mode_var.get().strip().casefold()
        if mode not in {PERFORMANCE_MODE_SILENT, PERFORMANCE_MODE_LIVE}:
            self.performance_mode_var.set(PERFORMANCE_MODE_SILENT)
            return
        live_enabled = mode == PERFORMANCE_MODE_LIVE
        state = "normal" if live_enabled else "disabled"
        if not live_enabled:
            self.update_open_workbook_var.set("0")
            self.excel_event_monitoring_var.set("0")
            self.allow_live_update_with_autosave_var.set("0")
        for attr in (
            "update_open_workbook_check",
            "excel_event_monitoring_check",
            "allow_live_autosave_check",
            "excel_save_debounce_entry",
            "excel_backend_menu",
        ):
            widget = getattr(self, attr, None)
            if widget is not None:
                try:
                    widget.configure(state=state)
                except Exception:
                    pass

    def _install_dir(self) -> Path:
        return expand_path(self.install_dir_var.get().strip() or DEFAULT_INSTALL_DIR)

    def _normalized_achats_local_file(self) -> str:
        raw = self.achats_local_file_var.get().strip()
        if not raw:
            return ""
        return str(expand_path(raw))

    def _normalized_achats_etranger_file(self) -> str:
        raw = self.achats_etranger_file_var.get().strip()
        if not raw:
            return ""
        return str(expand_path(raw))

    def _normalized_seller_previous_file(self) -> str:
        raw = self.seller_previous_file_var.get().strip()
        if not raw:
            return ""
        return str(expand_path(raw))

    def _normalized_backup_dir(self) -> str:
        raw = self.backup_dir_var.get().strip()
        if raw:
            return str(expand_path(raw))
        return str(self._install_dir() / "backups")

    def _config_file_path(self) -> Path:
        return self._install_dir() / "config.json"

    def _runtime_status_path(self) -> Path:
        return self._install_dir() / "runtime_status.json"

    def _resolve_api_key(self) -> str:
        typed = self.odoo_api_key_var.get().strip()
        if typed:
            return typed
        if self.credential_target:
            return read_secret(self.credential_target)
        raise AgentCredentialError("No Odoo API key is stored. Enter a new API key first.")

    def _refresh_credential_state(self) -> None:
        if self.credential_target and credential_exists(self.credential_target):
            self.api_key_is_stored = True
            self.credential_state_var.set(f"Stored in Windows Credential Manager ({self.credential_target})")
        else:
            self.api_key_is_stored = False
            self.credential_state_var.set("No API key stored")

    def _credential_target_for_save(self, typed_key: str) -> str:
        configured_target = str(self.credential_target or "").strip()
        expected_target = make_credential_target(
            self._config_file_path(),
            self.odoo_db_var.get().strip(),
            self.odoo_login_var.get().strip(),
        )
        if configured_target == expected_target and (not typed_key or credential_exists(configured_target)):
            return configured_target
        return expected_target

    def _build_config(self) -> dict[str, Any]:
        install_dir = self._install_dir()
        config = default_config(install_dir)
        config["odoo"]["url"] = self.odoo_url_var.get().strip()
        config["odoo"]["db"] = self.odoo_db_var.get().strip()
        config["odoo"]["login"] = self.odoo_login_var.get().strip()
        config["odoo"]["record_url_example"] = self.record_url_var.get().strip()
        config["odoo"]["credential_target"] = self._credential_target_for_save(self.odoo_api_key_var.get().strip())
        config["manual"]["last_file"] = self.manual_file_var.get().strip()
        config["background"]["watch_mode"] = WATCH_MODE_SELECTED_WORKBOOKS
        config["background"]["achats_local_file"] = self._normalized_achats_local_file()
        config["background"]["achats_etranger_file"] = self._normalized_achats_etranger_file()
        config["background"]["seller_previous_file"] = self._normalized_seller_previous_file()
        config["background"]["watch_file"] = ""
        config["background"]["watch_folder"] = default_watch_folder()
        config["background"]["recursive"] = False
        performance_mode = self.performance_mode_var.get().strip().casefold()
        if performance_mode not in {PERFORMANCE_MODE_SILENT, PERFORMANCE_MODE_LIVE}:
            performance_mode = PERFORMANCE_MODE_SILENT
        live_mode = performance_mode == PERFORMANCE_MODE_LIVE
        config["background"]["performance_mode"] = performance_mode
        config["background"]["process_existing_on_start"] = self.process_existing_var.get() == "1"
        config["background"]["update_open_workbook"] = live_mode and self.update_open_workbook_var.get() == "1"
        config["background"]["excel_event_monitoring"] = live_mode and self.excel_event_monitoring_var.get() == "1"
        config["background"]["excel_session_backend"] = normalize_excel_session_backend(
            self.excel_session_backend_var.get()
        )
        config["background"]["excel_save_debounce_seconds"] = int(
            self.excel_save_debounce_var.get().strip() or str(DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS)
        )
        config["background"]["allow_live_update_with_autosave"] = live_mode and self.allow_live_update_with_autosave_var.get() == "1"
        config["background"]["visible_excel"] = self.visible_excel_var.get() == "1"
        config["background"]["write_report_file"] = self.write_report_var.get() == "1"
        config["background"]["stable_backup_name"] = self.stable_backup_var.get() == "1"
        config["background"]["settle_seconds"] = int(self.settle_seconds_var.get().strip() or "15")
        config["background"]["retry_delay_seconds"] = int(self.retry_seconds_var.get().strip() or "45")
        config["paths"]["backup_dir"] = self._normalized_backup_dir()
        existing_updates = self.config.get("updates", {}) if isinstance(self.config.get("updates"), dict) else {}
        config["updates"]["manifest_url"] = self._update_manifest_url()
        config["updates"]["last_checked_at"] = str(existing_updates.get("last_checked_at") or "")
        config["updates"]["last_seen_version"] = str(existing_updates.get("last_seen_version") or "")
        return config

    def _load_config(self) -> None:
        config, messages = load_normalized_config(self.config_path)
        self.config = config
        self.install_dir_var.set(str(self.config_path.parent))
        self.odoo_url_var.set(str(config["odoo"].get("url") or DEFAULT_ODOO_URL))
        self.odoo_db_var.set(str(config["odoo"].get("db") or ""))
        self.odoo_login_var.set(str(config["odoo"].get("login") or ""))
        self.record_url_var.set(str(config["odoo"].get("record_url_example") or ""))
        self.manual_file_var.set(str(config.get("manual", {}).get("last_file") or ""))
        background = config.get("background", {})
        self.watch_mode_var.set(WATCH_MODE_SELECTED_WORKBOOKS)
        self.achats_local_file_var.set(str(background.get("achats_local_file") or ""))
        self.achats_etranger_file_var.set(str(background.get("achats_etranger_file") or ""))
        self.seller_previous_file_var.set(str(background.get("seller_previous_file") or ""))
        self.watch_file_var.set(str(background.get("watch_file") or ""))
        self.watch_folder_var.set(str(background.get("watch_folder") or default_watch_folder()))
        self.backup_dir_var.set(str(config.get("paths", {}).get("backup_dir") or self._install_dir() / "backups"))
        self.performance_mode_var.set(str(background.get("performance_mode") or PERFORMANCE_MODE_SILENT))
        self.process_existing_var.set("1" if background.get("process_existing_on_start", False) else "0")
        self.update_open_workbook_var.set("1" if background.get("update_open_workbook", False) else "0")
        self.excel_event_monitoring_var.set("1" if background.get("excel_event_monitoring", False) else "0")
        self.excel_session_backend_var.set(normalize_excel_session_backend(background.get("excel_session_backend")))
        self.excel_save_debounce_var.set(str(background.get("excel_save_debounce_seconds", DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS)))
        self.allow_live_update_with_autosave_var.set("1" if background.get("allow_live_update_with_autosave", False) else "0")
        self.recursive_var.set("1" if background.get("recursive", False) else "0")
        self.visible_excel_var.set("1" if background.get("visible_excel", False) else "0")
        self.write_report_var.set("1" if background.get("write_report_file", False) else "0")
        self.stable_backup_var.set("1" if background.get("stable_backup_name", True) else "0")
        self.settle_seconds_var.set(str(background.get("settle_seconds", 15)))
        self.retry_seconds_var.set(str(background.get("retry_delay_seconds", 45)))
        updates = config.get("updates", {}) if isinstance(config.get("updates"), dict) else {}
        self.update_manifest_url_var.set(str(updates.get("manifest_url") or DEFAULT_UPDATE_URL))
        last_seen = str(updates.get("last_seen_version") or "").strip()
        if last_seen:
            self.update_state_var.set(f"Current: {APP_VERSION} - Latest seen: {last_seen}")
        self.credential_target = str(config["odoo"].get("credential_target") or "")
        self._migrate_legacy_api_key_if_needed()
        self._refresh_credential_state()
        self._sync_watch_mode_ui()
        self._sync_performance_mode_ui()
        for message in messages:
            self.append_status(message)
        self.append_status(f"Loaded config from {self.config_path}")

    def _migrate_legacy_api_key_if_needed(self) -> None:
        if not self.config_path.exists():
            return
        try:
            raw = json.loads(self.config_path.read_text(encoding="utf-8-sig"))
        except Exception as exc:
            self.append_status(f"Skipped legacy API key migration because config is unreadable: {exc}")
            return
        legacy_api_key = str(raw.get("odoo", {}).get("api_key") or "").strip()
        if not legacy_api_key:
            return
        if not self.credential_target:
            self.credential_target = make_credential_target(
                self.config_path,
                self.odoo_db_var.get().strip(),
                self.odoo_login_var.get().strip(),
            )
            self.config["odoo"]["credential_target"] = self.credential_target
        store_secret(self.credential_target, self.odoo_login_var.get().strip(), legacy_api_key)
        self.append_status("Migrated legacy API key into Windows Credential Manager.")
        normalized = self._build_config()
        save_normalized_config(self.config_path, normalized)

    def _save_config_internal(self) -> Path:
        config_path = self._config_file_path()
        config = self._build_config()
        credential_target = config["odoo"]["credential_target"]
        typed_key = self.odoo_api_key_var.get().strip()
        if typed_key:
            store_secret(credential_target, config["odoo"]["login"], typed_key)
            self.odoo_api_key_var.set("")
            self.append_status("Stored API key in Windows Credential Manager.")
        self.credential_target = credential_target
        save_normalized_config(config_path, config)
        self.config_path = config_path
        self.config = config
        self._refresh_credential_state()
        self._refresh_overview_cards()
        self.append_status(f"Saved config to {config_path}")
        return config_path

    def _ensure_runtime_files(self) -> Path:
        install_dir = self._install_dir()
        install_dir.mkdir(parents=True, exist_ok=True)
        (install_dir / "reports").mkdir(parents=True, exist_ok=True)
        Path(self._normalized_backup_dir()).mkdir(parents=True, exist_ok=True)

        if getattr(sys, "frozen", False):
            self.append_status(f"Running as compiled executable. Prepared {install_dir}")
        else:
            for name in SUPPORT_FILES:
                source = (self.script_dir / name).resolve()
                destination = (install_dir / name).resolve()
                if source != destination and source.exists():
                    shutil.copy2(source, destination)
            self.append_status(f"Copied runtime files to {install_dir}")
        return install_dir

    def _validate_common(self) -> None:
        if not self.odoo_url_var.get().strip():
            raise ValueError("Odoo URL is required.")
        if not self.odoo_db_var.get().strip():
            raise ValueError("Odoo database is required.")
        if not self.odoo_login_var.get().strip():
            raise ValueError("Odoo login is required.")
        validate_odoo_settings(
            self.odoo_url_var.get().strip(),
            self.odoo_db_var.get().strip(),
            self.odoo_login_var.get().strip(),
            self.record_url_var.get().strip(),
        )
        backup_dir = Path(self._normalized_backup_dir())
        backup_dir.mkdir(parents=True, exist_ok=True)

    def _ensure_python_dependencies(self, *, include_background: bool) -> None:
        missing: list[str] = []
        if importlib.util.find_spec("pythoncom") is None:
            missing.append("pywin32")
        if include_background and importlib.util.find_spec("watchdog") is None:
            missing.append("watchdog")
        if importlib.util.find_spec("openpyxl") is None:
            missing.append("openpyxl")
        current_backend = self.excel_session_backend_var.get().strip().casefold()
        if normalize_excel_session_backend(current_backend) != current_backend:
            self.excel_session_backend_var.set(DEFAULT_EXCEL_SESSION_BACKEND)
        if missing:
            raise RuntimeError(
                "Missing Python packages: "
                + ", ".join(sorted(set(missing)))
                + ". Re-run the launcher/install script so it can install them."
            )

    def _validate_api_key_presence(self) -> None:
        if not self.odoo_api_key_var.get().strip() and not (self.credential_target and credential_exists(self.credential_target)):
            raise AgentCredentialError("No Odoo API key is stored. Enter a new API key first.")

    def _validate_watch_target(self) -> None:
        configured = self._configured_selected_workbooks()
        if not configured:
            raise ValueError("Choose at least one workbook for background processing.")
        for raw_path in configured:
            workbook = Path(raw_path)
            if not is_supported_workbook(workbook):
                raise ValueError(f"Invalid workbook type: {workbook.name}")
            if workbook.exists():
                if not workbook.is_file():
                    raise ValueError(f"Invalid workbook path: {workbook}")
                continue
            if not workbook.parent.exists() or not workbook.parent.is_dir():
                raise ValueError(f"Workbook parent folder does not exist: {workbook.parent}")

    def _validate_manual_file(self) -> Path:
        workbook = Path(self.manual_file_var.get().strip())
        if not workbook.exists():
            raise FileNotFoundError("Selected Excel file does not exist.")
        if not is_supported_workbook(workbook):
            raise ValueError("Selected file must be .xlsx, .xlsm, or .xls.")
        return workbook.resolve()

    def _configured_achats_workbook(self, kind: str) -> Path:
        if kind == "local":
            raw_path = self._normalized_achats_local_file()
            display_name = "ACHATS LOCAL"
        elif kind == "etranger":
            raw_path = self._normalized_achats_etranger_file()
            display_name = "ACHATS ETRANGER"
        else:
            raw_path = self._normalized_seller_previous_file()
            display_name = "Seller / Previous"
        if not raw_path:
            raise FileNotFoundError(f"Choose the {display_name} workbook first.")
        workbook = Path(raw_path)
        if not workbook.exists() or not workbook.is_file():
            raise FileNotFoundError(f"{display_name} workbook does not exist: {workbook}")
        if not is_supported_workbook(workbook):
            raise ValueError(f"{display_name} workbook must be .xlsx, .xlsm, or .xls.")
        return workbook.resolve()

    def _authenticate_odoo(self) -> int:
        client = OdooClient(
            self.odoo_url_var.get().strip(),
            self.odoo_db_var.get().strip(),
            self.odoo_login_var.get().strip(),
            self._resolve_api_key(),
        )
        return client.authenticate()

    def save_config(self) -> None:
        try:
            self._validate_common()
            config_path = self._save_config_internal()
            messagebox.showinfo(APP_NAME, f"Config saved to:\n{config_path}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))

    def install_agent(self) -> None:
        def work() -> None:
            try:
                self._validate_common()
                self._validate_api_key_presence()
                self._ensure_python_dependencies(include_background=True)
                self._validate_watch_target()
                uid = self._authenticate_odoo()
                self.append_status(f"Odoo authentication succeeded. UID={uid}")

                install_dir = self._ensure_runtime_files()
                config_path = self._save_config_internal()
                stopped = UiRuntime.stop_agent_processes()
                if stopped:
                    self.append_status(f"Stopped {stopped} running agent process(es)")

                UiRuntime.create_startup_shortcut(current_pythonw(), install_dir, config_path)
                self.append_status(f"Created startup shortcut: {STARTUP_SHORTCUT}")
                UiRuntime.start_agent_process(current_pythonw(), install_dir, config_path)
                self.append_status("Started background agent.")
                self.root.after(0, lambda: messagebox.showinfo(APP_NAME, "Agent installed and started."))
            except Exception as exc:
                self.append_status(f"Install failed: {exc}")
                self.root.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def start_agent(self) -> None:
        def work() -> None:
            try:
                self._validate_common()
                self._validate_api_key_presence()
                self._ensure_python_dependencies(include_background=True)
                self._validate_watch_target()
                uid = self._authenticate_odoo()
                self.append_status(f"Odoo authentication succeeded. UID={uid}")
                self._ensure_runtime_files()
                config_path = self._save_config_internal()
                stopped = UiRuntime.stop_agent_processes()
                if stopped:
                    self.append_status(f"Stopped {stopped} running agent process(es)")
                UiRuntime.start_agent_process(current_pythonw(), self._install_dir(), config_path)
                self.append_status("Start command sent to background agent.")
            except Exception as exc:
                self.append_status(f"Start failed: {exc}")
                self.root.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def stop_agent(self) -> None:
        def work() -> None:
            try:
                stopped = UiRuntime.stop_agent_processes()
                self.append_status(f"Stopped {stopped} agent process(es).")
            except Exception as exc:
                self.append_status(f"Stop failed: {exc}")
                self.root.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def test_odoo(self) -> None:
        def work() -> None:
            try:
                self._validate_common()
                self._validate_api_key_presence()
                uid = self._authenticate_odoo()
                self.append_status(f"Odoo authentication succeeded. UID={uid}")
                self.root.after(0, lambda: messagebox.showinfo(APP_NAME, "Odoo connection succeeded."))
            except Exception as exc:
                self.append_status(f"Odoo test failed: {exc}")
                self.root.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def _record_update_check(self, info: UpdateInfo) -> None:
        config = self._build_config()
        config["updates"]["last_checked_at"] = dt.datetime.now().isoformat(timespec="seconds")
        config["updates"]["last_seen_version"] = info.version
        save_normalized_config(self._config_file_path(), config)
        self.config = config

    def _update_manifest_url(self) -> str:
        value = self.update_manifest_url_var.get().strip()
        if not value:
            value = DEFAULT_UPDATE_URL
            self.update_manifest_url_var.set(value)
        return value

    def check_for_updates(self) -> None:
        def work() -> None:
            try:
                manifest_url = self._update_manifest_url()
                self.append_status("Checking for updates...")
                info = check_for_update(manifest_url, APP_VERSION)
                self._last_update_info = info
                self._record_update_check(info)
                if info.has_update:
                    message = f"Update available: {APP_VERSION} -> {info.version}"
                    detail = info.notes.strip() or info.download_url
                else:
                    message = f"No update available. Current version: {APP_VERSION}"
                    detail = f"Latest version in manifest: {info.version}"
                self.root.after(0, lambda: self.update_state_var.set(message))
                self.root.after(0, lambda: self.update_detail_var.set(detail))
                self.append_status(message)
            except Exception as exc:
                message = str(exc)
                self.append_status(f"Update check failed: {message}")
                self.root.after(0, lambda: self.update_state_var.set("Update check failed"))
                self.root.after(0, lambda: self.update_detail_var.set(message))
                self.root.after(0, lambda: messagebox.showerror(APP_NAME, message))

        threading.Thread(target=work, daemon=True).start()

    def install_update(self) -> None:
        def work() -> None:
            try:
                if not getattr(sys, "frozen", False):
                    raise RuntimeError("Automatic install is available only in the packaged OdooExcelAgent.exe.")
                info = self._last_update_info
                manifest_url = self._update_manifest_url()
                if info is None:
                    self.append_status("Checking for updates before install...")
                    info = check_for_update(manifest_url, APP_VERSION)
                    self._last_update_info = info
                    self._record_update_check(info)
                if not info.has_update:
                    raise RuntimeError(f"No newer version is available. Current version: {APP_VERSION}.")
                self._save_config_internal()
                self.append_status(f"Downloading update {info.version}...")
                asset_path = download_update_asset(info)
                self.append_status("Preparing update package...")
                prepared = prepare_update_payload(asset_path)
                stopped = UiRuntime.stop_agent_processes()
                if stopped:
                    self.append_status(f"Stopped {stopped} running agent process(es)")
                current_exe = Path(sys.executable).resolve()
                install_dir = current_exe.parent
                config_path = self._config_file_path()
                schedule_update_install(
                    current_exe=current_exe,
                    payload_dir=prepared.payload_dir,
                    install_dir=install_dir,
                    config_path=config_path,
                    restart=True,
                )
                self.append_status("Update scheduled. The app will close and reopen after replacement.")
                self.root.after(0, lambda: messagebox.showinfo(APP_NAME, "Update downloaded. The app will close and reopen after installing."))
                self.root.after(800, self.root.destroy)
            except Exception as exc:
                message = str(exc)
                self.append_status(f"Install update failed: {message}")
                self.root.after(0, lambda: self.update_state_var.set("Install update failed"))
                self.root.after(0, lambda: self.update_detail_var.set(message))
                self.root.after(0, lambda: messagebox.showerror(APP_NAME, message))

        threading.Thread(target=work, daemon=True).start()

    def _process_workbook_path(self, workbook: Path, source_label: str, workbook_slot: str = "") -> None:
        def work() -> None:
            try:
                self._validate_common()
                self._validate_api_key_presence()
                self._ensure_python_dependencies(include_background=False)
                workbook_path = workbook.expanduser().resolve()
                if not workbook_path.exists():
                    raise FileNotFoundError(f"Selected Excel file does not exist: {workbook_path}")
                if not is_supported_workbook(workbook_path):
                    raise ValueError("Selected file must be .xlsx, .xlsm, or .xls.")
                self._authenticate_odoo()
                config_path = self._save_config_internal()
                summary = process_workbook(
                    workbook_path=workbook_path,
                    odoo_url=self.odoo_url_var.get().strip(),
                    odoo_db=self.odoo_db_var.get().strip(),
                    odoo_login=self.odoo_login_var.get().strip(),
                    odoo_api_key=self._resolve_api_key(),
                    record_url_example=self.record_url_var.get().strip(),
                    backup_dir=Path(self._normalized_backup_dir()),
                    write_report_file=False,
                    stable_backup_name=self.stable_backup_var.get() == "1",
                    apply=True,
                    visible_excel=self.visible_excel_var.get() == "1",
                    allow_open_workbook_update=self.update_open_workbook_var.get() == "1",
                    excel_session_backend=normalize_excel_session_backend(self.excel_session_backend_var.get()),
                    excel_save_debounce_seconds=int(
                        self.excel_save_debounce_var.get().strip() or str(DEFAULT_EXCEL_SAVE_DEBOUNCE_SECONDS)
                    ),
                    allow_live_update_with_autosave=self.allow_live_update_with_autosave_var.get() == "1",
                    workbook_slot=workbook_slot,
                    performance_mode=self.performance_mode_var.get().strip() or PERFORMANCE_MODE_SILENT,
                )
                counts = ", ".join(f"{key}={value}" for key, value in sorted(summary.status_counts.items())) or "no matches"
                mode = "live Excel update" if summary.live_update_used else summary.workbook_state.replace("_", " ")
                self.append_status(f"{source_label}: processed {workbook_path.name}: {counts} ({mode})")
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        APP_NAME,
                        f"Processed: {workbook_path.name}\nLinked: {summary.linked_count}\nWorkbook state: {summary.workbook_state}\nStatuses: {counts}",
                    ),
                )
                self.append_status(f"Config in use: {config_path}")
            except Exception as exc:
                self.append_status(f"{source_label}: processing failed: {exc}")
                self.root.after(0, lambda: messagebox.showerror(APP_NAME, str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def process_selected_file(self) -> None:
        workbook = Path(self.manual_file_var.get().strip())
        self._process_workbook_path(workbook, source_label="Manual run")

    def process_achats_local_file(self) -> None:
        try:
            workbook = self._configured_achats_workbook("local")
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.manual_file_var.set(str(workbook))
        self._process_workbook_path(workbook, source_label="ACHATS LOCAL", workbook_slot=WORKBOOK_SLOT_ACHATS_LOCAL)

    def process_achats_etranger_file(self) -> None:
        try:
            workbook = self._configured_achats_workbook("etranger")
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.manual_file_var.set(str(workbook))
        self._process_workbook_path(workbook, source_label="ACHATS ETRANGER", workbook_slot=WORKBOOK_SLOT_ACHATS_ETRANGER)

    def process_seller_previous_file(self) -> None:
        try:
            workbook = self._configured_achats_workbook("seller")
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.manual_file_var.set(str(workbook))
        self._process_workbook_path(workbook, source_label="Seller / Previous", workbook_slot=WORKBOOK_SLOT_SELLER_PREVIOUS)

    def clear_stored_api_key(self) -> None:
        if not self.credential_target:
            messagebox.showinfo(APP_NAME, "No stored API key was found.")
            return
        delete_secret(self.credential_target)
        self._refresh_credential_state()
        self.append_status("Removed stored API key from Windows Credential Manager.")

    def refresh_status(self) -> None:
        try:
            processes = UiRuntime.list_agent_processes()
            runtime = read_runtime_status(get_runtime_status_path(self._config_file_path()))
            if processes:
                if runtime.last_issue_code in {"waiting_for_close", "excel_waiting_close", "excel_read_only", "excel_autosave_deferred", "excel_ambiguous_instance"}:
                    status_text = "Waiting for Excel"
                elif runtime.last_issue_code == "excel_backend_unavailable":
                    status_text = "Excel backend issue"
                elif runtime.last_issue_code == "missing_required_header":
                    status_text = "Header missing"
                elif runtime.last_issue_code == "excel_live_updated":
                    status_text = "Live updated"
                elif runtime.last_issue_code:
                    status_text = "Running"
                else:
                    status_text = "Running"
            else:
                mapping = {
                    "missing_api_key": "Setup required",
                    "odoo_auth_failed": "Odoo auth failed",
                    "invalid_watch_folder": "Invalid watch target",
                    "invalid_watch_target": "Invalid watch target",
                    "startup_failed": "Startup failed",
                    "stopped": "Stopped",
                    "running": "Stopped",
                }
                status_text = mapping.get(runtime.last_issue_code) or mapping.get(runtime.state, runtime.state.replace("_", " ").title())
            self.root.after(0, lambda: self._refresh_overview_cards(runtime=runtime, status_text=status_text))
        except Exception as exc:
            self.root.after(0, lambda: self._refresh_overview_cards(status_text="Unknown"))
            self.root.after(0, lambda: self.status_detail_var.set("Could not read runtime status."))
            self.append_status(f"Status refresh failed: {exc}")

    def open_selected_file(self) -> None:
        path = Path(self.manual_file_var.get().strip())
        if path.exists():
            os.startfile(str(path))
        else:
            messagebox.showinfo(APP_NAME, "Select a valid Excel file first.")

    def _open_configured_workbook(self, path_var: StringVar) -> None:
        raw_path = str(path_var.get() or "").strip()
        if not raw_path:
            messagebox.showinfo(APP_NAME, "Choose a valid workbook first.")
            return
        path = Path(raw_path)
        if path.exists():
            os.startfile(str(path))
        else:
            messagebox.showinfo(APP_NAME, f"Workbook does not exist:\n{path}")

    def open_watch_target(self) -> None:
        configured = self._configured_selected_workbooks()
        if not configured:
            messagebox.showinfo(APP_NAME, "Choose at least one workbook first.")
            return
        parent = Path(configured[0]).parent
        if parent.exists():
            os.startfile(str(parent))
        else:
            messagebox.showinfo(APP_NAME, f"Folder does not exist:\n{parent}")

    def open_install_folder(self) -> None:
        os.startfile(str(self._install_dir()))

    def open_log(self) -> None:
        path = self._install_dir() / "agent.log"
        if path.exists():
            os.startfile(str(path))
        else:
            messagebox.showinfo(APP_NAME, "Log file does not exist yet.")

    def open_backups(self) -> None:
        path = Path(self._normalized_backup_dir())
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(str(path))

    def open_config_file(self) -> None:
        config_path = self._config_file_path()
        if config_path.exists():
            os.startfile(str(config_path))
        else:
            messagebox.showinfo(APP_NAME, "Config file does not exist yet.")

    def on_close(self) -> None:
        try:
            running = bool(UiRuntime.list_agent_processes())
        except Exception:
            running = False
        if running:
            messagebox.showinfo(APP_NAME, "The setup window will close. The background agent will keep running in the system tray.")
        self.root.destroy()


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", default="", help="Optional path to config.json")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    root = ctk.CTk()
    app = AgentControlApp(root, config_arg=args.config)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

