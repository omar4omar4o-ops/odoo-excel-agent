# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = [
    'customtkinter',
    'win32com',
    'win32com.client',
    'watchdog',
    'openpyxl',
    'pythoncom',
    'pywintypes',
    'win32timezone',
    'odoo_excel_agent_ui',
    'odoo_excel_background',
    'link_odoo_vendor_bills',
    'odoo_excel_agent_support',
    'odoo_excel_updater',
]
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

excludes = [
    'xlwings',
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
    'pandas', 'numpy', 'scipy', 'matplotlib', 'plotly',
    'IPython', 'jupyter', 'jupyter_client', 'jupyter_core',
    'streamlit', 'tensorflow', 'torch', 'torchaudio', 'torchvision',
    'sklearn', 'skimage', 'cv2', 'datasets', 'transformers',
    'yt_dlp', 'curl_cffi', 'onnxruntime',
]


a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='OdooExcelAgent',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
