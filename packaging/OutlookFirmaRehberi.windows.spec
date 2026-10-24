# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata


root = Path.cwd()
config_path = root / "outlook_rehberi" / "app_config.json"
datas = []
for package in ("msal", "keyring", "certifi"):
    datas += collect_data_files(package)
    datas += copy_metadata(package)
if config_path.is_file():
    datas.append((str(config_path), "outlook_rehberi"))

hiddenimports = [
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
    "keyring.backends.Windows",
    *collect_submodules("msal"),
    *collect_submodules("keyring.backends"),
]

a = Analysis(
    [str(root / "outlook_rehberi/__main__.py")],
    pathex=[str(root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["win32com", "pythoncom", "tkinter"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="OutlookFirmaRehberi",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon=str(root / "assets/icon.ico"),
)
