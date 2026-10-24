# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import os

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata


root = Path.cwd()
app_version = os.environ.get("OUTLOOK_REHBERI_VERSION", "5.0.0")
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
    "keyring.backends.macOS",
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
    [],
    exclude_binaries=True,
    name="OutlookFirmaRehberi",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
collection = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="OutlookFirmaRehberi",
)
app = BUNDLE(
    collection,
    name="Outlook Firma Rehberi.app",
    icon=str(root / "assets/icon.icns"),
    bundle_identifier="com.sseymen.outlookfirmarehberi",
    info_plist={
        "CFBundleDisplayName": "Outlook Firma Rehberi",
        "CFBundleShortVersionString": app_version,
        "NSHighResolutionCapable": True,
        "LSMinimumSystemVersion": "12.0",
    },
)
