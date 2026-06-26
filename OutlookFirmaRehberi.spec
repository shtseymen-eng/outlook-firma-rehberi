# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules

datas     = []
binaries  = []
hiddenimports = [
    'win32com', 'win32com.client', 'pythoncom', 'pywintypes',
    'openpyxl', 'openpyxl.styles', 'openpyxl.utils',
]

for pkg in ['win32com', 'openpyxl']:
    d, b, h = collect_all(pkg)
    datas    += d
    binaries += b
    hiddenimports += h

hiddenimports += collect_submodules('win32com')
hiddenimports  = list(set(hiddenimports))

a = Analysis(
    ['uygulama.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'pandas', 'PIL', 'scipy'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='OutlookFirmaRehberi',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # Pencereli uygulama — siyah konsol yok
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
)
