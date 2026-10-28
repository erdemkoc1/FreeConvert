# -*- mode: python ; coding: utf-8 -*-
"""FreeConvert kaldırıcısı (uninstall.exe) için PyInstaller tanımı."""
import os

ROOT = SPECPATH
APP = os.path.join(ROOT, "tools", "uninstall.py")
ICON = os.path.join(ROOT, "assets", "icon", "freeconvert.ico")

a = Analysis(
    [APP],
    pathex=[ROOT],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name="uninstall",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=ICON,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
