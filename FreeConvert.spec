# -*- mode: python ; coding: utf-8 -*-
"""FreeConvert için tek dosyalık Windows paketleme tanımı."""
import os

from PyInstaller.utils.hooks import collect_all

ROOT = SPECPATH
APP = os.path.join(ROOT, "app.py")
TESSERACT_RUNTIME = os.path.join(ROOT, "tesseract_runtime")

datas = [
    (TESSERACT_RUNTIME, "tesseract_runtime"),
    (os.path.join(ROOT, "assets", "fonts"), "assets/fonts"),
    (os.path.join(ROOT, "THIRD_PARTY_NOTICES.md"), "."),
    (os.path.join(ROOT, "LICENSE"), "."),
]
binaries = []
hiddenimports = []

# PDFium tabanlı rendering ve permissive PDF metin/tablo parser'ları.
for package in ("pdfplumber", "pdfminer", "pypdfium2"):
    tmp_ret = collect_all(package)
    datas += tmp_ret[0]
    binaries += tmp_ret[1]
    hiddenimports += tmp_ret[2]

a = Analysis(
    [APP],
    pathex=[ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# COLLECT yoktur: scripts, binaries ve datas doğrudan exe içine konur.
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="FreeConvert",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
