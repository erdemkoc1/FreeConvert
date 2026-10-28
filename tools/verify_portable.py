"""Portable exe paketinde yasaklı paket ve eksik runtime kontrolü."""
from __future__ import annotations

import sys
from pathlib import Path

from PyInstaller.archive.readers import CArchiveReader

FORBIDDEN_PREFIXES = (
    "fitz/", "pymupdf/", "docx2pdf/", "win32com/",
)
REQUIRED = {
    "assets/fonts/NotoSans-Regular.ttf",
    "assets/fonts/NotoSans-Bold.ttf",
    "assets/fonts/NotoSans-Italic.ttf",
    "assets/fonts/OFL.txt",
    "assets/icon/freeconvert.ico",
    "tesseract_runtime/tesseract.exe",
    "tesseract_runtime/tessdata/eng.traineddata",
    "tesseract_runtime/tessdata/tur.traineddata",
    "pypdfium2_raw/pdfium.dll",
    "THIRD_PARTY_NOTICES.md",
    "LICENSE",
}


def main(path: str) -> int:
    exe = Path(path)
    if not exe.is_file():
        print(f"Missing output: {exe}")
        return 2
    archive = CArchiveReader(str(exe))
    names = {name.replace("\\", "/") for name in archive.toc}
    forbidden = sorted(
        name for name in names
        if name.lower().startswith(FORBIDDEN_PREFIXES)
    )
    missing = sorted(REQUIRED - names)
    print(f"Portable entries: {len(names)}")
    print(f"Forbidden package entries: {forbidden}")
    print(f"Missing required entries: {missing}")
    if forbidden or missing:
        return 1
    print("Portable package audit: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1] if len(sys.argv) == 2 else ""))
