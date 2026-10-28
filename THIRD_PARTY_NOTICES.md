# Third-party notices

FreeConvert is distributed as a portable offline application. The runtime
components below are permissive/open-source components; no AGPL or commercial
runtime dependency is included.

## Runtime components

| Component | Use | License |
|---|---|---|
| Python | Application runtime, bundled by PyInstaller | PSF License |
| python-docx | DOCX reading/writing | MIT |
| ReportLab | DOCX-to-PDF layout generation | BSD-style license |
| Pillow | Image conversion and PDFium bitmap output | HPND / Pillow license |
| pdfplumber | PDF words, lines, tables and cell geometry | MIT |
| pdfminer.six | PDF text/object parsing used by pdfplumber | MIT |
| pypdfium2 / PDFium | PDF page access and rendering | BSD-3-Clause / Apache-2.0 notices |
| Tesseract OCR | Bundled Turkish/English OCR runtime | Apache-2.0 |
| Noto Sans | Unicode font for Turkish and other Latin text | SIL Open Font License 1.1; see `assets/fonts/OFL.txt` |
| tkinterdnd2 | Optional drag-and-drop support | MIT |

The Tesseract language data notice is included at
`tesseract_runtime/tessdata/LICENSE`.

## Build-only tool

PyInstaller is used to build the executable. Its GPL-with-special-exception
build tooling is not a runtime conversion dependency and does not cause the
application's own source code to be relicensed.

## Removed dependencies

The project intentionally does **not** bundle or call:

- PyMuPDF / MuPDF (AGPL/commercial dual-licensed)
- Microsoft Word COM or `docx2pdf`
- LibreOffice
- OnlyOffice x2t
- Java, Ghostscript, Tabula, Camelot ML models, or cloud conversion services

All conversion code is local. The only runtime subprocess is the bundled
Tesseract executable used for OCR; it is embedded in the one-file build and is
never required as a separately installed system program.
