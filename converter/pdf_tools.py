"""PDF araçları: birleştirme ve düz metin çıkarma. Tamamen yerel.

Buluta veri göndermez; PDFium ve pdfminer tabanlı pdfplumber kullanır.
"""
from __future__ import annotations

import os
import tempfile


def convert_pdfs_to_pdf(inputs: list[str], output_path: str) -> tuple[bool, str]:
    """Birden çok PDF'i list sırasıyla tek PDF'te birleştirir."""
    ins = [p for p in inputs
           if os.path.splitext(p)[1].lower() == ".pdf"
           and os.path.exists(p)]
    if not ins:
        return False, "Birleştirilecek PDF bulunamadı."
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    try:
        import pypdfium2 as pdfium
    except Exception as e:
        return False, f"Gerekli kütüphane yüklenemedi: {e}"
    out = None
    try:
        out = pdfium.PdfDocument.new()
        pages = 0
        for p in ins:
            src = pdfium.PdfDocument(p)
            try:
                out.import_pages(src)
                pages += len(src)
            finally:
                src.close()
        if pages == 0:
            return False, "PDF'lerde içerik bulunamadı."
        out.save(output_path)
        if len(ins) == 1:
            return True, f"Tek PDF ({pages} sayfa) kaydedildi."
        return True, f"{len(ins)} PDF birleştirildi ({pages} sayfa)."
    except Exception as e:
        return False, f"Birleştirme hatası: {e}"
    finally:
        if out is not None:
            out.close()


def _ocr_pdf_page(page) -> str:
    """PDFium sayfasını geçici PNG'ye çevirip yerel OCR ile okur."""
    bitmap = None
    image = None
    tmp = ""
    try:
        bitmap = page.render(scale=2.0)
        image = bitmap.to_pil().convert("RGB")
        fd, tmp = tempfile.mkstemp(prefix="ocr_", suffix=".png")
        os.close(fd)
        image.save(tmp, "PNG")
        from converter.ocr import ocr_image_to_text
        ok, text = ocr_image_to_text(tmp)
        return text.strip() if ok else ""
    except Exception:
        return ""
    finally:
        if image is not None:
            image.close()
        if bitmap is not None:
            bitmap.close()
        try:
            if tmp and os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        try:
            page.close()
        except Exception:
            pass


def convert_pdf_to_text(input_path: str, output_path: str) -> tuple[bool, str]:
    """PDF'teki metni .txt olarak çıkarır. Taranmış sayfalarda OCR denenir."""
    if os.path.splitext(input_path)[1].lower() != ".pdf":
        return False, "Sadece .pdf destekleniyor."
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    try:
        import pdfplumber
    except Exception as e:
        return False, f"Gerekli kütüphane yüklenemedi: {e}"
    try:
        doc = pdfplumber.open(input_path)
    except Exception as e:
        return False, f"PDF açılamadı: {e}"
    try:
        import pypdfium2 as pdfium
    except Exception:
        pdfium = None
    render_doc = None
    parts: list[str] = []
    n_ocr = 0
    try:
        for i, page in enumerate(doc.pages):
            try:
                t = (page.extract_text() or "").strip()
            except Exception:
                t = ""
            if not t and pdfium is not None:
                # Vektör metni yok -> taranmış sayfa olabilir, yerel OCR dene
                try:
                    if render_doc is None:
                        render_doc = pdfium.PdfDocument(input_path)
                    t = _ocr_pdf_page(render_doc[i]).strip()
                    if t:
                        n_ocr += 1
                except Exception:
                    t = ""
            if t:
                parts.append(t)
    finally:
        doc.close()
        if render_doc is not None:
            render_doc.close()
    if not parts:
        return False, ("Metin bulunamadı (taranmış PDF olabilir ve "
                       "OCR kapalı olabilir).")
    text = "\n\n".join(parts)
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception as e:
        return False, f"Yazma hatası: {e}"
    info = f"{len(parts)} sayfa metin çıkarıldı"
    if n_ocr:
        info += f" ({n_ocr} sayfa OCR)"
    return True, info + "."
