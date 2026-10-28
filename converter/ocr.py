"""Yerel OCR (Tesseract CLI, LSTM yapay sinir ağı).

Tesseract motoru ve tur/eng dil verisi uygulamanın içine gömülüdür
(tesseract_runtime/); kullanıcı ek bir kurulum yapmaz, ağ erişimi yoktur.
Sistemde ayrıca kurulu Tesseract varsa o değil, gömülü olan kullanılır.

Tesseract runtime bulunamazsa tüm fonksiyonlar temiz şekilde (False, ...)
döner; arayüz bu durumu gösterir, diğer işlevler etkilenmez.
"""
from __future__ import annotations

import csv
import io
import os
import subprocess
import tempfile

from converter._common import find_tesseract, tessdata_dir


def tesseract_available() -> bool:
    """Gömülü OCR motorunu gerçekten çalıştırıp tur/eng dillerini doğrular."""
    tess = find_tesseract()
    td = tessdata_dir()
    if not tess or not td:
        return False
    try:
        result = subprocess.run(
            [tess, "--tessdata-dir", td, "--list-langs"],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            timeout=15, check=False)
        if result.returncode != 0:
            return False
        languages = result.stdout.decode("utf-8", errors="replace").lower()
        return "tur" in languages and "eng" in languages
    except Exception:
        return False


def ocr_image_to_text(image_path: str) -> tuple[bool, str]:
    """Görseldeki yazıyı okur. Dönüş: (ok, metin veya hata)."""
    tess = find_tesseract()
    if not tess:
        return False, "Tesseract kurulu değil (yerel OCR kapalı)."
    td = tessdata_dir()
    for lang in ("tur+eng", "eng", None):
        try:
            cmd = [tess, os.path.abspath(image_path), "stdout"]
            if td:
                cmd += ["--tessdata-dir", td]
            if lang:
                cmd += ["-l", lang]
            r = subprocess.run(cmd, stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE,
                               timeout=180, check=False)
            if r.returncode == 0:
                return True, r.stdout.decode("utf-8",
                                             errors="replace").strip()
        except Exception:
            continue
    return False, "OCR başarısız."


def ocr_image_to_words(image_path: str) -> tuple[bool, list[tuple]]:
    """Tesseract TSV çıktısından kelime koordinatlarını döndürür.

    Farklı page segmentation modları denenir; tabloda en çok güvenilir
    kelime koordinatı kümesi seçilir.
    """
    tess = find_tesseract()
    if not tess:
        return False, []
    td = tessdata_dir()
    best: list[tuple] = []
    succeeded = False
    for lang in ("tur+eng", "eng"):
        for psm in ("6", "11", "3"):
            try:
                cmd = [tess, os.path.abspath(image_path), "stdout"]
                if td:
                    cmd += ["--tessdata-dir", td]
                cmd += ["-l", lang, "--psm", psm,
                        "-c", "tessedit_create_tsv=1"]
                result = subprocess.run(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    timeout=180, check=False)
                if result.returncode != 0:
                    continue
                succeeded = True
                text = result.stdout.decode("utf-8", errors="replace")
                words: list[tuple] = []
                for row in csv.reader(io.StringIO(text), delimiter="\t"):
                    if len(row) < 12 or row[0] != "5":
                        continue
                    try:
                        conf = float(row[10])
                        value = row[11].strip()
                        if conf < 30 or not value:
                            continue
                        left, top, width, height = (
                            float(row[6]), float(row[7]), float(row[8]),
                            float(row[9]))
                        words.append((left, top, left + width, top + height,
                                      value, conf))
                    except (TypeError, ValueError):
                        continue
                if len(words) > len(best):
                    best = words
            except Exception:
                continue
    return succeeded, best


def convert_image_to_text(input_path: str,
                          output_path: str) -> tuple[bool, str]:
    """Görseli .txt metne çevirir. Dönüş: (ok, bilgi)."""
    from converter.images import _IMG_EXTS
    if os.path.splitext(input_path)[1].lower() not in _IMG_EXTS:
        return False, "Görsel dosyası değil."
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    ok, text = ocr_image_to_text(input_path)
    if not ok:
        return False, text
    if not text.strip():
        return False, "Görselde metin bulunamadı."
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(text + "\n")
        return True, os.path.basename(output_path)
    except Exception as e:
        return False, f"Yazma hatası: {e}"


def ocr_pdf_page_words(page, zoom: float = 2.0) -> list[tuple]:
    """Render edilen PDF sayfasından OCR kelime kutularını çıkarır."""
    if not tesseract_available():
        return []
    tmp = ""
    bitmap = None
    image = None
    try:
        bitmap = page.render(scale=zoom)
        image = bitmap.to_pil()
        fd, tmp = tempfile.mkstemp(prefix="ocr_words_", suffix=".png")
        os.close(fd)
        image.save(tmp)
        ok, pixels = ocr_image_to_words(tmp)
        if not ok:
            return []
        scale = 1.0 / max(0.1, zoom)
        return [(x0 * scale, y0 * scale, x1 * scale, y1 * scale,
                 text, conf) for x0, y0, x1, y1, text, conf in pixels]
    except Exception:
        return []
    finally:
        try:
            if image is not None:
                image.close()
        except Exception:
            pass
        try:
            if bitmap is not None:
                bitmap.close()
        except Exception:
            pass
        try:
            if tmp and os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass


def ocr_pdf_page(page, zoom: float = 2.0) -> str:
    """PDFium sayfasını render edip OCR'lar. Başarısızsa ''."""
    if not tesseract_available():
        return ""
    tmp = ""
    bitmap = None
    image = None
    try:
        bitmap = page.render(scale=zoom)
        image = bitmap.to_pil()
        fd, tmp = tempfile.mkstemp(prefix="ocr_", suffix=".png")
        os.close(fd)
        image.save(tmp)
        ok, text = ocr_image_to_text(tmp)
        return text if ok else ""
    except Exception:
        return ""
    finally:
        try:
            if image is not None:
                image.close()
        except Exception:
            pass
        try:
            if bitmap is not None:
                bitmap.close()
        except Exception:
            pass
        try:
            if tmp and os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass


def ocr_pdf_page_text(page, zoom: float = 2.0) -> str:
    """PDFium sayfası için düz OCR metni."""
    return ocr_pdf_page(page, zoom)
