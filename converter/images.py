"""Görsel dönüştürücüler. Tamamen yerel (Pillow + PDFium).

- Görsel -> Görsel (png/jpg, kalite ayarlı)
- Görseller -> PDF (tek dosyada birleştirme)
- PDF -> Görsel (sayfaları png/jpg çıkarma)

Buluta veri göndermez.
"""
from __future__ import annotations

import os

_IMG_EXTS = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".tif")


def _open_rgb(path: str):
    """EXIF yönü düzeltilmiş RGB görsel + saydamlık beyaz zemine basılır."""
    from PIL import Image, ImageOps
    img = Image.open(path)
    img = ImageOps.exif_transpose(img)
    if img.mode in ("RGBA", "LA", "PA"):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        alpha = img.split()[-1]
        bg.paste(img.convert("RGB"), mask=alpha)
        return bg
    return img.convert("RGB")


def convert_image_to_image(input_path: str, output_path: str,
                           quality: int = 90) -> tuple[bool, str]:
    ext_in = os.path.splitext(input_path)[1].lower()
    ext_out = os.path.splitext(output_path)[1].lower()
    if ext_in not in _IMG_EXTS:
        return False, f"Desteklenmeyen girdi: {ext_in}"
    if ext_out not in (".png", ".jpg", ".jpeg"):
        return False, f"Desteklenmeyen çıktı: {ext_out}"
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    try:
        img = _open_rgb(input_path)
        if ext_out in (".jpg", ".jpeg"):
            img.save(output_path, "JPEG", quality=max(10, min(100, quality)),
                     optimize=True)
        else:
            img.save(output_path, "PNG", optimize=True)
        return True, os.path.basename(output_path)
    except Exception as e:
        return False, f"Görsel hatası: {e}"


def convert_images_to_pdf(inputs: list[str], output_path: str) -> tuple[bool, str]:
    ins = [p for p in inputs
           if os.path.splitext(p)[1].lower() in _IMG_EXTS
           and os.path.exists(p)]
    if not ins:
        return False, "Çevrilecek görsel bulunamadı."
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    try:
        imgs = [_open_rgb(p) for p in ins]
        imgs[0].save(output_path, "PDF", save_all=True,
                     append_images=imgs[1:] if len(imgs) > 1 else [])
        return True, f"{len(imgs)} görsel birleştirildi."
    except Exception as e:
        return False, f"PDF hatası: {e}"


def convert_pdf_to_images(input_path: str, outdir: str, fmt: str = "png",
                          dpi: int = 150, quality: int = 90,
                          out_names: list[str] | None = None) -> tuple[bool, object]:
    """Sayfaları görsel çıkarır. Dönüş: (ok, [dosya...] veya hata yazısı).

    out_names verilir ve sayfa sayısıyla eşleşirse aynen kullanılır
    (çakışma önceden çözülmüş kabul edilir).
    """
    if os.path.splitext(input_path)[1].lower() != ".pdf":
        return False, "Sadece .pdf destekleniyor."
    fmt = fmt.lower().lstrip(".")
    if fmt == "jpeg":
        fmt = "jpg"
    if fmt not in ("png", "jpg"):
        return False, f"Desteklenmeyen çıktı: {fmt}"
    os.makedirs(outdir, exist_ok=True)
    try:
        import pypdfium2 as pdfium
    except Exception as e:
        return False, f"Gerekli kütüphane yüklenemedi: {e}"
    try:
        src = pdfium.PdfDocument(input_path)
    except Exception as e:
        return False, f"PDF açılamadı: {e}"
    outs: list[str] = []
    try:
        base = os.path.splitext(os.path.basename(input_path))[0]
        scale = max(72, min(300, dpi)) / 72.0
        use_names = (list(out_names) if out_names
                     and len(out_names) == len(src) else None)
        for i in range(len(src)):
            page = src[i]
            bitmap = None
            image = None
            try:
                bitmap = page.render(scale=scale)
                image = bitmap.to_pil().convert("RGB")
                if use_names:
                    dst = use_names[i]
                else:
                    dst = os.path.join(outdir, f"{base}_p{i + 1}.{fmt}")
                    if os.path.exists(dst):
                        root, ext = os.path.splitext(dst)
                        k = 2
                        while os.path.exists(f"{root}_{k}{ext}"):
                            k += 1
                        dst = f"{root}_{k}{ext}"
                if fmt == "jpg":
                    image.save(dst, "JPEG",
                               quality=max(10, min(100, quality)),
                               optimize=True)
                else:
                    image.save(dst, "PNG", optimize=True)
                outs.append(dst)
            finally:
                if image is not None:
                    image.close()
                if bitmap is not None:
                    bitmap.close()
                page.close()
        return True, outs
    except Exception as e:
        return False, f"Dönüştürme hatası: {e}"
    finally:
        src.close()
