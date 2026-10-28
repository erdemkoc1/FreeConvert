"""Word (.docx) -> PDF dönüştürücü.

Dış program kullanmaz. Tüm dönüşüm python-docx + ReportLab ile yapılır;
uygulamanın tek portable exe içinde çalışmasını korur.
"""
from __future__ import annotations

import io
import os
import sys
from xml.sax.saxutils import escape

from converter._common import pdf_has_content


def _paragraph_markup(paragraph) -> str:
    """DOCX run içeriğini ReportLab güvenli Paragraph markup'ine çevirir."""
    pieces: list[str] = []
    for run in paragraph.runs:
        text = escape(run.text or "").replace("\n", "<br/>")
        if not text:
            continue
        if run.bold:
            text = f"<b>{text}</b>"
        if run.italic:
            text = f"<i>{text}</i>"
        if run.underline:
            text = f"<u>{text}</u>"
        pieces.append(text)
    if pieces:
        return "".join(pieces)
    return escape(paragraph.text or "").replace("\n", "<br/>")


def _paragraph_style(paragraph, styles):
    name = (paragraph.style.name if paragraph.style else "") or ""
    lower = name.lower()
    if "heading 1" in lower or "başlık 1" in lower:
        return styles["DCHeading1"]
    if "heading 2" in lower or "başlık 2" in lower:
        return styles["DCHeading2"]
    if "heading" in lower or "başlık" in lower:
        return styles["DCHeading3"]
    if "caption" in lower or ("başlık" in lower and "resim" in lower):
        return styles["DCCaption"]
    return styles["DCNormal"]


def _has_page_break(paragraph) -> bool:
    """Paragraftaki sayfa sonu işaretini (w:br type=page) algılar."""
    return bool(paragraph._p.xpath('.//w:br[@w:type="page"]'))


def _paragraph_image_blobs(doc, paragraph) -> list[bytes]:
    """Paragraftaki gömülü görsellerin ham baytlarını sırayla döndürür."""
    from docx.oxml.ns import qn
    blobs: list[bytes] = []
    for blip in paragraph._p.xpath(".//a:blip"):
        rel_id = blip.get(qn("r:embed"))
        part = doc.part.related_parts.get(rel_id) if rel_id else None
        if part is not None:
            blobs.append(part.blob)
    return blobs


def _cell_markup(text: str) -> str:
    return escape(text or "").replace("\n", "<br/>") or " "


def _document_order(doc):
    """Paragraf ve tabloları DOCX XML gövde sırasına göre döndürür."""
    paragraphs = {id(p._p): p for p in doc.paragraphs}
    tables = {id(t._tbl): t for t in doc.tables}
    for child in doc.element.body.iterchildren():
        item = paragraphs.get(id(child))
        if item is not None:
            yield ("paragraph", item)
            continue
        item = tables.get(id(child))
        if item is not None:
            yield ("table", item)


def _builtin_docx_to_pdf(input_path: str, output_path: str) -> None:
    """DOCX belgesini ReportLab akışına aktarır."""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        Image as RLImage,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    doc = Document(input_path)

    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    resource_root = getattr(sys, "_MEIPASS", None) or os.path.dirname(
        os.path.dirname(os.path.abspath(__file__)))
    font_dir = os.path.join(resource_root, "assets", "fonts")
    regular = os.path.join(font_dir, "NotoSans-Regular.ttf")
    bold = os.path.join(font_dir, "NotoSans-Bold.ttf")
    italic = os.path.join(font_dir, "NotoSans-Italic.ttf")
    if all(os.path.exists(path) for path in (regular, bold, italic)):
        pdfmetrics.registerFont(TTFont("NotoSans", regular))
        pdfmetrics.registerFont(TTFont("NotoSans-Bold", bold))
        pdfmetrics.registerFont(TTFont("NotoSans-Italic", italic))
        pdfmetrics.registerFontFamily(
            "NotoSans", normal="NotoSans", bold="NotoSans-Bold",
            italic="NotoSans-Italic", boldItalic="NotoSans-Bold")
        font_name, bold_font, italic_font = (
            "NotoSans", "NotoSans-Bold", "NotoSans-Italic")
    else:
        font_name, bold_font, italic_font = (
            "Helvetica", "Helvetica-Bold", "Helvetica-Oblique")

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        "DCNormal", parent=styles["Normal"], fontName=font_name,
        fontSize=10.5, leading=14, alignment=TA_JUSTIFY, spaceAfter=5))
    styles.add(ParagraphStyle(
        "DCHeading1", parent=styles["Heading1"], fontName=bold_font,
        fontSize=16, leading=19, spaceBefore=8, spaceAfter=7))
    styles.add(ParagraphStyle(
        "DCHeading2", parent=styles["Heading2"], fontName=bold_font,
        fontSize=13, leading=16, spaceBefore=6, spaceAfter=5))
    styles.add(ParagraphStyle(
        "DCHeading3", parent=styles["Heading3"], fontName=bold_font,
        fontSize=11, leading=14, spaceBefore=4, spaceAfter=4))
    styles.add(ParagraphStyle(
        "DCCaption", parent=styles["Normal"], fontName=italic_font,
        fontSize=9, leading=11, alignment=TA_CENTER, spaceAfter=5))

    section = doc.sections[0] if doc.sections else None
    if section is not None:
        page_width = float(section.page_width.pt)
        page_height = float(section.page_height.pt)
        left = float(section.left_margin.pt)
        right = float(section.right_margin.pt)
        top = float(section.top_margin.pt)
        bottom = float(section.bottom_margin.pt)
    else:
        page_width, page_height = A4
        left = right = 2 * cm
        top = bottom = 2 * cm

    # A4 dışı veya çok dar belgelerde ReportLab'ın sayfa boyutunu koru.
    page_size = (max(page_width, 72), max(page_height, 72))
    story = []
    for kind, item in _document_order(doc):
        if kind == "paragraph":
            markup = _paragraph_markup(item)
            has_break = _has_page_break(item)
            if has_break and markup in ("", "<br/>"):
                markup = ""
            blobs = _paragraph_image_blobs(doc, item)
            if not markup.strip() and not blobs and not has_break:
                story.append(Spacer(1, 4))
                continue
            # Görsel ve sayfa sonu, paragraftaki konumlarıyla sırayla eklenir.
            for blob in blobs:
                story.append(Spacer(1, 5))
                story.append(RLImage(io.BytesIO(blob), width=14 * cm,
                                     height=10 * cm, kind="proportional"))
            if markup.strip():
                style = _paragraph_style(item, styles)
                if item.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                    style = ParagraphStyle(
                        f"{style.name}_C", parent=style, alignment=TA_CENTER)
                elif item.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                    style = ParagraphStyle(
                        f"{style.name}_R", parent=style, alignment=TA_RIGHT)
                elif item.alignment == WD_ALIGN_PARAGRAPH.LEFT:
                    style = ParagraphStyle(
                        f"{style.name}_L", parent=style, alignment=TA_LEFT)
                story.append(Paragraph(markup, style))
            if has_break:
                story.append(PageBreak())
            continue

        data = []
        for row in item.rows:
            data.append([Paragraph(_cell_markup(cell.text), styles["DCNormal"])
                         for cell in row.cells])
        if not data:
            continue
        try:
            table = Table(data, repeatRows=1, splitByRow=1)
            table.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]))
            story.extend((Spacer(1, 4), table, Spacer(1, 5)))
        except Exception:
            # Bozuk/geniş bir tablo tüm belgeyi düşürmesin.
            for row in item.rows:
                story.append(Paragraph(
                    " | ".join(_cell_markup(c.text) for c in row.cells),
                    styles["DCNormal"]))

    if not story:
        story.append(Paragraph(" ", styles["DCNormal"]))
    metadata = doc.core_properties
    pdf = SimpleDocTemplate(
        output_path, pagesize=page_size,
        leftMargin=left, rightMargin=right,
        topMargin=top, bottomMargin=bottom,
        title=os.path.basename(input_path),
        author=metadata.author or "")
    pdf.build(story)


def convert_word_to_pdf(input_path: str, output_path: str,
                        mode: str = "adaptif") -> tuple[bool, str]:
    """DOCX'i PDF'e çevirir. ``mode`` geriye uyum için yok sayılır."""
    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".doc":
        return False, ".doc eski biçimi açık kaynak motorla desteklenmiyor; .docx kullanın."
    if ext != ".docx":
        return False, "Sadece .docx destekleniyor."
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    try:
        if os.path.exists(output_path):
            os.remove(output_path)
    except Exception:
        pass
    try:
        _builtin_docx_to_pdf(input_path, output_path)
        if not pdf_has_content(output_path):
            return False, "PDF oluşturuldu ancak içerik doğrulanamadı."
        return True, "Dahili offline dönüştürücü ile (tek exe)."
    except Exception as exc:
        return False, f"Dahili dönüştürme hatası: {exc}"
