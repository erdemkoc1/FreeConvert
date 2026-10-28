"""PDF tablolarını yerel olarak çıkarmak için pdfplumber katmanı.

Bu modül optional bir iyileştirmedir: pdfplumber kurulu değilse
çağıran taraf mevcut düz metin fallback'ine devam eder. Bulut servis
kullanılmaz.
"""
from __future__ import annotations

from typing import Any


# Sıralama önemlidir: önce güçlü çizgili tablolar, sonra metin hizalı
# tablolar. Böylece normal metin kolonları yanlışlıkla tablo sayılmaz.
_SETTINGS: tuple[dict[str, Any], ...] = (
    {},
    {
        "vertical_strategy": "lines",
        "horizontal_strategy": "text",
        "snap_y_tolerance": 4,
        "text_x_tolerance": 2,
        "text_y_tolerance": 3,
    },
    {
        "vertical_strategy": "text",
        "horizontal_strategy": "lines",
        "snap_x_tolerance": 4,
        "text_x_tolerance": 2,
        "text_y_tolerance": 3,
    },
    {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "text_x_tolerance": 2,
        "text_y_tolerance": 3,
        "intersection_tolerance": 5,
    },
)


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def _compact_rows(rows: list[list[Any]]) -> list[list[str]]:
    """Metin stratejisinin ürettiği tamamen boş ayraç satırlarını at."""
    compact: list[list[str]] = []
    for row in rows:
        cells = [_text(cell) for cell in row]
        if any(cells):
            compact.append(cells)
    return compact


def _row_quality(rows: list[list[Any]]) -> tuple[float, int, int]:
    if not rows:
        return -1.0, 0, 0
    width = max(len(row) for row in rows)
    if width < 2 or len(rows) < 2:
        return -1.0, len(rows), width
    normalized = [
        [_text(cell) for cell in row] + [""] * (width - len(row))
        for row in rows
    ]
    nonempty = sum(1 for row in normalized for cell in row if cell)
    text_cells = sum(
        1 for row in normalized for cell in row
        if any(ch.isalnum() for ch in cell)
    )
    # Çok küçük 2x2 kesişimler genellikle iki sütunlu normal metindir.
    minimum_nonempty = 3 if len(rows) == 2 and width == 2 else 2
    if nonempty < minimum_nonempty or text_cells < 2:
        return -1.0, len(rows), width
    density = nonempty / (len(rows) * width)
    if density < 0.25:
        return -1.0, len(rows), width
    column_coverage = sum(
        1 for col in range(width)
        if sum(bool(row[col]) for row in normalized) >= 2
    ) / width
    if column_coverage < 0.5:
        return -1.0, len(rows), width
    # Daha çok satıt/sütun ve hücre doluluğu iyi tabloyu öne çıkarır.
    score = 10.0 + min(len(rows), 20) + min(width, 12) * 0.6
    score += density * 5.0 + column_coverage * 3.0
    return score, len(rows), width


def _bbox_area(bbox: tuple[float, float, float, float]) -> float:
    return max(0.0, bbox[2] - bbox[0]) * max(0.0, bbox[3] - bbox[1])


def _overlap(a: tuple[float, float, float, float],
             b: tuple[float, float, float, float]) -> float:
    x0, y0 = max(a[0], b[0]), max(a[1], b[1])
    x1, y1 = min(a[2], b[2]), min(a[3], b[3])
    if x1 <= x0 or y1 <= y0:
        return 0.0
    return (x1 - x0) * (y1 - y0)


def extract_page_tables(page: Any) -> list[dict[str, Any]]:
    """Bir pdfplumber Page için güvenilir tablo adaylarını döndürür.

    Dönüş sözlükleri: ``bbox``, ``rows``, ``score`` ve ``strategy``.
    pdfplumber'ın Table nesnesi sayfaya bağlı olduğu için dönüşümde yalnızca
    saf Python verileri tutulur.
    """
    candidates: list[dict[str, Any]] = []
    for index, settings in enumerate(_SETTINGS):
        try:
            tables = page.find_tables(settings)
        except Exception:
            continue
        for table in tables:
            try:
                raw_matrix = [
                    [_text(cell) for cell in row]
                    for row in table.extract(
                        x_tolerance=2, y_tolerance=3)
                ]
                rows = _compact_rows(raw_matrix)
                raw_rows = list(table.rows)
                raw_cell_boxes = [
                    [tuple(float(v) for v in cell) if cell is not None else None
                     for cell in getattr(row, "cells", [])]
                    for row in raw_rows
                ]
                cell_boxes = [
                    boxes for matrix_row, boxes in zip(raw_matrix, raw_cell_boxes)
                    if any(matrix_row)
                ]
                bbox = tuple(float(v) for v in table.bbox)
            except Exception:
                continue
            score, row_count, column_count = _row_quality(rows)
            if score < 0:
                continue
            # Metin-stratejisi sonuçları çok kısıtlanır: aksi halde normal
            # iki sütunlu sayfa tablo gibi görünebilir.
            if index >= 2 and (row_count < 3 or column_count < 2
                                or column_count > 12):
                continue
            if index >= 2:
                values = [cell for row in rows for cell in row if cell]
                lengths = sorted(len(cell) for cell in values)
                median_length = (lengths[len(lengths) // 2]
                                 if lengths else 0)
                short_ratio = (sum(1 for length in lengths if length <= 2)
                               / len(lengths)) if lengths else 1.0
                if median_length < 3 or short_ratio > 0.45:
                    continue
            candidates.append({
                "bbox": bbox,
                "rows": rows,
                "cell_boxes": cell_boxes,
                "score": score + (3.0 if index == 0 else 0.0),
                "strategy": "lines" if index in (0, 1) else "text",
            })

    # Aynı bölgeyi farklı stratejilerin birkaç kez üretmesini engelle.
    selected: list[dict[str, Any]] = []
    # Güçlü çizgi tablosu varsa onu önce seç; metin stratejisi bütün sayfayı
    # kapsayıp yanlışlıkla daha yüksek puan alabilir.
    for candidate in sorted(
            candidates,
            key=lambda item: (item["strategy"] == "lines", item["score"]),
            reverse=True):
        duplicate = False
        for current in selected:
            inter = _overlap(candidate["bbox"], current["bbox"])
            smaller = min(_bbox_area(candidate["bbox"]),
                          _bbox_area(current["bbox"]))
            if smaller and inter / smaller >= 0.65:
                duplicate = True
                break
        if not duplicate:
            selected.append(candidate)

    selected.sort(key=lambda item: (item["bbox"][1], item["bbox"][0]))
    return selected
