from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _color(value, default="000000"):
    if not isinstance(value, str):
        return default
    value = value.strip().lstrip("#")
    return value.upper() if len(value) == 6 else default


def _page_elements(gem_page):
    elements = []
    if not isinstance(gem_page, dict):
        return elements
    direct = gem_page.get("elements")
    if isinstance(direct, list):
        elements.extend(x for x in direct if isinstance(x, dict))
    columns = gem_page.get("columns")
    if isinstance(columns, list):
        for col in columns:
            if isinstance(col, dict) and isinstance(col.get("elements"), list):
                elements.extend(x for x in col["elements"] if isinstance(x, dict))
    elements.sort(key=lambda e: (_safe_float(e.get("y")), _safe_float(e.get("x"))))
    return elements


def _column_elements(gem_page):
    columns = gem_page.get("columns") if isinstance(gem_page, dict) else None
    if not isinstance(columns, list) or not columns:
        return [_page_elements(gem_page)]
    result = []
    for col in columns:
        if not isinstance(col, dict):
            continue
        items = col.get("elements", [])
        if isinstance(items, list):
            result.append(sorted((x for x in items if isinstance(x, dict)), key=lambda e: (_safe_float(e.get("y")), _safe_float(e.get("x")))))
    return result or [_page_elements(gem_page)]


def _add_text(cell, element):
    text = str(element.get("text") or "")
    if not text.strip():
        return
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    align = str(element.get("align") or "left").lower()
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    run = p.add_run(text)
    run.bold = bool(element.get("bold"))
    run.italic = bool(element.get("italic"))
    run.underline = bool(element.get("underline"))
    run.font.size = Pt(max(6, _safe_float(element.get("font_size"), 10.5)))
    try:
        run.font.name = str(element.get("font_family") or "Arial")
    except Exception:
        pass
    try:
        run.font.color.rgb = __import__("docx").shared.RGBColor.from_string(_color(element.get("color")))
    except Exception:
        pass


def _add_table(cell, element):
    rows = element.get("rows")
    if not isinstance(rows, list) or not rows:
        return
    clean_rows = []
    for row in rows:
        clean_rows.append(row if isinstance(row, list) else [row])
    cols = max((len(r) for r in clean_rows), default=0)
    if not cols:
        return
    table = cell.add_table(rows=len(clean_rows), cols=cols)
    table.style = "Table Grid"
    table.autofit = True
    for r, row in enumerate(clean_rows):
        for c in range(cols):
            value = row[c] if c < len(row) else ""
            target = table.cell(r, c)
            target.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            target.text = "" if value is None else str(value)
            for p in target.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                for run in p.runs:
                    run.font.size = Pt(max(7, _safe_float(element.get("font_size"), 9)))
    cell.add_paragraph().paragraph_format.space_after = Pt(0)


def _extract_images(page, work_dir):
    """Extract page images and retain their real PDF bounding boxes.

    Gemini's image_index follows the PDF's top-to-bottom, left-to-right image
    order. Keeping the actual bbox lets the renderer correct Gemini's occasional
    size estimate errors while still using Gemini for semantic/layout decisions.
    """
    result = []
    for idx, info in enumerate(page.get_images(full=True)):
        try:
            xref = info[0]
            data = page.parent.extract_image(xref)
            ext = data.get("ext", "png")
            path = Path(work_dir) / f"p{page.number + 1}_img{idx}.{ext}"
            path.write_bytes(data["image"])
            rects = page.get_image_rects(xref)
            rect = rects[0] if rects else None
            result.append((idx, path, rect))
        except Exception:
            continue
    # Match the prompt's visual order rather than PyMuPDF object order.
    result.sort(key=lambda item: (
        item[2].y0 if item[2] is not None else 10**9,
        item[2].x0 if item[2] is not None else 10**9,
    ))
    return {order: (path, rect) for order, (_, path, rect) in enumerate(result)}


def _add_image(cell, image_info, element, page_rect):
    image_path, actual_rect = image_info
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)

    # Gemini supplies normalized coordinates; use them for horizontal placement
    # and vertical spacing. This is intentionally a flow-based placement so the
    # resulting DOCX remains editable and stable across Word versions.
    x = _safe_float(element.get("x"), 0.0)
    y = _safe_float(element.get("y"), 0.0)
    w = _safe_float(element.get("w"), 0.0)

    if actual_rect is not None:
        actual_w = max(1.0, actual_rect.width / page_rect.width)
        actual_h = max(1.0, actual_rect.height / page_rect.height)
        # Prefer the real PDF image width when Gemini's estimate is obviously off.
        if w <= 0 or abs(w - actual_w) > 0.08:
            w = actual_w
        # If Gemini has no usable vertical coordinate, use the PDF bbox.
        if y <= 0:
            y = max(0.0, actual_rect.y0 / page_rect.height)
        if x <= 0:
            x = max(0.0, actual_rect.x0 / page_rect.width)
    else:
        actual_h = 0.0

    w = min(max(w or 0.25, 0.03), 0.95)
    page_width = max(1.0, page_rect.width)
    left_pt = max(0.0, min(x * page_width, page_width * 0.85))
    if left_pt:
        p.paragraph_format.left_indent = Pt(left_pt)

    # Keep a small amount of the PDF's vertical position in the flow. Do not add
    # huge blank areas: Word's normal paragraph flow should remain usable.
    if y > 0:
        p.paragraph_format.space_before = Pt(min(y * page_rect.height, 48))

    align = str(element.get("align") or "left").lower()
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.LEFT)

    try:
        p.add_run().add_picture(str(image_path), width=Inches((w * page_width) / 72.0))
    except Exception:
        pass


def _render_elements(cell, elements, image_map, page_rect):
    for element in elements:
        kind = str(element.get("type") or "text").lower()
        if kind in {"text", "heading", "paragraph", "bullet"}:
            if kind == "bullet":
                element = dict(element)
                element["text"] = "• " + str(element.get("text") or "")
            _add_text(cell, element)
        elif kind == "table":
            _add_table(cell, element)
        elif kind == "image":
            idx = int(_safe_float(element.get("image_index"), 0))
            if idx in image_map:
                _add_image(cell, image_map[idx], element, page_rect)
        elif kind == "line":
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run("────────────────────────────────")
            run.font.size = Pt(6)
            run.font.color.rgb = __import__("docx").shared.RGBColor.from_string(_color(element.get("color"), "808080"))


def render_editable_pdf(source_pdf, layout, output):
    """Render Gemini's PDF understanding into an editable DOCX.

    Text and tables stay native Word objects; PDF images are extracted as separate
    images. Image dimensions and approximate placement are corrected from the
    source PDF's real image bounding boxes rather than relying only on Gemini's
    visual size estimate.
    """
    source_pdf = Path(source_pdf)
    output = Path(output)
    pages = layout.get("pages", []) if isinstance(layout, dict) else []
    if not isinstance(pages, list):
        pages = []
    pdf = fitz.open(source_pdf)
    doc = Document()
    work_dir = output.parent / f".pdf-images-{output.stem}"
    work_dir.mkdir(parents=True, exist_ok=True)
    try:
        for page_index, page in enumerate(pdf):
            if page_index:
                doc.add_section(WD_SECTION.NEW_PAGE)
            section = doc.sections[-1]
            section.page_width = Pt(page.rect.width)
            section.page_height = Pt(page.rect.height)
            section.top_margin = Pt(24)
            section.bottom_margin = Pt(24)
            section.left_margin = Pt(24)
            section.right_margin = Pt(24)

            gem_page = pages[page_index] if page_index < len(pages) and isinstance(pages[page_index], dict) else {}
            image_map = _extract_images(page, work_dir)
            columns = _column_elements(gem_page)

            if len(columns) > 1:
                layout_table = doc.add_table(rows=1, cols=len(columns))
                layout_table.autofit = True
                for col_index, elements in enumerate(columns):
                    cell = layout_table.cell(0, col_index)
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                    cell.text = ""
                    _render_elements(cell, elements, image_map, page.rect)
            else:
                target = doc.add_table(rows=1, cols=1).cell(0, 0)
                target.text = ""
                _render_elements(target, columns[0], image_map, page.rect)
        doc.save(output)
    finally:
        pdf.close()
        for p in work_dir.glob("*"):
            try:
                p.unlink()
            except Exception:
                pass
        try:
            work_dir.rmdir()
        except Exception:
            pass
