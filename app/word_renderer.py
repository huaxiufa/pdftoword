from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor


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
    if not isinstance(gem_page, dict):
        return []
    elements = []
    direct = gem_page.get("elements")
    if isinstance(direct, list):
        elements.extend(x for x in direct if isinstance(x, dict))
    columns = gem_page.get("columns")
    if isinstance(columns, list):
        for col in columns:
            if isinstance(col, dict) and isinstance(col.get("elements"), list):
                elements.extend(x for x in col["elements"] if isinstance(x, dict))
    return sorted(elements, key=lambda e: (_safe_float(e.get("y")), _safe_float(e.get("x"))))


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


def _set_cell_text(cell, text, element=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    run = p.add_run("" if text is None else str(text))
    if element:
        run.bold = bool(element.get("bold"))
        run.italic = bool(element.get("italic"))
        run.underline = bool(element.get("underline"))
        run.font.size = Pt(max(7, _safe_float(element.get("font_size"), 9)))
        try:
            run.font.name = str(element.get("font_family") or "Arial")
            run.font.color.rgb = RGBColor.from_string(_color(element.get("color")))
        except Exception:
            pass


def _add_text(cell, element):
    text = str(element.get("text") or "")
    if not text.strip():
        return
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1
    align = str(element.get("align") or "left").lower()
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    run = p.add_run(text)
    run.bold = bool(element.get("bold"))
    run.italic = bool(element.get("italic"))
    run.underline = bool(element.get("underline"))
    run.font.size = Pt(max(6, _safe_float(element.get("font_size"), 10.5)))
    try:
        run.font.name = str(element.get("font_family") or "Arial")
        run.font.color.rgb = RGBColor.from_string(_color(element.get("color")))
    except Exception:
        pass


def _normalize_table(element):
    """Normalize Gemini table output while preserving row/column order and spans."""
    cells = element.get("cells")
    if isinstance(cells, list) and cells:
        max_row = 0
        max_col = 0
        normalized = []
        for item in cells:
            if not isinstance(item, dict):
                continue
            r = max(0, int(_safe_float(item.get("row"), 0)))
            c = max(0, int(_safe_float(item.get("col"), 0)))
            rs = max(1, int(_safe_float(item.get("row_span"), 1)))
            cs = max(1, int(_safe_float(item.get("col_span"), 1)))
            normalized.append((r, c, rs, cs, item.get("text", "")))
            max_row = max(max_row, r + rs)
            max_col = max(max_col, c + cs)
        if normalized and max_col:
            return max_row, max_col, normalized

    rows = element.get("rows")
    if not isinstance(rows, list):
        return 0, 0, []
    clean = []
    for row in rows:
        if isinstance(row, list):
            clean.append(row)
        elif isinstance(row, dict):
            clean.append([row.get("text", "")])
        else:
            clean.append([row])
    cols = max((len(r) for r in clean), default=0)
    cells = []
    for r, row in enumerate(clean):
        for c in range(cols):
            cells.append((r, c, 1, 1, row[c] if c < len(row) else ""))
    return len(clean), cols, cells


def _add_table(cell, element):
    rows, cols, source_cells = _normalize_table(element)
    if not rows or not cols:
        return
    table = cell.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    table.autofit = False
    for r in range(rows):
        for c in range(cols):
            target = table.cell(r, c)
            target.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            target.width = Inches(1.0)
    for r, c, rs, cs, value in source_cells:
        if r >= rows or c >= cols:
            continue
        target = table.cell(r, c)
        if rs > 1 or cs > 1:
            try:
                target = target.merge(table.cell(min(rows - 1, r + rs - 1), min(cols - 1, c + cs - 1)))
            except Exception:
                pass
        _set_cell_text(target, value, element)
    cell.add_paragraph().paragraph_format.space_after = Pt(0)


def _extract_images(page, work_dir):
    """Return images in visual PDF order, with their real placement rectangles."""
    result = []
    seen = set()
    for info in page.get_images(full=True):
        try:
            xref = info[0]
            rects = page.get_image_rects(xref)
            data = page.parent.extract_image(xref)
            ext = data.get("ext", "png")
            path = Path(work_dir) / f"p{page.number + 1}_img{xref}.{ext}"
            path.write_bytes(data["image"])
            if rects:
                for rect in rects:
                    key = (xref, round(rect.x0, 3), round(rect.y0, 3), round(rect.x1, 3), round(rect.y1, 3))
                    if key not in seen:
                        result.append((path, rect))
                        seen.add(key)
        except Exception:
            continue
    result.sort(key=lambda item: (item[1].y0, item[1].x0))
    return {i: item for i, item in enumerate(result)}


def _add_image(cell, image_info, element, page_rect):
    image_path, actual_rect = image_info
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1

    # Use the PDF's actual image rectangle as the source of truth for dimensions.
    # Gemini still decides that this element is an image and where it belongs semantically.
    actual_w_pt = max(1.0, actual_rect.width)
    actual_h_pt = max(1.0, actual_rect.height)
    page_w_pt = max(1.0, page_rect.width)
    page_h_pt = max(1.0, page_rect.height)
    element_w = _safe_float(element.get("w"), 0.0)
    width_pt = actual_w_pt
    if element_w > 0 and abs(element_w - actual_w_pt / page_w_pt) < 0.18:
        width_pt = element_w * page_w_pt
    width_pt = max(1.0, min(width_pt, page_w_pt))

    x = max(0.0, min(actual_rect.x0 / page_w_pt, 1.0))
    if _safe_float(element.get("x"), 0.0) > 0 and abs(_safe_float(element.get("x")) - x) < 0.15:
        x = _safe_float(element.get("x"))
    left_pt = x * page_w_pt
    if left_pt > 0:
        p.paragraph_format.left_indent = Pt(left_pt)

    # Vertical spacing is capped so a bad Gemini coordinate cannot create a huge gap.
    y = actual_rect.y0 / page_h_pt
    if _safe_float(element.get("y"), 0.0) > 0 and abs(_safe_float(element.get("y")) - y) < 0.15:
        y = _safe_float(element.get("y"))
    p.paragraph_format.space_before = Pt(min(max(0.0, y * page_h_pt), 72))

    align = str(element.get("align") or "left").lower()
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    try:
        p.add_run().add_picture(str(image_path), width=Inches(width_pt / 72.0))
    except Exception:
        pass


def _render_elements(cell, elements, image_map, page_rect):
    for element in elements:
        kind = str(element.get("type") or "text").lower()
        if kind in {"text", "heading", "paragraph", "bullet"}:
            item = dict(element)
            if kind == "bullet":
                item["text"] = "• " + str(item.get("text") or "")
            _add_text(cell, item)
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
            try:
                run.font.color.rgb = RGBColor.from_string(_color(element.get("color"), "808080"))
            except Exception:
                pass


def render_editable_pdf(source_pdf, layout, output):
    """Render Gemini's structured understanding into editable Word objects.

    Text is native Word text, tables are native Word tables, and original PDF
    images are inserted as separate pictures. The page itself is never rasterized.
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
            section.top_margin = Pt(12)
            section.bottom_margin = Pt(12)
            section.left_margin = Pt(12)
            section.right_margin = Pt(12)

            gem_page = pages[page_index] if page_index < len(pages) and isinstance(pages[page_index], dict) else {}
            image_map = _extract_images(page, work_dir)
            columns = _column_elements(gem_page)
            if len(columns) > 1:
                layout_table = doc.add_table(rows=1, cols=len(columns))
                layout_table.autofit = True
                for col_index, elements in enumerate(columns):
                    c = layout_table.cell(0, col_index)
                    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                    c.text = ""
                    _render_elements(c, elements, image_map, page.rect)
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
