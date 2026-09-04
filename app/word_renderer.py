from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from PIL import Image


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
        if isinstance(row, list):
            clean_rows.append(row)
        else:
            clean_rows.append([row])
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
    result = []
    for idx, info in enumerate(page.get_images(full=True)):
        try:
            xref = info[0]
            data = page.parent.extract_image(xref)
            ext = data.get("ext", "png")
            path = Path(work_dir) / f"p{page.number + 1}_img{idx}.{ext}"
            path.write_bytes(data["image"])
            result.append((idx, path))
        except Exception:
            continue
    return result


def _add_image(cell, image_path, element, page_width_pt):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    align = str(element.get("align") or "left").lower()
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    width = _safe_float(element.get("w"), 0.25) * page_width_pt
    width = min(max(width, 20), page_width_pt)
    try:
        p.add_run().add_picture(str(image_path), width=Pt(width))
    except Exception:
        pass


def _render_elements(cell, elements, image_map, page_width_pt):
    # Remove the initial empty paragraph's visual impact, but keep it for DOCX validity.
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
                _add_image(cell, image_map[idx], element, page_width_pt)
        elif kind == "line":
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run("────────────────────────────────")
            run.font.size = Pt(6)
            run.font.color.rgb = __import__("docx").shared.RGBColor.from_string(_color(element.get("color"), "808080"))


def render_editable_pdf(source_pdf, layout, output):
    """Render Gemini's PDF understanding into an editable DOCX.

    This renderer intentionally uses only high-level python-docx APIs. It never
    creates VML text boxes or whole-page screenshots, avoiding fragile XML issues.
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
            image_map = dict(_extract_images(page, work_dir))
            columns = _column_elements(gem_page)

            # A one-row table gives us a stable multi-column page structure while
            # keeping every text node and table editable in Word.
            if len(columns) > 1:
                layout_table = doc.add_table(rows=1, cols=len(columns))
                layout_table.autofit = True
                for col_index, elements in enumerate(columns):
                    cell = layout_table.cell(0, col_index)
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                    cell.text = ""
                    _render_elements(cell, elements, image_map, page.rect.width)
            else:
                cell = doc.add_paragraph()._p.getparent()  # unused; kept out of rendering path
                target = doc.add_table(rows=1, cols=1).cell(0, 0)
                target.text = ""
                _render_elements(target, columns[0], image_map, page.rect.width)
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
