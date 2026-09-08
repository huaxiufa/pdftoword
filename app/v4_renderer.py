from __future__ import annotations

import math
import tempfile
from copy import deepcopy
from pathlib import Path

import fitz
from PIL import Image
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

EMU_PER_PT = 12700


def _remove_cell_margins(cell) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = tcMar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tcMar.append(node)
        node.set(qn("w:w"), "0")
        node.set(qn("w:type"), "dxa")


def _remove_table_borders(table) -> None:
    tblPr = table._tbl.tblPr
    borders = tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        node = borders.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            borders.append(node)
        node.set(qn("w:val"), "nil")


def _font_flags(span: dict) -> tuple[bool, bool]:
    flags = int(span.get("flags") or 0)
    return bool(flags & 16), bool(flags & 2)


def _set_run_font(run, span: dict) -> None:
    font = span.get("font") or "Arial"
    size = float(span.get("size") or 10)
    run.font.name = font
    run.font.size = Pt(max(1, size))
    bold, italic = _font_flags(span)
    run.bold = bold
    run.italic = italic
    color = span.get("color")
    if isinstance(color, int):
        run.font.color.rgb = RGBColor((color >> 16) & 255, (color >> 8) & 255, color & 255)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.rFonts
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:eastAsia"), font)


def _image_bytes_from_block(doc: fitz.Document, block: dict) -> bytes | None:
    data = block.get("image")
    if data:
        try:
            if block.get("mask"):
                base = fitz.Pixmap(data)
                mask = fitz.Pixmap(block["mask"])
                pix = fitz.Pixmap(base, mask)
                return pix.tobytes("png")
            return data
        except Exception:
            pass
    xref = int(block.get("xref") or 0)
    if xref:
        try:
            info = doc.extract_image(xref)
            return info["image"]
        except Exception:
            try:
                pix = fitz.Pixmap(doc, xref)
                if pix.alpha or pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                return pix.tobytes("png")
            except Exception:
                pass
    return None


def _crop_image(path: Path, source_rect: fitz.Rect, page_rect: fitz.Rect) -> tuple[Path, fitz.Rect]:
    clipped = source_rect & page_rect
    if clipped.is_empty or clipped == source_rect:
        return path, source_rect
    with Image.open(path) as image:
        image = image.convert("RGBA")
        sx = image.width / max(source_rect.width, 0.001)
        sy = image.height / max(source_rect.height, 0.001)
        box = (
            max(0, round((clipped.x0 - source_rect.x0) * sx)),
            max(0, round((clipped.y0 - source_rect.y0) * sy)),
            min(image.width, round((clipped.x1 - source_rect.x0) * sx)),
            min(image.height, round((clipped.y1 - source_rect.y0) * sy)),
        )
        cropped = image.crop(box)
        target = path.with_name(path.stem + "-clipped.png")
        cropped.save(target, "PNG")
    return target, clipped


def _rotation_from_matrix(matrix) -> int:
    if not matrix:
        return 0
    try:
        angle = math.degrees(math.atan2(float(matrix[1]), float(matrix[0])))
    except Exception:
        return 0
    normalized = angle % 360
    choices = [0, 90, 180, 270]
    return min(choices, key=lambda value: abs(((normalized - value + 180) % 360) - 180))


def _rotate_image(path: Path, angle: int) -> Path:
    if angle == 0:
        return path
    with Image.open(path) as image:
        rotated = image.convert("RGBA").rotate(-angle, expand=True, resample=Image.Resampling.BICUBIC)
        target = path.with_name(path.stem + f"-rot{angle}.png")
        rotated.save(target, "PNG")
    return target


def _add_anchor(cell, image_path: Path, rect: fitz.Rect, page_rect: fitz.Rect) -> None:
    width = max(1.0, rect.width)
    height = max(1.0, rect.height)
    paragraph = cell.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1
    run = paragraph.add_run()
    inline = run.add_picture(str(image_path), width=Inches(width / 72), height=Inches(height / 72))
    drawing = inline._inline

    anchor = OxmlElement("wp:anchor")
    anchor.set("distT", "0")
    anchor.set("distB", "0")
    anchor.set("distL", "0")
    anchor.set("distR", "0")
    anchor.set("behindDoc", "0")
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "1")
    anchor.set("allowOverlap", "1")

    extent = OxmlElement("wp:extent")
    extent.set("cx", str(round(width * EMU_PER_PT)))
    extent.set("cy", str(round(height * EMU_PER_PT)))
    anchor.append(extent)

    simple_pos = OxmlElement("wp:simplePos")
    simple_pos.set("x", "0")
    simple_pos.set("y", "0")
    anchor.append(simple_pos)

    pos_h = OxmlElement("wp:positionH")
    pos_h.set("relativeFrom", "page")
    off_h = OxmlElement("wp:posOffset")
    off_h.text = str(round(rect.x0 * EMU_PER_PT))
    pos_h.append(off_h)
    anchor.append(pos_h)

    pos_v = OxmlElement("wp:positionV")
    pos_v.set("relativeFrom", "page")
    off_v = OxmlElement("wp:posOffset")
    off_v.text = str(round(rect.y0 * EMU_PER_PT))
    pos_v.append(off_v)
    anchor.append(pos_v)

    wrap = OxmlElement("wp:wrapNone")
    anchor.append(wrap)
    for child in list(drawing):
        anchor.append(deepcopy(child))
    drawing.getparent().replace(drawing, anchor)


def _extract_images(page: fitz.Page, work_dir: Path) -> list[dict]:
    work_dir.mkdir(parents=True, exist_ok=True)
    results: list[dict] = []
    seen: set[tuple[int, float, float, float, float]] = set()
    blocks = page.get_text("dict", sort=False).get("blocks", [])
    for block_index, block in enumerate(blocks):
        if block.get("type") != 1 or not block.get("bbox"):
            continue
        rect = fitz.Rect(block["bbox"])
        key = (int(block.get("xref") or 0), round(rect.x0, 2), round(rect.y0, 2), round(rect.x1, 2), round(rect.y1, 2))
        if key in seen:
            continue
        blob = _image_bytes_from_block(page.parent, block)
        if not blob:
            continue
        path = work_dir / f"image-{block_index + 1:04d}.png"
        path.write_bytes(blob)
        clipped_path, clipped_rect = _crop_image(path, rect, page.rect)
        matrix = block.get("transform")
        angle = _rotation_from_matrix(matrix)
        rotated_path = _rotate_image(clipped_path, angle)
        seen.add(key)
        results.append({"path": rotated_path, "rect": clipped_rect, "angle": angle})

    # Fallback for displayed XObjects that are not present in the text page.
    try:
        infos = page.get_image_info(xrefs=True)
    except Exception:
        infos = []
    for index, info in enumerate(infos):
        bbox = info.get("bbox")
        if not bbox:
            continue
        rect = fitz.Rect(bbox)
        xref = int(info.get("xref") or 0)
        key = (xref, round(rect.x0, 2), round(rect.y0, 2), round(rect.x1, 2), round(rect.y1, 2))
        if key in seen:
            continue
        try:
            data = page.parent.extract_image(xref)["image"]
        except Exception:
            continue
        path = work_dir / f"xref-{index + 1:04d}.png"
        path.write_bytes(data)
        clipped_path, clipped_rect = _crop_image(path, rect, page.rect)
        rotated_path = _rotate_image(clipped_path, _rotation_from_matrix(info.get("transform")))
        seen.add(key)
        results.append({"path": rotated_path, "rect": clipped_rect, "angle": _rotation_from_matrix(info.get("transform"))})

    results.sort(key=lambda item: (round(item["rect"].y0, 2), round(item["rect"].x0, 2)))
    return results


def _render_text_block(cell, block: dict, cursor_y: float) -> float:
    bbox = fitz.Rect(block.get("bbox", (0, 0, 0, 0)))
    paragraph = cell.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.left_indent = Pt(max(0, bbox.x0))
    fmt.right_indent = Pt(0)
    fmt.space_before = Pt(max(0, bbox.y0 - cursor_y))
    fmt.space_after = Pt(0)
    fmt.keep_together = True

    lines = block.get("lines") or []
    max_bottom = bbox.y1
    first_line = True
    for line in lines:
        if not first_line:
            paragraph.add_run().add_break()
        first_line = False
        for span in line.get("spans", []):
            text = span.get("text", "")
            if not text:
                continue
            run = paragraph.add_run(text)
            _set_run_font(run, span)
        max_bottom = max(max_bottom, float(fitz.Rect(line.get("bbox", bbox)).y1))
    return max_bottom


def _render_page(doc: Document, page: fitz.Page, page_index: int, work_dir: Path) -> None:
    if page_index:
        doc.add_section(WD_SECTION.NEW_PAGE)
    section = doc.sections[-1]
    section.page_width = Inches(page.rect.width / 72)
    section.page_height = Inches(page.rect.height / 72)
    section.top_margin = Inches(0)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0)
    section.right_margin = Inches(0)
    section.header_distance = Inches(0)
    section.footer_distance = Inches(0)

    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.allow_autofit = False
    table.width = Inches(page.rect.width / 72)
    _remove_table_borders(table)
    cell = table.cell(0, 0)
    cell.width = Inches(page.rect.width / 72)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    _remove_cell_margins(cell)

    blocks = page.get_text("dict", sort=False).get("blocks", [])
    text_blocks = [b for b in blocks if b.get("type") == 0 and b.get("lines")]
    images = _extract_images(page, work_dir)

    objects: list[tuple[float, int, dict]] = []
    for block in text_blocks:
        objects.append((float(fitz.Rect(block["bbox"]).y0), 0, block))
    for image in images:
        objects.append((float(image["rect"].y0), 1, image))
    objects.sort(key=lambda item: (round(item[0], 2), item[1]))

    cursor_y = 0.0
    for _y, kind, obj in objects:
        if kind == 0:
            cursor_y = _render_text_block(cell, obj, cursor_y)
        else:
            rect = obj["rect"]
            _add_anchor(cell, obj["path"], rect, page.rect)
            cursor_y = max(cursor_y, rect.y1)


def convert_pdf_to_docx(source_pdf: Path, output: Path, progress=None) -> None:
    source_pdf = Path(source_pdf)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()

    with fitz.open(source_pdf) as pdf, tempfile.TemporaryDirectory(prefix="pdftoword-v4-") as tmp:
        total = len(pdf)
        for index, page in enumerate(pdf):
            if progress:
                percent = 10 + int((index / max(total, 1)) * 80)
                progress({"stage": "render", "percent": percent, "message": f"V4 正在解析第 {index + 1}/{total} 页…"})
            _render_page(doc, page, index, Path(tmp) / f"page-{index + 1}")

    if progress:
        progress({"stage": "save", "percent": 95, "message": "正在生成 DOCX…"})
    doc.save(output)
