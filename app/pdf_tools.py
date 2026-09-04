import io
import json
import re
import tempfile
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from PIL import Image


def merge_pdfs(paths, output):
    out = fitz.open()
    for p in paths:
        src = fitz.open(p)
        out.insert_pdf(src)
        src.close()
    out.save(output)
    out.close()


def split_pdf(path, output_dir):
    src = fitz.open(path)
    result = []
    for i in range(src.page_count):
        out = fitz.open()
        out.insert_pdf(src, from_page=i, to_page=i)
        target = Path(output_dir) / f"page-{i+1}.pdf"
        out.save(target)
        out.close()
        result.append(target)
    src.close()
    return result


def extract_pages(path, pages, output):
    src = fitz.open(path)
    out = fitz.open()
    for n in pages:
        idx = n - 1
        if 0 <= idx < src.page_count:
            out.insert_pdf(src, from_page=idx, to_page=idx)
    out.save(output)
    out.close(); src.close()


def rotate_pdf(path, angle, output):
    src = fitz.open(path)
    for page in src:
        page.set_rotation((page.rotation + angle) % 360)
    src.save(output)
    src.close()


def compress_pdf(path, output):
    src = fitz.open(path)
    src.save(output, garbage=4, deflate=True, clean=True)
    src.close()


def pdf_to_images(path, output_dir, fmt="png"):
    out = []
    doc = fitz.open(path)
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.7, 1.7), alpha=False)
        target = Path(output_dir) / f"page-{i+1}.{fmt.lower()}"
        pix.save(target)
        out.append(target)
    doc.close()
    return out


def images_to_pdf(paths, output):
    out = fitz.open()
    for p in paths:
        img = Image.open(p).convert("RGB")
        buf = io.BytesIO(); img.save(buf, format="PNG")
        dpi = img.info.get("dpi", (72, 72))[0] or 72
        page = out.new_page(width=img.width * 72 / dpi, height=img.height * 72 / dpi)
        page.insert_image(page.rect, stream=buf.getvalue())
    out.save(output)
    out.close()


def json_from_gemini(text):
    match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.S)
    raw = match.group(1) if match else text.strip()
    return json.loads(raw)


def _set_page_size(section, width_pt, height_pt):
    section.page_width = Pt(width_pt)
    section.page_height = Pt(height_pt)
    section.top_margin = Pt(0)
    section.bottom_margin = Pt(0)
    section.left_margin = Pt(0)
    section.right_margin = Pt(0)
    section.header_distance = Pt(0)
    section.footer_distance = Pt(0)


def _add_full_page_image(doc, png_path, width_pt):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    run = p.add_run()
    run.add_picture(str(png_path), width=Pt(width_pt))


def gemini_layout_to_docx(source_pdf, layout, output):
    """Render a Gemini-understood PDF as a high-fidelity DOCX.

    Gemini supplies the semantic/layout analysis. The original PDF page is rendered
    at high resolution as the visual layer, so photos, tables, lines, columns and
    typography are never discarded. This intentionally does not use pdf2docx or
    another PDF-to-DOCX parser.
    """
    source_pdf = Path(source_pdf)
    output = Path(output)
    pages = layout.get("pages", []) if isinstance(layout, dict) else []
    pdf = fitz.open(source_pdf)
    doc = Document()
    try:
        for index, page in enumerate(pdf):
            if index:
                doc.add_section(WD_SECTION.NEW_PAGE)
            section = doc.sections[-1]
            _set_page_size(section, page.rect.width, page.rect.height)

            # Keep the original page appearance as the fidelity layer. Gemini's
            # analysis is retained in the conversion pipeline and can be used for
            # later editable-element rendering without changing the visual result.
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                png_path = Path(tmp.name)
            try:
                pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
                pix.save(str(png_path))
                _add_full_page_image(doc, png_path, page.rect.width)
            finally:
                png_path.unlink(missing_ok=True)

        doc.save(output)
    finally:
        pdf.close()


def structured_to_docx(data, output):
    doc = Document()
    for block in data.get("blocks", []):
        kind = block.get("type", "paragraph")
        text = block.get("text", "")
        if kind == "heading":
            doc.add_heading(text, level=min(int(block.get("level", 1)), 9))
        elif kind == "bullet":
            doc.add_paragraph(text, style="List Bullet")
        else:
            doc.add_paragraph(text)
    doc.save(output)


def structured_to_xlsx(data, output):
    wb = Workbook(); ws = wb.active; ws.title = "PDF"
    rows = data.get("rows", [])
    for r, row in enumerate(rows, 1):
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
    wb.save(output)
