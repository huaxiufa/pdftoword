import io
import json
import re
from pathlib import Path

import fitz
from docx import Document
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
        page = out.new_page(width=img.width * 72 / img.info.get("dpi", (72,72))[0], height=img.height * 72 / img.info.get("dpi", (72,72))[1])
        page.insert_image(page.rect, stream=buf.getvalue())
    out.save(output)
    out.close()


def json_from_gemini(text):
    match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.S)
    raw = match.group(1) if match else text.strip()
    return json.loads(raw)


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
