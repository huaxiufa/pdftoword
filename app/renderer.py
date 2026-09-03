from io import BytesIO
from pathlib import Path
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

from app.model import DocumentModel


def _style_run(run, span):
    run.font.name = span.font or "Arial"
    run.font.size = Pt(span.size or 10)
    run.bold = span.bold
    run.italic = span.italic


def _clean_bullet(text):
    return re.sub(r"^(?:•|·|▪|‣|◦|●|-|–|—)\s*", "", text)


def render_docx(model: DocumentModel, output: Path):
    doc = Document()
    for page_index, page in enumerate(model.pages):
        if page_index:
            doc.add_page_break()
        section = doc.sections[0]
        section.page_width = Inches(page.width / 72)
        section.page_height = Inches(page.height / 72)
        for block in page.blocks:
            if block.kind == "image" and block.image:
                p = doc.add_paragraph()
                run = p.add_run()
                run.add_picture(BytesIO(block.image), width=Inches(block.bbox.width / 72))
                continue

            if block.kind == "heading":
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(4)
                for line in block.lines:
                    for span in line.spans:
                        r = p.add_run(span.text)
                        _style_run(r, span)
                        r.bold = True
                continue

            if block.kind == "list":
                for line in block.lines:
                    p = doc.add_paragraph(style="List Bullet")
                    for span in line.spans:
                        r = p.add_run(_clean_bullet(span.text))
                        _style_run(r, span)
                continue

            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            for i, line in enumerate(block.lines):
                if i:
                    p.add_run().add_break()
                for span in line.spans:
                    r = p.add_run(span.text)
                    _style_run(r, span)
    doc.save(output)
