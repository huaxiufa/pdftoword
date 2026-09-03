from io import BytesIO
from pathlib import Path
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Inches, Pt

from app.model import DocumentModel


def _style_run(run, span):
    run.font.name = span.font or "Arial"
    run.font.size = Pt(span.size or 10)
    run.bold = span.bold
    run.italic = span.italic
    # Keep East Asia font mapping consistent with the Latin font where possible.
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is not None:
        rfonts.set("w:eastAsia", span.font or "Arial")


def _clean_bullet(text):
    return re.sub(r"^(?:•|·|▪|‣|◦|●|-|–|—)\s*", "", text)


def _configure_section(section, page):
    section.page_width = Inches(page.width / 72)
    section.page_height = Inches(page.height / 72)
    section.top_margin = Inches(0)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0)
    section.right_margin = Inches(0)
    section.header_distance = Inches(0)
    section.footer_distance = Inches(0)


def _new_paragraph(doc, block, cursor_y):
    p = doc.add_paragraph()
    # Convert the PDF's absolute x/y into Word paragraph positioning. This is
    # intentionally conservative: text remains editable while approximate
    # coordinates are preserved for visual regression.
    p.paragraph_format.left_indent = Pt(max(0, block.bbox.x0))
    gap = max(0, block.bbox.y0 - cursor_y)
    if gap:
        p.paragraph_format.space_before = Pt(gap)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    return p


def render_docx(model: DocumentModel, output: Path):
    doc = Document()
    cursor_y = 0.0

    for page_index, page in enumerate(model.pages):
        if page_index:
            section = doc.add_section(WD_SECTION.NEW_PAGE)
            cursor_y = 0.0
        else:
            section = doc.sections[0]
        _configure_section(section, page)

        for block in page.blocks:
            p = _new_paragraph(doc, block, cursor_y)

            if block.kind == "image" and block.image:
                run = p.add_run()
                run.add_picture(BytesIO(block.image), width=Inches(block.bbox.width / 72))
                cursor_y = max(cursor_y, block.bbox.y1)
                continue

            if block.kind == "heading":
                for line in block.lines:
                    for span in line.spans:
                        r = p.add_run(span.text)
                        _style_run(r, span)
                        r.bold = True
                cursor_y = max(cursor_y, block.bbox.y1)
                continue

            if block.kind == "list":
                # Keep the bullet as editable text instead of relying on Word's
                # automatic list indentation, which is a major source of drift.
                p.add_run("• ")
                for line in block.lines:
                    for span in line.spans:
                        r = p.add_run(_clean_bullet(span.text))
                        _style_run(r, span)
                cursor_y = max(cursor_y, block.bbox.y1)
                continue

            for i, line in enumerate(block.lines):
                if i:
                    p.add_run().add_break()
                for span in line.spans:
                    r = p.add_run(span.text)
                    _style_run(r, span)
            cursor_y = max(cursor_y, block.bbox.y1)

    doc.save(output)
