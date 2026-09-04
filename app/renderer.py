from io import BytesIO
from pathlib import Path
import re

from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

from app.model import DocumentModel


EMU_PER_PT = 12700


def _style_run(run, span, font_scale=1.0):
    font_name = span.font or "Arial"
    run.font.name = font_name
    run.font.size = Pt((span.size or 10) * font_scale)
    run.bold = span.bold
    run.italic = span.italic
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is not None:
        rfonts.set(qn("w:eastAsia"), font_name)


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


def _new_paragraph(doc, block, cursor_y, vertical_scale=1.0):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(max(0, block.bbox.x0))
    gap = max(0, block.bbox.y0 - cursor_y) * vertical_scale
    if gap:
        p.paragraph_format.space_before = Pt(gap)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    return p


def _anchor_image_paragraph(doc, block):
    """Add a page-relative floating image directly under w:body.

    The previous renderer put the image in normal inline flow. That makes the
    paragraph's position depend on preceding text and, in particular, prevents
    faithful placement when the PDF image starts above/left of the page. Keep
    the complete source image and use the PDF rectangle itself as the Word
    anchor coordinate instead of cropping or clamping it.
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1

    run = p.add_run()
    inline = run.add_picture(
        BytesIO(block.image),
        width=Inches(max(0.01, block.bbox.width / 72)),
        height=Inches(max(0.01, block.bbox.height / 72)),
    )
    drawing = inline._inline

    anchor = OxmlElement("wp:anchor")
    for name, value in (
        ("distT", "0"),
        ("distB", "0"),
        ("distL", "0"),
        ("distR", "0"),
        ("simplePos", "0"),
        ("relativeHeight", "251658240"),
        ("behindDoc", "0"),
        ("locked", "0"),
        ("layoutInCell", "0"),
        ("allowOverlap", "1"),
    ):
        anchor.set(name, value)

    simple = OxmlElement("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    anchor.append(simple)

    pos_h = OxmlElement("wp:positionH")
    pos_h.set("relativeFrom", "page")
    off_h = OxmlElement("wp:posOffset")
    off_h.text = str(round(block.bbox.x0 * EMU_PER_PT))
    pos_h.append(off_h)
    anchor.append(pos_h)

    pos_v = OxmlElement("wp:positionV")
    pos_v.set("relativeFrom", "page")
    off_v = OxmlElement("wp:posOffset")
    off_v.text = str(round(block.bbox.y0 * EMU_PER_PT))
    pos_v.append(off_v)
    anchor.append(pos_v)

    for child in list(drawing):
        anchor.append(child)
    anchor.append(OxmlElement("wp:wrapNone"))
    drawing.getparent().replace(drawing, anchor)
    return p


def render_docx(model: DocumentModel, output: Path, font_scale=1.0, vertical_scale=1.0):
    """Render editable DOCX while preserving PDF page-relative image geometry."""
    doc = Document()

    for page_index, page in enumerate(model.pages):
        if page_index:
            section = doc.add_section(WD_SECTION.NEW_PAGE)
        else:
            section = doc.sections[0]
        _configure_section(section, page)
        cursor_y = 0.0

        for block in page.blocks:
            if block.kind == "image" and block.image:
                _anchor_image_paragraph(doc, block)
                cursor_y = max(cursor_y, block.bbox.y1)
                continue

            p = _new_paragraph(doc, block, cursor_y, vertical_scale)

            if block.kind == "heading":
                for line in block.lines:
                    for span in line.spans:
                        r = p.add_run(span.text)
                        _style_run(r, span, font_scale)
                        r.bold = True
                cursor_y = max(cursor_y, block.bbox.y1)
                continue

            if block.kind == "list":
                p.add_run("• ")
                for line in block.lines:
                    for span in line.spans:
                        r = p.add_run(_clean_bullet(span.text))
                        _style_run(r, span, font_scale)
                cursor_y = max(cursor_y, block.bbox.y1)
                continue

            for i, line in enumerate(block.lines):
                if i:
                    p.add_run().add_break()
                for span in line.spans:
                    r = p.add_run(span.text)
                    _style_run(r, span, font_scale)
            cursor_y = max(cursor_y, block.bbox.y1)

    doc.save(output)
