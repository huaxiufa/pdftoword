from __future__ import annotations

import os
import tempfile
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


EMU_PER_PT = 12700


def _set_page_layout(section, width_pt: float, height_pt: float) -> None:
    section.page_width = Inches(width_pt / 72)
    section.page_height = Inches(height_pt / 72)
    for name in (
        "top_margin",
        "bottom_margin",
        "left_margin",
        "right_margin",
        "header_distance",
        "footer_distance",
    ):
        setattr(section, name, Inches(0))


def _set_paragraph_exact(paragraph) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1
    p_pr = paragraph._p.get_or_add_pPr()
    spacing = p_pr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        p_pr.append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "exact")


def _set_run_exact_size(run, width_pt: float, height_pt: float) -> None:
    drawing = run._element.xpath(".//wp:inline")
    if not drawing:
        return
    inline = drawing[0]
    extent = inline.find(qn("wp:extent"))
    if extent is None:
        extent = OxmlElement("wp:extent")
        inline.insert(0, extent)
    extent.set("cx", str(round(width_pt * EMU_PER_PT)))
    extent.set("cy", str(round(height_pt * EMU_PER_PT)))


def _set_picture_no_wrap(run) -> None:
    inline = run._element.xpath(".//wp:inline")
    if not inline:
        return
    inline = inline[0]
    # Keep the page image inline so Word does not introduce wrapping or
    # floating-position drift. The paragraph itself is exactly zero-height
    # around the image.
    docPr = inline.find(qn("wp:docPr"))
    if docPr is not None:
        docPr.set("descr", "Rendered PDF page - V5 visual fidelity")


def _render_page(pdf_page: fitz.Page, output_png: Path, dpi: int) -> None:
    # Render the complete PDF display list instead of extracting individual
    # images. This is intentional: MuPDF composes transparency masks, clipping,
    # image transforms, vector graphics, fonts, annotations and other PDF
    # drawing operations exactly as a PDF viewer sees them.
    pix = pdf_page.get_pixmap(
        dpi=dpi,
        colorspace=fitz.csRGB,
        alpha=False,
        annots=True,
    )
    pix.save(str(output_png))


def _add_page_image(doc: Document, section, image_path: Path, width_pt: float, height_pt: float) -> None:
    paragraph = doc.add_paragraph()
    _set_paragraph_exact(paragraph)
    paragraph.alignment = 0
    run = paragraph.add_run()
    run.add_picture(
        str(image_path),
        width=Inches(width_pt / 72),
        height=Inches(height_pt / 72),
    )
    _set_run_exact_size(run, width_pt, height_pt)
    _set_picture_no_wrap(run)


def convert_pdf_to_docx(source_pdf: Path, output: Path, progress=None):
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)

    dpi = int(os.getenv("PDF_RENDER_DPI", "220"))
    dpi = max(120, min(300, dpi))

    doc = Document()
    # Remove the default empty paragraph before the first rendered page.
    body = doc._body._element
    for child in list(body):
        if child.tag == qn("w:p"):
            body.remove(child)

    with fitz.open(source_pdf) as pdf, tempfile.TemporaryDirectory(prefix="pdftoword-v5-") as tmp:
        total = len(pdf)
        for index, page in enumerate(pdf):
            if index:
                doc.add_section(WD_SECTION.NEW_PAGE)
            section = doc.sections[-1]
            width_pt = float(page.rect.width)
            height_pt = float(page.rect.height)
            _set_page_layout(section, width_pt, height_pt)

            page_dir = Path(tmp) / f"page-{index + 1:04d}"
            page_dir.mkdir(parents=True, exist_ok=True)
            png = page_dir / "page.png"

            if progress:
                progress({
                    "stage": "render",
                    "percent": 8 + int(index / max(total, 1) * 84),
                    "message": f"V5 正在按 PDF 原始视觉效果渲染第 {index + 1}/{total} 页…",
                })

            _render_page(page, png, dpi)
            _add_page_image(doc, section, png, width_pt, height_pt)

    if progress:
        progress({"stage": "save", "percent": 96, "message": "正在生成高保真 DOCX…"})
    doc.save(output)
