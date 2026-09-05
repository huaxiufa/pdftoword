from __future__ import annotations

import io
import json
import re
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from lxml import html

EMU_PER_INCH = 914400
PT_PER_INCH = 72


class CoordinateDocxRenderer:
    """Rebuild each PDF page as an editable Word canvas.

    Text is emitted as editable Word text boxes positioned from PaddleOCR's
    layout coordinates. PDF embedded images are inserted as floating images
    at their original page rectangles. Tables are reconstructed as editable
    Word tables inside positioned VML text boxes when table HTML is available.
    """

    def __init__(self, pdf_path: Path):
        self.pdf_path = Path(pdf_path)
        self.pdf = fitz.open(self.pdf_path)

    def render(self, json_files: list[Path], out_path: Path) -> None:
        doc = Document()
        # Remove the initial empty section content; we create one section/page.
        for section_index, json_path in enumerate(json_files):
            page_index = self._page_index(json_path, section_index)
            page = self.pdf[page_index]
            if section_index == 0:
                section = doc.sections[0]
            else:
                section = doc.add_section(WD_SECTION.NEW_PAGE)
            self._configure_section(section, page.rect.width, page.rect.height)
            self._clear_body_page(doc, section)
            data = json.loads(json_path.read_text(encoding="utf-8"))
            self._render_page(doc, page, data)
        self.pdf.close()
        doc.save(out_path)

    def _configure_section(self, section, width_pt: float, height_pt: float):
        section.page_width = Pt(width_pt)
        section.page_height = Pt(height_pt)
        section.top_margin = Pt(0)
        section.bottom_margin = Pt(0)
        section.left_margin = Pt(0)
        section.right_margin = Pt(0)
        section.header_distance = Pt(0)
        section.footer_distance = Pt(0)

    def _clear_body_page(self, doc, section):
        # A zero-size paragraph acts as the anchor for the floating shapes.
        body = section._sectPr.getparent()
        # Remove only body paragraphs/tables created by the default section.
        for child in list(body):
            if child.tag in (qn("w:p"), qn("w:tbl")):
                body.remove(child)
        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = " "
        r.append(t)
        p.append(r)
        body.insert(0, p)

    def _render_page(self, doc, page, data):
        page_w, page_h = page.rect.width, page.rect.height
        # 1. Original PDF images: preserve pixels and coordinates.
        for img in page.get_images(full=True):
            xref = img[0]
            for rect in page.get_image_rects(xref):
                try:
                    pix = fitz.Pixmap(self.pdf, xref)
                    if pix.alpha:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    blob = pix.tobytes("png")
                    self._add_floating_image(doc, blob, rect, page_w, page_h)
                except Exception:
                    continue

        # 2. OCR/layout objects. Text and table content remains editable.
        blocks = data.get("parsing_res_list") or data.get("layout_parsing_results") or []
        if isinstance(blocks, dict):
            blocks = blocks.get("parsing_res_list", [])
        for block in blocks:
            bbox = block.get("block_bbox") or block.get("bbox")
            if not bbox or len(bbox) < 4:
                continue
            label = str(block.get("block_label") or block.get("sub_label") or "text").lower()
            if label in {"image", "figure", "chart"}:
                continue
            if label == "table":
                table_html = block.get("block_content")
                if isinstance(table_html, str) and "<table" in table_html.lower():
                    self._add_table_box(doc, table_html, bbox, page_w, page_h)
                else:
                    text = str(table_html or "").strip()
                    if text:
                        self._add_text_box(doc, text, bbox, page_w, page_h)
                continue
            text = str(block.get("block_content") or "").strip()
            if text:
                self._add_text_box(doc, text, bbox, page_w, page_h)

    def _add_text_box(self, doc, text, bbox, page_w, page_h):
        x0, y0, x1, y1 = map(float, bbox[:4])
        width = max(1.0, x1 - x0)
        height = max(1.0, y1 - y0)
        font_pt = max(6.0, min(36.0, height * 0.72))
        self._add_vml_box(doc, x0, y0, width, height, page_w, page_h,
                           paragraphs=[(text, font_pt, False)])

    def _add_table_box(self, doc, table_html, bbox, page_w, page_h):
        x0, y0, x1, y1 = map(float, bbox[:4])
        width = max(1.0, x1 - x0)
        height = max(1.0, y1 - y0)
        root = html.fromstring(table_html)
        table = root if root.tag.lower() == "table" else root.find(".//table")
        if table is None:
            self._add_text_box(doc, re.sub(r"<[^>]+>", " ", table_html), bbox, page_w, page_h)
            return
        rows = table.xpath(".//tr")
        payload = []
        for row in rows:
            cells = row.xpath("./th|./td")
            payload.append([(c.text_content().strip(), c.tag.lower() == "th") for c in cells])
        if not payload:
            return
        # Render table into a VML textbox using a real Word table.
        p = self._anchor_paragraph(doc)
        pict = OxmlElement("w:pict")
        shape = self._shape(x0, y0, width, height, page_w, page_h)
        txbx = OxmlElement("v:textbox")
        txbx.set(qn("style"), "mso-fit-shape-to-text:t")
        content = OxmlElement("w:txbxContent")
        tbl = OxmlElement("w:tbl")
        self._build_table(tbl, payload)
        content.append(tbl)
        txbx.append(content)
        shape.append(txbx)
        pict.append(shape)
        p._p.append(pict)

    def _build_table(self, tbl, rows):
        tblPr = OxmlElement("w:tblPr")
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "808080")
            borders.append(el)
        tblPr.append(borders)
        tbl.append(tblPr)
        for row in rows:
            tr = OxmlElement("w:tr")
            for text, header in row:
                tc = OxmlElement("w:tc")
                p = OxmlElement("w:p")
                r = OxmlElement("w:r")
                rPr = OxmlElement("w:rPr")
                if header:
                    b = OxmlElement("w:b")
                    rPr.append(b)
                r.append(rPr)
                t = OxmlElement("w:t")
                t.text = text
                r.append(t)
                p.append(r)
                tc.append(p)
                tr.append(tc)
            tbl.append(tr)

    def _add_floating_image(self, doc, blob, rect, page_w, page_h):
        x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1
        p = self._anchor_paragraph(doc)
        run = p.add_run()
        inline = run.add_picture(io.BytesIO(blob), width=Pt(max(1, x1 - x0)), height=Pt(max(1, y1 - y0)))
        inline._inline.getparent().remove(inline._inline)
        drawing = inline._inline
        self._convert_inline_to_anchor(drawing, x0, y0, x1 - x0, y1 - y0, page_w, page_h)
        run._r.append(drawing)

    def _convert_inline_to_anchor(self, inline, x, y, width, height, page_w, page_h):
        anchor = OxmlElement("wp:anchor")
        anchor.set("distT", "0"); anchor.set("distB", "0"); anchor.set("distL", "0"); anchor.set("distR", "0")
        anchor.set("simplePos", "0"); anchor.set("relativeHeight", "0"); anchor.set("behindDoc", "0")
        anchor.set("locked", "0"); anchor.set("layoutInCell", "1"); anchor.set("allowOverlap", "1")
        simple = OxmlElement("wp:simplePos"); simple.set("x", "0"); simple.set("y", "0"); anchor.append(simple)
        for axis, value in (("x", x), ("y", y)):
            pos = OxmlElement(f"wp:position{axis.upper()}")
            pos.set("relativeFrom", "page")
            off = OxmlElement("wp:posOffset"); off.text = str(int(value * EMU_PER_INCH / PT_PER_INCH)); pos.append(off)
            anchor.append(pos)
        extent = OxmlElement("wp:extent")
        extent.set("cx", str(int(width * EMU_PER_INCH / PT_PER_INCH)))
        extent.set("cy", str(int(height * EMU_PER_INCH / PT_PER_INCH)))
        anchor.append(extent)
        effect = OxmlElement("wp:effectExtent")
        for k in ("l", "t", "r", "b"): effect.set(k, "0")
        anchor.append(effect)
        wrap = OxmlElement("wp:wrapNone"); anchor.append(wrap)
        for child in list(inline):
            anchor.append(child)
        inline.getparent().replace(inline, anchor)

    def _add_vml_box(self, doc, x, y, width, height, page_w, page_h, paragraphs):
        p = self._anchor_paragraph(doc)
        pict = OxmlElement("w:pict")
        shape = self._shape(x, y, width, height, page_w, page_h)
        txbx = OxmlElement("v:textbox")
        content = OxmlElement("w:txbxContent")
        for text, font_pt, bold in paragraphs:
            wp = OxmlElement("w:p")
            wr = OxmlElement("w:r")
            rp = OxmlElement("w:rPr")
            sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(max(2, int(font_pt * 2))))
            rp.append(sz)
            if bold: rp.append(OxmlElement("w:b"))
            wr.append(rp)
            wt = OxmlElement("w:t"); wt.text = text; wr.append(wt)
            wp.append(wr); content.append(wp)
        txbx.append(content); shape.append(txbx); pict.append(shape); p._p.append(pict)

    def _shape(self, x, y, width, height, page_w, page_h):
        shape = OxmlElement("v:shape")
        shape.set(qn("id"), "_x0000_s" + str(abs(hash((x, y, width, height))) % 1000000))
        shape.set(qn("type"), "#_x0000_t202")
        shape.set(qn("style"), f"position:absolute;margin-left:{x}pt;margin-top:{y}pt;width:{width}pt;height:{height}pt;z-index:1;mso-wrap-style:none")
        shape.set(qn("stroked"), "f")
        shape.set(qn("filled"), "f")
        return shape

    def _anchor_paragraph(self, doc):
        body = doc.sections[-1]._sectPr.getparent()
        p = OxmlElement("w:p")
        body.append(p)
        return type("P", (), {"_p": p})()

    def _page_index(self, path: Path, fallback: int) -> int:
        m = re.search(r"page-(\d+)", path.name)
        return int(m.group(1)) if m else fallback
