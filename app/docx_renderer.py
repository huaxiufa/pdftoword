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
from docx.shared import Pt
from lxml import html

EMU_PER_INCH = 914400
PT_PER_INCH = 72


class CoordinateDocxRenderer:
    """Rebuild each PDF page as an editable Word coordinate canvas."""

    def __init__(self, pdf_path: Path):
        self.pdf_path = Path(pdf_path)
        self.pdf = fitz.open(self.pdf_path)

    def render(self, json_files: list[Path], out_path: Path) -> None:
        doc = Document()
        for page_no, json_path in enumerate(json_files):
            page_index = self._page_index(json_path, page_no)
            page = self.pdf[page_index]
            section = doc.sections[0] if page_no == 0 else doc.add_section(WD_SECTION.NEW_PAGE)
            self._configure_section(section, page.rect.width, page.rect.height)
            data = json.loads(json_path.read_text(encoding="utf-8"))
            self._render_page(doc, page, data)
        self.pdf.close()
        doc.save(out_path)

    def _configure_section(self, section, width_pt: float, height_pt: float):
        section.page_width = Pt(width_pt)
        section.page_height = Pt(height_pt)
        section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Pt(0)
        section.header_distance = section.footer_distance = Pt(0)

    def _render_page(self, doc, page, data):
        page_w, page_h = page.rect.width, page.rect.height

        # Preserve original PDF images instead of rasterizing the whole page.
        seen = set()
        for img in page.get_images(full=True):
            xref = img[0]
            for rect in page.get_image_rects(xref):
                key = (xref, round(rect.x0, 2), round(rect.y0, 2), round(rect.x1, 2), round(rect.y1, 2))
                if key in seen:
                    continue
                seen.add(key)
                try:
                    pix = fitz.Pixmap(self.pdf, xref)
                    if pix.alpha:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    self._add_floating_image(doc, pix.tobytes("png"), rect, page_w, page_h)
                except Exception:
                    pass

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
            text = str(block.get("block_content") or "").strip()
            if label == "table" and "<table" in text.lower():
                self._add_table_box(doc, text, bbox, page_w, page_h)
            elif text:
                self._add_text_box(doc, text, bbox, page_w, page_h)

    def _add_text_box(self, doc, text, bbox, page_w, page_h):
        x0, y0, x1, y1 = map(float, bbox[:4])
        width, height = max(1, x1 - x0), max(1, y1 - y0)
        font_pt = max(6, min(36, height * 0.72))
        p = doc.add_paragraph()
        self._zero_paragraph(p)
        pict = OxmlElement("w:pict")
        shape = self._shape(x0, y0, width, height)
        txbx = OxmlElement("v:textbox")
        content = OxmlElement("w:txbxContent")
        wp = OxmlElement("w:p")
        wr = OxmlElement("w:r")
        rp = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(max(2, int(font_pt * 2)))); rp.append(sz)
        wr.append(rp)
        wt = OxmlElement("w:t"); wt.text = text; wr.append(wt)
        wp.append(wr); content.append(wp); txbx.append(content); shape.append(txbx); pict.append(shape)
        p._p.append(pict)

    def _add_table_box(self, doc, table_html, bbox, page_w, page_h):
        root = html.fromstring(table_html)
        table = root if root.tag.lower() == "table" else root.find(".//table")
        if table is None:
            return self._add_text_box(doc, root.text_content().strip(), bbox, page_w, page_h)
        rows = []
        for row in table.xpath(".//tr"):
            cells = row.xpath("./th|./td")
            rows.append([(c.text_content().strip(), c.tag.lower() == "th") for c in cells])
        if not rows:
            return
        x0, y0, x1, y1 = map(float, bbox[:4])
        p = doc.add_paragraph(); self._zero_paragraph(p)
        pict = OxmlElement("w:pict"); shape = self._shape(x0, y0, x1-x0, y1-y0)
        txbx = OxmlElement("v:textbox"); content = OxmlElement("w:txbxContent")
        tbl = OxmlElement("w:tbl"); self._build_table(tbl, rows)
        content.append(tbl); txbx.append(content); shape.append(txbx); pict.append(shape); p._p.append(pict)

    def _build_table(self, tbl, rows):
        tbl_pr = OxmlElement("w:tblPr")
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "4"); el.set(qn("w:space"), "0"); el.set(qn("w:color"), "808080")
            borders.append(el)
        tbl_pr.append(borders); tbl.append(tbl_pr)
        for cells in rows:
            tr = OxmlElement("w:tr")
            for text, header in cells:
                tc = OxmlElement("w:tc"); p = OxmlElement("w:p"); r = OxmlElement("w:r")
                if header: r_pr = OxmlElement("w:rPr"); r_pr.append(OxmlElement("w:b")); r.append(r_pr)
                t = OxmlElement("w:t"); t.text = text; r.append(t); p.append(r); tc.append(p); tr.append(tc)
            tbl.append(tr)

    def _add_floating_image(self, doc, blob, rect, page_w, page_h):
        x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1
        p = doc.add_paragraph(); self._zero_paragraph(p)
        run = p.add_run()
        inline_shape = run.add_picture(io.BytesIO(blob), width=Pt(max(1, x1-x0)), height=Pt(max(1, y1-y0)))
        inline = inline_shape._inline
        drawing = inline
        anchor = OxmlElement("wp:anchor")
        for k, v in {"distT":"0","distB":"0","distL":"0","distR":"0","simplePos":"0","relativeHeight":"0","behindDoc":"0","locked":"0","layoutInCell":"1","allowOverlap":"1"}.items():
            anchor.set(k, v)
        simple = OxmlElement("wp:simplePos"); simple.set("x","0"); simple.set("y","0"); anchor.append(simple)
        for tag, value in (("positionH", x0), ("positionV", y0)):
            pos = OxmlElement(f"wp:{tag}"); pos.set("relativeFrom", "page")
            off = OxmlElement("wp:posOffset"); off.text = str(int(value * EMU_PER_INCH / PT_PER_INCH)); pos.append(off); anchor.append(pos)
        extent = OxmlElement("wp:extent"); extent.set("cx", str(int((x1-x0)*EMU_PER_INCH/PT_PER_INCH))); extent.set("cy", str(int((y1-y0)*EMU_PER_INCH/PT_PER_INCH))); anchor.append(extent)
        eff = OxmlElement("wp:effectExtent")
        for k in ("l","t","r","b"): eff.set(k,"0")
        anchor.append(eff); anchor.append(OxmlElement("wp:wrapNone"))
        for child in list(drawing): anchor.append(child)
        drawing.getparent().replace(drawing, anchor)

    def _shape(self, x, y, width, height):
        shape = OxmlElement("v:shape")
        shape.set(qn("id"), "_x0000_s" + str(abs(hash((x,y,width,height))) % 900000 + 100000))
        shape.set(qn("type"), "#_x0000_t202")
        shape.set(qn("style"), f"position:absolute;margin-left:{x}pt;margin-top:{y}pt;width:{width}pt;height:{height}pt;mso-wrap-style:none")
        shape.set(qn("stroked"), "f"); shape.set(qn("filled"), "f")
        return shape

    def _zero_paragraph(self, p):
        fmt = p.paragraph_format
        fmt.space_before = Pt(0); fmt.space_after = Pt(0); fmt.line_spacing = 1

    def _page_index(self, path: Path, fallback: int) -> int:
        m = re.search(r"page-(\d+)", path.name)
        return int(m.group(1)) if m else fallback
