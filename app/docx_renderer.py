from __future__ import annotations

import io
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

EMU_PER_PT = 12700


class CoordinateDocxRenderer:
    def __init__(self, pdf_path: Path):
        self.pdf = fitz.open(pdf_path)

    def render(self, pages: list[dict], out_path: Path) -> None:
        doc = Document()
        for i, item in enumerate(pages):
            page = self.pdf[item["page_index"]]
            section = doc.sections[0] if i == 0 else doc.add_section(WD_SECTION.NEW_PAGE)
            section.page_width = Pt(page.rect.width)
            section.page_height = Pt(page.rect.height)
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Pt(0)
            section.header_distance = section.footer_distance = Pt(0)
            self._page(doc, page, item)
        self.pdf.close()
        doc.save(out_path)

    def _page(self, doc, page, item):
        # Keep non-background PDF images as real, floating Word images.
        seen = set()
        for image in page.get_images(full=True):
            xref = image[0]
            for rect in page.get_image_rects(xref):
                if rect.width * rect.height >= page.rect.width * page.rect.height * 0.92:
                    continue
                key = (xref, round(rect.x0, 1), round(rect.y0, 1), round(rect.x1, 1), round(rect.y1, 1))
                if key in seen:
                    continue
                seen.add(key)
                try:
                    pix = fitz.Pixmap(self.pdf, xref)
                    if pix.alpha:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    self._image(doc, pix.tobytes("png"), rect.x0, rect.y0, rect.width, rect.height)
                except Exception:
                    pass

        for region in item.get("regions", []):
            x0, y0, x1, y1 = region["bbox"]
            for line in region.get("lines", []):
                text = str(line.get("text", "")).strip()
                if not text:
                    continue
                bx = line.get("bbox") or region["bbox"]
                self._text(doc, text, bx)

    def _text(self, doc, text, bbox):
        x0, y0, x1, y1 = map(float, bbox[:4])
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1
        pict = OxmlElement("w:pict")
        shape = OxmlElement("v:shape")
        shape.set(qn("id"), f"_x0000_s{abs(hash((x0,y0,x1,y1,text))) % 800000 + 100000}")
        shape.set(qn("type"), "#_x0000_t202")
        shape.set("style", f"position:absolute;margin-left:{x0}pt;margin-top:{y0}pt;width:{max(1,x1-x0)}pt;height:{max(1,y1-y0)}pt;mso-wrap-style:none")
        shape.set(qn("stroked"), "f")
        shape.set(qn("filled"), "f")
        textbox = OxmlElement("v:textbox")
        content = OxmlElement("w:txbxContent")
        wp = OxmlElement("w:p")
        wr = OxmlElement("w:r")
        rp = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(max(12, min(72, int(max(6, y1-y0) * 1.25)))) )
        rp.append(sz)
        wr.append(rp)
        wt = OxmlElement("w:t")
        wt.text = text
        wr.append(wt)
        wp.append(wr)
        content.append(wp)
        textbox.append(content)
        shape.append(textbox)
        pict.append(shape)
        p._p.append(pict)

    def _image(self, doc, blob, x, y, w, h):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run()
        inline = run.add_picture(io.BytesIO(blob), width=Pt(max(1,w)), height=Pt(max(1,h)))
        old = inline._inline
        anchor = OxmlElement("wp:anchor")
        for k,v in {"distT":"0","distB":"0","distL":"0","distR":"0","simplePos":"0","relativeHeight":"0","behindDoc":"0","locked":"0","layoutInCell":"1","allowOverlap":"1"}.items(): anchor.set(k,v)
        simple = OxmlElement("wp:simplePos"); simple.set("x","0"); simple.set("y","0"); anchor.append(simple)
        for tag, value in (("positionH",x),("positionV",y)):
            pos = OxmlElement(f"wp:{tag}"); pos.set("relativeFrom","page")
            off = OxmlElement("wp:posOffset"); off.text = str(int(value*EMU_PER_PT)); pos.append(off); anchor.append(pos)
        extent = OxmlElement("wp:extent"); extent.set("cx",str(int(w*EMU_PER_PT))); extent.set("cy",str(int(h*EMU_PER_PT))); anchor.append(extent)
        eff = OxmlElement("wp:effectExtent")
        for k in ("l","t","r","b"): eff.set(k,"0")
        anchor.append(eff); anchor.append(OxmlElement("wp:wrapNone"))
        for child in list(old): anchor.append(child)
        old.getparent().replace(old, anchor)
