from __future__ import annotations

import io
from pathlib import Path

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import nsmap, qn
from docx.shared import Pt

EMU_PER_PT = 12700
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"

# python-docx's OxmlElement accepts namespace prefixes (e.g. ``wps:wsp``),
# but the newer WordprocessingShape namespace is not registered by default.
nsmap.setdefault("wps", WPS_URI)


def el(tag: str):
    return OxmlElement(tag)


def text_box(x: float, y: float, w: float, h: float, text: str, ident: int):
    anchor = el("wp:anchor")
    for k, v in {"distT":"0","distB":"0","distL":"0","distR":"0","simplePos":"0","relativeHeight":"1","behindDoc":"0","locked":"0","layoutInCell":"1","allowOverlap":"1"}.items():
        anchor.set(k, v)
    simple = el("wp:simplePos"); simple.set("x", "0"); simple.set("y", "0"); anchor.append(simple)
    for tag, value in (("wp:positionH", x), ("wp:positionV", y)):
        pos = el(tag); pos.set("relativeFrom", "page")
        off = el("wp:posOffset"); off.text = str(int(value * EMU_PER_PT)); pos.append(off); anchor.append(pos)
    extent = el("wp:extent"); extent.set("cx", str(max(1, int(w * EMU_PER_PT)))); extent.set("cy", str(max(1, int(h * EMU_PER_PT)))); anchor.append(extent)
    effect = el("wp:effectExtent")
    for k in ("l","t","r","b"): effect.set(k, "0")
    anchor.append(effect); anchor.append(el("wp:wrapNone"))
    docpr = el("wp:docPr"); docpr.set("id", str(ident)); docpr.set("name", f"Text {ident}"); anchor.append(docpr)
    anchor.append(el("wp:cNvGraphicFramePr"))

    graphic = el("a:graphic"); data = el("a:graphicData"); data.set("uri", WPS_URI)
    wsp = el("wps:wsp")
    wsp.append(el("wps:cNvSpPr"))
    sppr = el("wps:spPr"); xfrm = el("a:xfrm")
    off = el("a:off"); off.set("x", "0"); off.set("y", "0")
    ext = el("a:ext"); ext.set("cx", str(max(1, int(w * EMU_PER_PT)))); ext.set("cy", str(max(1, int(h * EMU_PER_PT))))
    xfrm.extend([off, ext]); sppr.append(xfrm)
    geom = el("a:prstGeom"); geom.set("prst", "rect"); geom.append(el("a:avLst")); sppr.append(geom)
    sppr.append(el("a:noFill")); wsp.append(sppr)

    txbx = el("wps:txbx"); content = el("w:txbxContent"); p = el("w:p")
    ppr = el("w:pPr"); spacing = el("w:spacing"); spacing.set(qn("w:before"), "0"); spacing.set(qn("w:after"), "0"); spacing.set(qn("w:line"), "240"); ppr.append(spacing); p.append(ppr)
    r = el("w:r"); rpr = el("w:rPr"); size = el("w:sz"); size.set(qn("w:val"), str(max(12, min(72, int(max(6, h) * 1.15))))); rpr.append(size); r.append(rpr)
    wt = el("w:t"); wt.text = text; r.append(wt); p.append(r); content.append(p); txbx.append(content); wsp.append(txbx)
    body = el("wps:bodyPr"); body.set("lIns","0"); body.set("tIns","0"); body.set("rIns","0"); body.set("bIns","0"); wsp.append(body)
    data.append(wsp); graphic.append(data); anchor.append(graphic)
    return anchor


class CoordinateDocxRenderer:
    def __init__(self, pdf_path: Path):
        self.pdf = fitz.open(pdf_path)
        self.ident = 1000

    def render(self, pages: list[dict], out_path: Path) -> None:
        doc = Document()
        for i, item in enumerate(pages):
            page = self.pdf[item["page_index"]]
            section = doc.sections[0] if i == 0 else doc.add_section(WD_SECTION.NEW_PAGE)
            section.page_width = Pt(page.rect.width); section.page_height = Pt(page.rect.height)
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Pt(0)
            section.header_distance = section.footer_distance = Pt(0)
            self._page(doc, page, item)
        self.pdf.close(); doc.save(out_path)

    def _page(self, doc, page, item):
        seen = set()
        for image in page.get_images(full=True):
            xref = image[0]
            for rect in page.get_image_rects(xref):
                if rect.width * rect.height >= page.rect.width * page.rect.height * 0.92: continue
                key = (xref, round(rect.x0,1), round(rect.y0,1), round(rect.x1,1), round(rect.y1,1))
                if key in seen: continue
                seen.add(key)
                try:
                    pix = fitz.Pixmap(self.pdf, xref)
                    if pix.alpha: pix = fitz.Pixmap(fitz.csRGB, pix)
                    self._image(doc, pix.tobytes("png"), rect.x0, rect.y0, rect.width, rect.height)
                except Exception: pass
        for region in item.get("regions", []):
            for line in region.get("lines", []):
                text = str(line.get("text", "")).strip(); bbox = line.get("bbox") or region.get("bbox")
                if text and bbox: self._text(doc, text, bbox)

    def _text(self, doc, text, bbox):
        x0,y0,x1,y1 = map(float, bbox[:4]); p = doc.add_paragraph(); drawing = el("w:drawing")
        drawing.append(text_box(x0, y0, max(2,x1-x0), max(2,y1-y0), text, self.ident)); self.ident += 1; p._p.append(drawing)

    def _image(self, doc, blob, x, y, w, h):
        p = doc.add_paragraph(); run = p.add_run(); inline = run.add_picture(io.BytesIO(blob), width=Pt(max(1,w)), height=Pt(max(1,h))); old = inline._inline
        anchor = el("wp:anchor")
        for k,v in {"distT":"0","distB":"0","distL":"0","distR":"0","simplePos":"0","relativeHeight":"0","behindDoc":"0","locked":"0","layoutInCell":"1","allowOverlap":"1"}.items(): anchor.set(k,v)
        simple=el("wp:simplePos"); simple.set("x","0"); simple.set("y","0"); anchor.append(simple)
        for tag,value in (("wp:positionH",x),("wp:positionV",y)):
            pos=el(tag); pos.set("relativeFrom","page"); off=el("wp:posOffset"); off.text=str(int(value*EMU_PER_PT)); pos.append(off); anchor.append(pos)
        extent=el("wp:extent"); extent.set("cx",str(int(max(1,w)*EMU_PER_PT))); extent.set("cy",str(int(max(1,h)*EMU_PER_PT))); anchor.append(extent)
        effect=el("wp:effectExtent")
        for k in ("l","t","r","b"): effect.set(k,"0")
        anchor.append(effect); anchor.append(el("wp:wrapNone")); dp=el("wp:docPr"); dp.set("id",str(self.ident)); dp.set("name",f"Image {self.ident}"); self.ident+=1; anchor.append(dp); anchor.append(el("wp:cNvGraphicFramePr"))
        graphic=el("a:graphic"); data=el("a:graphicData"); data.set("uri",PICTURE_URI)
        for child in list(old): data.append(child)
        graphic.append(data); anchor.append(graphic); old.getparent().replace(old,anchor)
