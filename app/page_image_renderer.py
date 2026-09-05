from copy import deepcopy

from docx.oxml import OxmlElement
from docx.shared import Inches, Pt

from app import word_renderer

EMU_PER_PT = 12700


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _add_page_anchored_image(doc, image_info, element, page_rect):
    """Render an image at Gemini's normalized page coordinates.

    The live pdf-to-word endpoint uses app.word_renderer, not app.renderer.
    This hook therefore changes the code path that is actually executed.
    Image geometry comes from the layout element (x/y/w/h), while image bytes
    continue to come from the PDF extractor.
    """
    path = image_info["path"]
    page_w, page_h = page_rect.width, page_rect.height
    x_pt = _safe_float(element.get("x")) * page_w
    y_pt = _safe_float(element.get("y")) * page_h
    width_pt = max(0.01, _safe_float(element.get("w")) * page_w)
    height_pt = max(0.01, _safe_float(element.get("h")) * page_h)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    run = p.add_run()
    inline = run.add_picture(str(path), width=Inches(width_pt / 72), height=Inches(height_pt / 72))
    drawing = inline._inline

    anchor = OxmlElement("wp:anchor")
    for name, value in {
        "distT": "0", "distB": "0", "distL": "0", "distR": "0",
        "simplePos": "0", "relativeHeight": "251658240", "behindDoc": "0",
        "locked": "0", "layoutInCell": "0", "allowOverlap": "1",
    }.items():
        anchor.set(name, value)

    simple = OxmlElement("wp:simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    anchor.append(simple)

    pos_h = OxmlElement("wp:positionH")
    pos_h.set("relativeFrom", "page")
    off_h = OxmlElement("wp:posOffset")
    off_h.text = str(round(x_pt * EMU_PER_PT))
    pos_h.append(off_h)
    anchor.append(pos_h)

    pos_v = OxmlElement("wp:positionV")
    pos_v.set("relativeFrom", "page")
    off_v = OxmlElement("wp:posOffset")
    off_v.text = str(round(y_pt * EMU_PER_PT))
    pos_v.append(off_v)
    anchor.append(pos_v)

    for child in list(drawing):
        anchor.append(deepcopy(child))
    anchor.append(OxmlElement("wp:wrapNone"))
    drawing.getparent().replace(drawing, anchor)


def _patched_add_floating_image(doc, image_info, element, page_rect):
    return _add_page_anchored_image(doc, image_info, element, page_rect)


# render_editable_pdf resolves this module-level hook when it processes each
# image. Patching it here changes the production endpoint without duplicating
# the existing table/text renderer.
word_renderer._add_floating_image = _patched_add_floating_image
render_editable_pdf = word_renderer.render_editable_pdf
