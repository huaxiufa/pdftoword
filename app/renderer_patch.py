from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_east_asia_font(run, font_name: str):
    """Set Word's East Asian font safely through OOXML."""
    run.font.name = font_name
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:eastAsia"), font_name)
