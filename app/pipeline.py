from app.layout import analyze_layout
from app.parser import parse_pdf
from app.renderer import render_docx


def convert_pdf_to_docx(pdf_path, docx_path):
    model = parse_pdf(pdf_path)
    analyze_layout(model)
    render_docx(model, docx_path)
    return {
        "pages": len(model.pages),
        "blocks": sum(len(p.blocks) for p in model.pages),
    }
