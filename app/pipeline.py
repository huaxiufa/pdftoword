from pathlib import Path

from app.compare import compare_pdf_and_docx, save_report
from app.layout import analyze_layout
from app.parser import parse_pdf
from app.renderer import render_docx


def convert_pdf_to_docx(pdf_path, docx_path, work_dir=None):
    pdf_path = Path(pdf_path)
    docx_path = Path(docx_path)
    model = parse_pdf(pdf_path)
    analyze_layout(model)
    render_docx(model, docx_path)

    report = None
    if work_dir is not None:
        report = compare_pdf_and_docx(pdf_path, docx_path, Path(work_dir))
        save_report(report, Path(work_dir) / "comparison.json")

    return {
        "pages": len(model.pages),
        "blocks": sum(len(p.blocks) for p in model.pages),
        "comparison": report,
    }
