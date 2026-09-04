from pathlib import Path
import shutil

from app.compare import compare_pdf_and_docx, save_report
from app.layout import analyze_layout
from app.optimizer import LayoutOptimizer
from app.parser import parse_pdf
from app.renderer import render_docx


def convert_pdf_to_docx(pdf_path, docx_path, work_dir=None):
    """Convert PDF to editable DOCX with a measurable, bounded layout optimization loop."""
    pdf_path = Path(pdf_path)
    docx_path = Path(docx_path)
    model = parse_pdf(pdf_path)
    analyze_layout(model)

    if work_dir is None:
        render_docx(model, docx_path)
        return {
            "pages": len(model.pages),
            "blocks": sum(len(p.blocks) for p in model.pages),
            "comparison": None,
        }

    work_dir = Path(work_dir)
    optimizer = LayoutOptimizer(model, pdf_path, work_dir)
    best, history, history_path = optimizer.run()
    if best is None:
        raise RuntimeError("Optimization produced no candidate")

    shutil.copy2(best["candidate"], docx_path)
    final_report = compare_pdf_and_docx(pdf_path, docx_path, work_dir / "final", make_diagnostics=True)
    final_report["optimization"] = {
        "enabled": True,
        "iterations": len(history),
        "selected_font_scale": best["parameters"]["font_scale"],
        "selected_vertical_scale": best["parameters"]["vertical_scale"],
        "selection_rule": "prefer exact page count, then highest visual score",
        "history": str(history_path),
    }
    save_report(final_report, work_dir / "comparison.json")

    return {
        "pages": len(model.pages),
        "blocks": sum(len(p.blocks) for p in model.pages),
        "comparison": final_report,
    }
