from pathlib import Path
import shutil

from app.compare import compare_pdf_and_docx, save_report
from app.layout import analyze_layout
from app.parser import parse_pdf
from app.renderer import render_docx


def convert_pdf_to_docx(pdf_path, docx_path, work_dir=None):
    """Convert PDF to DOCX and automatically search a small set of layout-density candidates.

    The original PDF is never modified. Candidate DOCX files are rendered back to PDF and
    scored; the highest-scoring candidate is copied to the requested output path.
    """
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
    candidates_dir = work_dir / "candidates"
    candidates_dir.mkdir(parents=True, exist_ok=True)

    # Start conservative and progressively compact vertical layout/font size.
    candidates = [
        (1.00, 1.00),
        (1.00, 0.94),
        (1.00, 0.88),
        (0.98, 0.94),
        (0.96, 0.90),
        (0.94, 0.86),
    ]

    best = None
    for index, (font_scale, vertical_scale) in enumerate(candidates):
        candidate_docx = candidates_dir / f"candidate-{index}.docx"
        render_docx(
            model,
            candidate_docx,
            font_scale=font_scale,
            vertical_scale=vertical_scale,
        )
        candidate_report = compare_pdf_and_docx(
            pdf_path,
            candidate_docx,
            work_dir / f"candidate-{index}",
        )
        candidate_report["font_scale"] = font_scale
        candidate_report["vertical_scale"] = vertical_scale

        score = candidate_report["overall_score"]
        # Prefer exact page count when scores are close; otherwise use the visual score.
        exact_pages = candidate_report["original_pages"] == candidate_report["output_pages"]
        rank = (1 if exact_pages else 0, score)
        if best is None or rank > best["rank"]:
            best = {
                "rank": rank,
                "candidate": candidate_docx,
                "report": candidate_report,
            }

    shutil.copy2(best["candidate"], docx_path)

    # Re-render the selected final file and expose its authoritative comparison report.
    final_report = compare_pdf_and_docx(pdf_path, docx_path, work_dir / "final")
    final_report["optimization"] = {
        "enabled": True,
        "candidates_tested": len(candidates),
        "selected_font_scale": best["report"]["font_scale"],
        "selected_vertical_scale": best["report"]["vertical_scale"],
        "selection_rule": "prefer exact page count, then highest visual score",
    }
    save_report(final_report, work_dir / "comparison.json")

    return {
        "pages": len(model.pages),
        "blocks": sum(len(p.blocks) for p in model.pages),
        "comparison": final_report,
    }
