import json
from pathlib import Path

from app.compare import compare_pdf_and_docx
from app.renderer import render_docx


# Candidate grid intentionally favors density reductions because an extra output
# page is much more damaging than a small amount of whitespace compression.
DEFAULT_CANDIDATES = [
    (1.00, 1.00),
    (1.00, 0.97),
    (1.00, 0.94),
    (0.98, 0.94),
    (0.96, 0.92),
    (0.94, 0.90),
    (0.92, 0.88),
    (0.90, 0.86),
    (0.88, 0.84),
    (0.86, 0.82),
    (0.84, 0.80),
    (0.82, 0.78),
]


class LayoutOptimizer:
    """Search tunable DOCX density parameters against the rendered visual score."""

    def __init__(self, model, source_pdf: Path, work_dir: Path, candidates=None, max_iterations=12):
        self.model = model
        self.source_pdf = Path(source_pdf)
        self.work_dir = Path(work_dir)
        self.candidates = candidates or DEFAULT_CANDIDATES
        self.max_iterations = max_iterations

    @staticmethod
    def _rank(report):
        # Page-count correctness is the hard constraint. Only compare visual
        # quality after exact page count has been established.
        exact_pages = report["original_pages"] == report["output_pages"]
        return (1 if exact_pages else 0, report["overall_score"])

    def run(self):
        candidates_dir = self.work_dir / "candidates"
        candidates_dir.mkdir(parents=True, exist_ok=True)
        history = []
        best = None

        for index, (font_scale, vertical_scale) in enumerate(self.candidates[: self.max_iterations]):
            candidate_dir = self.work_dir / f"iteration-{index + 1:02d}"
            candidate_docx = candidates_dir / f"candidate-{index:02d}.docx"
            render_docx(self.model, candidate_docx, font_scale=font_scale, vertical_scale=vertical_scale)
            report = compare_pdf_and_docx(self.source_pdf, candidate_docx, candidate_dir, make_diagnostics=False)
            item = {
                "iteration": index + 1,
                "font_scale": font_scale,
                "vertical_scale": vertical_scale,
                "overall_score": report["overall_score"],
                "page_scores": report["page_scores"],
                "original_pages": report["original_pages"],
                "output_pages": report["output_pages"],
            }
            history.append(item)
            if best is None or self._rank(report) > self._rank(best["report"]):
                best = {"candidate": candidate_docx, "report": report, "parameters": item}

            # Never stop on score stability while the page count is still wrong.
            # The previous implementation could terminate after three similar
            # 4-page candidates and never reach the denser candidates that fit
            # the source's 3-page layout.
            if report["original_pages"] == report["output_pages"] and report["overall_score"] >= 0.985:
                break

        history_path = self.work_dir / "optimization-history.json"
        history_path.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
        return best, history, history_path
