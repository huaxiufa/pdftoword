import json
import shutil
import subprocess
from pathlib import Path

import fitz
import numpy as np


def _render_pdf(path: Path, dpi: int = 100):
    doc = fitz.open(path)
    scale = dpi / 72.0
    matrix = fitz.Matrix(scale, scale)
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        images.append(np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n))
    doc.close()
    return images


def _page_score(a, b):
    h = min(a.shape[0], b.shape[0])
    w = min(a.shape[1], b.shape[1])
    a = a[:h, :w, :3].astype(np.float32)
    b = b[:h, :w, :3].astype(np.float32)
    mae = np.mean(np.abs(a - b)) / 255.0
    score = max(0.0, 1.0 - mae)
    return round(float(score), 4)


def _convert_docx_to_pdf(docx: Path, out_dir: Path) -> Path:
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice is required for visual comparison")
    out_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx)],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    pdf = out_dir / f"{docx.stem}.pdf"
    if result.returncode != 0 or not pdf.exists():
        raise RuntimeError(f"DOCX rendering failed: {result.stderr.strip() or result.stdout.strip()}")
    return pdf


def compare_pdf_and_docx(pdf_path: Path, docx_path: Path, work_dir: Path):
    rendered_pdf = _convert_docx_to_pdf(docx_path, work_dir / "rendered")
    original_pages = _render_pdf(pdf_path)
    output_pages = _render_pdf(rendered_pdf)
    count = min(len(original_pages), len(output_pages))
    scores = [_page_score(original_pages[i], output_pages[i]) for i in range(count)]
    page_count_penalty = 0.0 if len(original_pages) == len(output_pages) else 0.15
    overall = max(0.0, round((sum(scores) / len(scores) if scores else 0.0) - page_count_penalty, 4))

    issues = []
    if len(original_pages) != len(output_pages):
        issues.append({"type": "page_count", "original": len(original_pages), "output": len(output_pages)})
    for index, score in enumerate(scores, start=1):
        if score < 0.90:
            issues.append({"type": "layout_mismatch", "page": index, "score": score})

    report = {
        "overall_score": overall,
        "page_scores": scores,
        "original_pages": len(original_pages),
        "output_pages": len(output_pages),
        "issues": issues,
        "rendered_pdf": str(rendered_pdf),
        "metric": "1 - normalized mean absolute pixel error at 100 DPI; page-count mismatch penalty 0.15",
    }
    return report


def save_report(report, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
