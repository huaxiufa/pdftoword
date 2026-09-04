import json
import shutil
import subprocess
from io import BytesIO
from pathlib import Path

import fitz
import numpy as np
from PIL import Image

DEFAULT_DPI = 100


def _render_pdf(path: Path, dpi: int = DEFAULT_DPI):
    doc = fitz.open(path)
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        images.append(np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n))
    doc.close()
    return images


def _page_score(a, b):
    h, w = min(a.shape[0], b.shape[0]), min(a.shape[1], b.shape[1])
    left = a[:h, :w, :3].astype(np.float32)
    right = b[:h, :w, :3].astype(np.float32)
    return max(0.0, round(float(1.0 - np.mean(np.abs(left - right)) / 255.0), 4))


def _diagnostic_images(a, b, directory: Path, page_number: int):
    directory.mkdir(parents=True, exist_ok=True)
    h, w = min(a.shape[0], b.shape[0]), min(a.shape[1], b.shape[1])
    left = a[:h, :w, :3].astype(np.float32)
    right = b[:h, :w, :3].astype(np.float32)
    overlay = np.clip((left + right) / 2.0, 0, 255).astype(np.uint8)
    diff = np.clip(np.abs(left - right).mean(axis=2) * 4.0, 0, 255).astype(np.uint8)
    overlay_path = directory / f"overlay-page-{page_number}.png"
    diff_path = directory / f"diff-page-{page_number}.png"
    Image.fromarray(overlay).save(overlay_path)
    Image.fromarray(diff).save(diff_path)
    return {"overlay": str(overlay_path), "diff": str(diff_path)}


def _make_debug_pdf(original: Path, rendered: Path, output: Path, dpi: int = DEFAULT_DPI):
    original_doc = fitz.open(original)
    rendered_doc = fitz.open(rendered)
    debug = fitz.open()
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    for index in range(min(len(original_doc), len(rendered_doc))):
        a = original_doc[index].get_pixmap(matrix=matrix, alpha=False)
        b = rendered_doc[index].get_pixmap(matrix=matrix, alpha=False)
        aa = np.frombuffer(a.samples, dtype=np.uint8).reshape(a.height, a.width, a.n)[: min(a.height, b.height), : min(a.width, b.width), :3]
        bb = np.frombuffer(b.samples, dtype=np.uint8).reshape(b.height, b.width, b.n)[: aa.shape[0], : aa.shape[1], :3]
        overlay = np.clip((aa.astype(np.float32) + bb.astype(np.float32)) / 2.0, 0, 255).astype(np.uint8)
        image = Image.fromarray(overlay)
        stream = BytesIO()
        image.save(stream, format="PNG")
        page = debug.new_page(width=overlay.shape[1] / (dpi / 72.0), height=overlay.shape[0] / (dpi / 72.0))
        page.insert_image(page.rect, stream=stream.getvalue())
    output.parent.mkdir(parents=True, exist_ok=True)
    debug.save(output)
    debug.close()
    original_doc.close()
    rendered_doc.close()


def _convert_docx_to_pdf(docx: Path, out_dir: Path) -> Path:
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice is required for visual comparison")
    out_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx)],
        capture_output=True, text=True, timeout=120, check=False,
    )
    pdf = out_dir / f"{docx.stem}.pdf"
    if result.returncode != 0 or not pdf.exists():
        raise RuntimeError(f"DOCX rendering failed: {result.stderr.strip() or result.stdout.strip()}")
    return pdf


def compare_pdf_and_docx(pdf_path: Path, docx_path: Path, work_dir: Path, make_diagnostics: bool = True):
    work_dir = Path(work_dir)
    rendered_pdf = _convert_docx_to_pdf(Path(docx_path), work_dir / "rendered")
    original_pages, output_pages = _render_pdf(Path(pdf_path)), _render_pdf(rendered_pdf)
    count = min(len(original_pages), len(output_pages))
    scores = [_page_score(original_pages[i], output_pages[i]) for i in range(count)]
    penalty = 0.0 if len(original_pages) == len(output_pages) else 0.15
    overall = max(0.0, round((sum(scores) / len(scores) if scores else 0.0) - penalty, 4))
    issues = []
    diagnostics = []
    if len(original_pages) != len(output_pages):
        issues.append({"type": "page_count", "original": len(original_pages), "output": len(output_pages)})
    for index, score in enumerate(scores, 1):
        if score < 0.90:
            issues.append({"type": "layout_mismatch", "page": index, "score": score})
        if make_diagnostics:
            diagnostics.append(_diagnostic_images(original_pages[index - 1], output_pages[index - 1], work_dir / "diagnostics", index))
    debug_pdf = work_dir / "debug.pdf"
    if make_diagnostics:
        _make_debug_pdf(Path(pdf_path), rendered_pdf, debug_pdf)
    return {
        "overall_score": overall,
        "page_scores": scores,
        "original_pages": len(original_pages),
        "output_pages": len(output_pages),
        "issues": issues,
        "rendered_pdf": str(rendered_pdf),
        "debug_pdf": str(debug_pdf) if debug_pdf.exists() else None,
        "diagnostics": diagnostics,
        "metric": "1 - normalized mean absolute pixel error at 100 DPI; page-count mismatch penalty 0.15",
    }


def save_report(report, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
