import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz

from .docx_renderer import CoordinateDocxRenderer


def _has_text(pdf_path: Path) -> bool:
    doc = fitz.open(pdf_path)
    try:
        return any(page.get_text("text").strip() for page in doc)
    finally:
        doc.close()


def pdf_to_docx(pdf_path: Path, out_path: Path) -> None:
    """PDF -> layout JSON -> editable DOCX.

    The PDF itself is rendered by PyMuPDF only for geometry/image extraction.
    PaddleOCR is responsible for OCR/layout/table understanding, including
    scanned PDFs. No pdf2docx fallback is used.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="paddle-layout-") as tmp:
        result_dir = Path(tmp) / "results"
        result_dir.mkdir()
        env = os.environ.copy()
        env.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
        cmd = [sys.executable, "-m", "app.paddle_worker", str(pdf_path), str(result_dir)]
        timeout = int(os.getenv("PADDLEOCR_TIMEOUT", "1800"))
        completed = subprocess.run(cmd, env=env, capture_output=True, text=True,
                                   timeout=timeout)
        if completed.returncode != 0:
            raise RuntimeError(
                "PaddleOCR failed\nSTDOUT:\n%s\nSTDERR:\n%s"
                % (completed.stdout[-12000:], completed.stderr[-12000:])
            )
        json_files = sorted(result_dir.glob("*.json"), key=lambda p: _page_number(p.name))
        if not json_files:
            raise RuntimeError("PaddleOCR produced no layout JSON results")
        renderer = CoordinateDocxRenderer(pdf_path)
        renderer.render(json_files, out_path)


def _page_number(name: str) -> int:
    stem = Path(name).stem
    digits = "".join(ch for ch in stem[::-1] if ch.isdigit())
    return int(digits[::-1]) if digits else 0
