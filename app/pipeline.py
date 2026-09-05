import os
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz

from .docx_renderer import CoordinateDocxRenderer


def _page_has_native_text(page) -> bool:
    text = page.get_text("text").strip()
    return len(text) >= 2


def _run_paddle(pdf_path: Path, result_dir: Path) -> None:
    env = os.environ.copy()
    env.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
    cmd = [sys.executable, "-m", "app.paddle_worker", str(pdf_path), str(result_dir)]
    timeout = int(os.getenv("PADDLEOCR_TIMEOUT", "1800"))
    completed = subprocess.run(
        cmd, env=env, capture_output=True, text=True, timeout=timeout
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "PaddleOCR failed\nSTDOUT:\n%s\nSTDERR:\n%s"
            % (completed.stdout[-12000:], completed.stderr[-12000:])
        )


def pdf_to_docx(pdf_path: Path, out_path: Path) -> None:
    """Fast hybrid PDF -> editable DOCX pipeline.

    Native-text pages never enter PaddleOCR: PyMuPDF supplies text/image/table
    coordinates directly. Only pages without usable native text are rendered
    through the isolated PaddleOCR PP-StructureV3 worker.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)
    source = fitz.open(pdf_path)
    try:
        native_pages = [i for i, page in enumerate(source) if _page_has_native_text(page)]
        scan_pages = [i for i in range(len(source)) if i not in set(native_pages)]

        if not scan_pages:
            CoordinateDocxRenderer(pdf_path).render_native(out_path)
            return

        with tempfile.TemporaryDirectory(prefix="paddle-layout-") as tmp:
            tmp = Path(tmp)
            scan_pdf = tmp / "scanned-pages.pdf"
            result_dir = tmp / "results"
            result_dir.mkdir()

            scan_doc = fitz.open()
            try:
                for page_index in scan_pages:
                    scan_doc.insert_pdf(source, from_page=page_index, to_page=page_index)
                scan_doc.save(scan_pdf)
            finally:
                scan_doc.close()

            _run_paddle(scan_pdf, result_dir)
            json_files = sorted(
                result_dir.glob("*.json"), key=lambda p: _page_number(p.name)
            )
            if len(json_files) != len(scan_pages):
                raise RuntimeError(
                    f"PaddleOCR returned {len(json_files)} pages for {len(scan_pages)} scanned pages"
                )
            CoordinateDocxRenderer(pdf_path).render_mixed(
                out_path, {page_index: json_files[n] for n, page_index in enumerate(scan_pages)}
            )
    finally:
        source.close()


def _page_number(name: str) -> int:
    stem = Path(name).stem
    digits = "".join(ch for ch in stem[::-1] if ch.isdigit())
    return int(digits[::-1]) if digits else 0
