from pathlib import Path
import os
import subprocess
import sys


def _is_text_pdf(source_pdf: Path) -> bool:
    """Return True when most PDF pages contain extractable text.

    This deliberately avoids running a document VLM on normal digital PDFs.
    For those files the PDF parser can preserve text, images, tables and page
    geometry much more directly. Scanned/image-only PDFs are sent to the OCR
    recovery worker instead.
    """
    import fitz

    doc = fitz.open(str(source_pdf))
    try:
        if not doc:
            return False
        text_pages = 0
        total_chars = 0
        for page in doc:
            chars = len(page.get_text("text").strip())
            total_chars += chars
            if chars >= 20:
                text_pages += 1
        return text_pages >= max(1, (len(doc) + 1) // 2) and total_chars >= 50
    finally:
        doc.close()


def _pdf2docx(source_pdf: Path, output: Path) -> None:
    """Recover a normal, text-based PDF directly into an editable DOCX."""
    from pdf2docx import Converter

    converter = Converter(str(source_pdf))
    try:
        converter.convert(str(output), start=0, end=None)
    finally:
        converter.close()


def _ocr_recovery(source_pdf: Path, output: Path) -> None:
    """Run PP-StructureV3 recovery in an isolated process for scanned PDFs."""
    env = os.environ.copy()
    env.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
    timeout = int(os.getenv("PADDLEOCR_TIMEOUT", "1800"))
    cmd = [
        sys.executable,
        "-m",
        "app.paddleocr_worker",
        str(source_pdf),
        str(output),
    ]

    proc = subprocess.Popen(
        cmd,
        cwd=str(Path(__file__).resolve().parent.parent),
        env=env,
        stdout=None,
        stderr=None,
        text=True,
    )
    try:
        return_code = proc.wait(timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        proc.kill()
        proc.wait()
        raise RuntimeError(
            f"PaddleOCR recovery timed out after {timeout} seconds"
        ) from exc
    if return_code != 0:
        raise RuntimeError(f"PaddleOCR recovery failed (exit {return_code})")


def render_editable_pdf(source_pdf, output):
    """Convert a PDF to an editable DOCX using the lightest suitable path.

    AUTO mode uses direct PDF parsing for text PDFs and PP-StructureV3 OCR
    recovery for scanned PDFs. Set PDF_TO_WORD_MODE to ``pdf2docx`` or ``ocr``
    to force a specific path.
    """
    source_pdf = Path(source_pdf).resolve()
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    mode = os.getenv("PDF_TO_WORD_MODE", "auto").strip().lower()
    if mode not in {"auto", "pdf2docx", "ocr"}:
        raise ValueError("PDF_TO_WORD_MODE must be auto, pdf2docx or ocr")

    use_pdf2docx = mode == "pdf2docx" or (mode == "auto" and _is_text_pdf(source_pdf))

    if use_pdf2docx:
        try:
            _pdf2docx(source_pdf, output)
            if output.exists() and output.stat().st_size > 0:
                return
        except Exception as exc:
            if mode == "pdf2docx":
                raise
            print(
                f"PDF parser recovery failed; falling back to OCR recovery: "
                f"{type(exc).__name__}: {exc}",
                flush=True,
            )

    _ocr_recovery(source_pdf, output)
    if not output.exists() or output.stat().st_size == 0:
        raise RuntimeError("PDF recovery produced no DOCX")
