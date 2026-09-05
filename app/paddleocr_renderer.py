from pathlib import Path
import os
import subprocess
import sys


def render_editable_pdf(source_pdf, output):
    """Convert a PDF to editable DOCX in an isolated PaddleOCR worker.

    PaddleOCR/PaddleX keeps global process state (PDX). Running the VL pipeline
    in a short-lived child process prevents that global state from colliding
    with the FastAPI application's imports and request lifecycle.
    """
    source_pdf = Path(source_pdf).resolve()
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    env = os.environ.copy()
    env.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
    proc = subprocess.run(
        [sys.executable, "-m", "app.paddleocr_worker", str(source_pdf), str(output)],
        cwd=str(Path(__file__).resolve().parent.parent),
        env=env,
        capture_output=True,
        text=True,
        timeout=int(os.getenv("PADDLEOCR_TIMEOUT", "900")),
    )
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "PaddleOCR worker failed").strip()
        # Keep API errors bounded while retaining the actual worker exception.
        raise RuntimeError(f"PaddleOCR worker failed (exit {proc.returncode}): {detail[-4000:]}")
    if not output.exists() or output.stat().st_size == 0:
        detail = (proc.stderr or proc.stdout or "no output document").strip()
        raise RuntimeError(f"PaddleOCR worker produced no DOCX: {detail[-2000:]}")
