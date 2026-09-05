import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import fitz

from .docx_renderer import CoordinateDocxRenderer


def pdf_to_docx(pdf_path: Path, out_path: Path, progress=None) -> None:
    """PDF -> PP-StructureV3 layout JSON -> editable coordinate DOCX."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="paddle-layout-") as tmp:
        tmp = Path(tmp)
        result_dir = tmp / "results"
        result_dir.mkdir()
        progress_file = tmp / "progress.json"
        log_file = tmp / "paddle.log"
        env = os.environ.copy()
        env.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
        cmd = [
            sys.executable,
            "-m",
            "app.paddle_worker",
            str(pdf_path),
            str(result_dir),
            str(progress_file),
        ]
        timeout = int(os.getenv("PADDLEOCR_TIMEOUT", "1800"))
        if progress:
            progress({"stage": "starting", "percent": 1, "message": "正在启动 PP-StructureV3…"})

        with log_file.open("w", encoding="utf-8") as log:
            proc = subprocess.Popen(
                cmd,
                env=env,
                stdout=log,
                stderr=subprocess.STDOUT,
                text=True,
            )
            started = time.monotonic()
            last = None
            try:
                while proc.poll() is None:
                    if time.monotonic() - started > timeout:
                        proc.kill()
                        proc.wait()
                        raise TimeoutError(f"PaddleOCR timed out after {timeout} seconds")
                    if progress_file.exists():
                        try:
                            current = json.loads(progress_file.read_text(encoding="utf-8"))
                            if current != last and progress:
                                progress(current)
                            last = current
                        except (OSError, json.JSONDecodeError):
                            pass
                    time.sleep(0.5)
                if progress_file.exists():
                    try:
                        current = json.loads(progress_file.read_text(encoding="utf-8"))
                        if progress:
                            progress(current)
                    except (OSError, json.JSONDecodeError):
                        pass
            finally:
                proc.wait()

        if proc.returncode != 0:
            output = log_file.read_text(encoding="utf-8", errors="replace")
            raise RuntimeError("PaddleOCR failed\n" + output[-12000:])

        json_files = sorted(result_dir.glob("*.json"), key=lambda p: _page_number(p.name))
        if not json_files:
            raise RuntimeError("PaddleOCR produced no layout JSON results")
        if progress:
            progress({"stage": "rendering", "percent": 96, "message": "正在生成可编辑 Word…"})
        renderer = CoordinateDocxRenderer(pdf_path)
        renderer.render(json_files, out_path)
        if progress:
            progress({"stage": "done", "percent": 100, "message": "转换完成"})


def _page_number(name: str) -> int:
    stem = Path(name).stem
    digits = "".join(ch for ch in stem[::-1] if ch.isdigit())
    return int(digits[::-1]) if digits else 0
