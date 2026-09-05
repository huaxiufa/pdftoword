from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from .docx_renderer import CoordinateDocxRenderer


def pdf_to_docx(pdf_path: Path, out_path: Path, progress=None) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ocr-yolo-") as tmp_name:
        tmp = Path(tmp_name)
        result_dir = tmp / "results"
        progress_file = tmp / "progress.json"
        log_file = tmp / "worker.log"
        env = os.environ.copy()
        timeout = int(env.get("CONVERSION_TIMEOUT", "1800"))
        cmd = [sys.executable, "-m", "app.worker", str(pdf_path), str(result_dir), str(progress_file)]
        if progress:
            progress({"stage":"starting","percent":1,"message":"正在启动 PP-OCRv5 + DocLayout-YOLO…"})
        with log_file.open("w", encoding="utf-8") as log:
            proc = subprocess.Popen(cmd, env=env, stdout=log, stderr=subprocess.STDOUT)
            started = time.monotonic()
            last = None
            while proc.poll() is None:
                if time.monotonic() - started > timeout:
                    proc.kill(); proc.wait()
                    raise TimeoutError(f"转换超过 {timeout} 秒")
                if progress_file.exists():
                    try:
                        current = json.loads(progress_file.read_text(encoding="utf-8"))
                        if current != last and progress:
                            progress(current)
                        last = current
                    except (OSError, json.JSONDecodeError):
                        pass
                time.sleep(0.4)
        if proc.returncode != 0:
            log = log_file.read_text(encoding="utf-8", errors="replace")
            raise RuntimeError("OCR/layout worker failed\n" + log[-12000:])
        json_files = sorted(result_dir.glob("page-*.json"))
        if not json_files:
            raise RuntimeError("没有生成 OCR/layout JSON")
        pages = [json.loads(p.read_text(encoding="utf-8")) for p in json_files]
        if progress:
            progress({"stage":"rendering","percent":96,"message":"正在生成可编辑 Word…"})
        CoordinateDocxRenderer(pdf_path).render(pages, out_path)
        if progress:
            progress({"stage":"done","percent":100,"message":"转换完成"})
