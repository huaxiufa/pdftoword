from __future__ import annotations

import asyncio
import json
import os
import shutil
import tempfile
import traceback
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

from .image_compressor import SUPPORTED, compress_files

app = FastAPI(title="Image Compressor")
WEB = Path(__file__).resolve().parent.parent / "web" / "index.html"
ROOT = Path(tempfile.gettempdir()) / "image-compressor"
ROOT.mkdir(parents=True, exist_ok=True)
JOBS: dict[str, dict] = {}
BUILD_VERSION = "2026-09-08-image-compressor-v1"


def run_job(job_id: str, sources: list[Path], output: Path, options: dict) -> None:
    def progress(percent: int, message: str) -> None:
        JOBS[job_id].update(percent=percent, message=message)

    try:
        summary = compress_files(sources, output, options, progress)
        JOBS[job_id].update(status="done", stage="done", percent=100, message="压缩完成", summary=summary)
    except Exception as exc:
        detail = traceback.format_exc()
        print(detail, flush=True)
        JOBS[job_id].update(status="error", stage="error", percent=100, message=f"{type(exc).__name__}: {exc}", error_detail=detail)
    finally:
        for source in sources:
            source.unlink(missing_ok=True)


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    response = HTMLResponse(WEB.read_text(encoding="utf-8"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


@app.get("/health")
def health() -> dict:
    return {"ok": True, "engine": "Pillow", "version": BUILD_VERSION}


@app.get("/version")
def version() -> dict:
    return {"engine": "Pillow", "version": BUILD_VERSION}


@app.post("/compress")
async def compress(files: list[UploadFile] = File(...), options: str = "{}") -> dict:
    valid = [f for f in files if f.filename and Path(f.filename).suffix.lower() in SUPPORTED]
    if not valid:
        raise HTTPException(400, "请上传 JPG、PNG 或 WebP 图片")
    try:
        parsed = json.loads(options or "{}")
    except json.JSONDecodeError:
        raise HTTPException(400, "压缩参数格式错误")

    max_mb = int(os.getenv("MAX_UPLOAD_MB", "100"))
    job_id = uuid.uuid4().hex
    root = ROOT / job_id
    root.mkdir(parents=True, exist_ok=True)
    sources: list[Path] = []
    total_size = 0
    try:
        for index, file in enumerate(valid):
            suffix = Path(file.filename).suffix.lower()
            source = root / f"input-{index}{suffix}"
            with source.open("wb") as target:
                while chunk := await file.read(1024 * 1024):
                    total_size += len(chunk)
                    if total_size > max_mb * 1024 * 1024:
                        raise HTTPException(413, f"总文件大小不能超过 {max_mb} MB")
                    target.write(chunk)
            sources.append(source)
    except Exception:
        shutil.rmtree(root, ignore_errors=True)
        raise

    output = root / "compressed-images.zip"
    JOBS[job_id] = {"status": "running", "stage": "queued", "percent": 0, "message": "任务已创建"}
    asyncio.create_task(asyncio.to_thread(run_job, job_id, sources, output, parsed))
    return {"task_id": job_id, "count": len(sources)}


@app.get("/progress/{job_id}")
def progress(job_id: str) -> dict:
    if job_id not in JOBS:
        raise HTTPException(404, "任务不存在")
    return JOBS[job_id]


@app.get("/result/{job_id}")
def result(job_id: str) -> FileResponse:
    item = JOBS.get(job_id)
    if not item or item.get("status") != "done":
        raise HTTPException(404, "结果尚未生成")
    path = ROOT / job_id / "compressed-images.zip"
    if not path.exists():
        raise HTTPException(404, "结果文件不存在")
    return FileResponse(path, filename="compressed-images.zip", media_type="application/zip")
