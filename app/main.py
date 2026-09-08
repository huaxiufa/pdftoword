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

from .excel_cleaner import clean_file

app = FastAPI(title="Excel / CSV Cleaner")
WEB = Path(__file__).resolve().parent.parent / "web" / "index.html"
ROOT = Path(tempfile.gettempdir()) / "excel-cleaner"
ROOT.mkdir(parents=True, exist_ok=True)
JOBS: dict[str, dict] = {}
BUILD_VERSION = "2026-09-08-excel-cleaner-v1"


def run_job(job_id: str, source: Path, output: Path, options: dict) -> None:
    def progress(percent: int, message: str) -> None:
        JOBS[job_id].update(percent=percent, message=message)

    try:
        summary = clean_file(source, output, options, progress)
        JOBS[job_id].update(status="done", stage="done", percent=100, message="清洗完成", summary=summary)
    except Exception as exc:
        detail = traceback.format_exc()
        print(detail, flush=True)
        JOBS[job_id].update(status="error", stage="error", percent=100, message=f"{type(exc).__name__}: {exc}", error_detail=detail)
    finally:
        source.unlink(missing_ok=True)


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    response = HTMLResponse(WEB.read_text(encoding="utf-8"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


@app.get("/health")
def health() -> dict:
    return {"ok": True, "engine": "pandas + openpyxl", "version": BUILD_VERSION}


@app.get("/version")
def version() -> dict:
    return {"engine": "pandas + openpyxl", "version": BUILD_VERSION}


@app.post("/clean")
async def clean(file: UploadFile = File(...), options: str = "{}") -> dict:
    if not file.filename or Path(file.filename).suffix.lower() not in {".csv", ".xlsx", ".xlsm"}:
        raise HTTPException(400, "请上传 CSV、XLSX 或 XLSM 文件")
    try:
        parsed_options = json.loads(options or "{}")
    except json.JSONDecodeError:
        raise HTTPException(400, "清洗参数格式错误")

    max_mb = int(os.getenv("MAX_UPLOAD_MB", "100"))
    job_id = uuid.uuid4().hex
    root = ROOT / job_id
    root.mkdir(parents=True, exist_ok=True)
    suffix = Path(file.filename).suffix.lower()
    source = root / f"input{suffix}"
    output = root / "cleaned.xlsx"

    try:
        with source.open("wb") as target:
            size = 0
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > max_mb * 1024 * 1024:
                    raise HTTPException(413, f"文件不能超过 {max_mb} MB")
                target.write(chunk)
    except Exception:
        shutil.rmtree(root, ignore_errors=True)
        raise

    JOBS[job_id] = {"status": "running", "stage": "queued", "percent": 0, "message": "任务已创建"}
    asyncio.create_task(asyncio.to_thread(run_job, job_id, source, output, parsed_options))
    return {"task_id": job_id}


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
    path = ROOT / job_id / "cleaned.xlsx"
    if not path.exists():
        raise HTTPException(404, "结果文件不存在")
    return FileResponse(path, filename="cleaned-data.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
