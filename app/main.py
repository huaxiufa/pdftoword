from __future__ import annotations

import asyncio
import os
import shutil
import tempfile
import traceback
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

from .v5_renderer import convert_pdf_to_docx

app = FastAPI(title="PDF to Word - V5")
WEB = Path(__file__).resolve().parent.parent / "web" / "index.html"
JOBS: dict[str, dict] = {}
BUILD_VERSION = "2026-09-08-pdftoword-v5"


def run_job(job_id: str, pdf: Path, docx: Path) -> None:
    def progress(data: dict) -> None:
        JOBS[job_id].update(data)

    try:
        convert_pdf_to_docx(pdf, docx, progress)
        JOBS[job_id].update(status="done", stage="done", percent=100, message="V5 高保真转换完成")
    except Exception as exc:
        detail = traceback.format_exc()
        print(detail, flush=True)
        JOBS[job_id].update(
            status="error",
            stage="error",
            message=f"{type(exc).__name__}: {exc}",
            error_detail=detail,
            percent=100,
        )
    finally:
        pdf.unlink(missing_ok=True)


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    response = HTMLResponse(WEB.read_text(encoding="utf-8"))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response


@app.get("/health")
def health() -> dict:
    return {"ok": True, "engine": "PyMuPDF", "renderer": "DOCX V5 visual fidelity", "version": BUILD_VERSION}


@app.get("/version")
def version() -> dict:
    return {"engine": "PyMuPDF", "renderer": "DOCX V5 visual fidelity", "version": BUILD_VERSION}


@app.post("/convert")
async def convert(file: UploadFile = File(...)) -> dict:
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "请上传 PDF 文件")

    max_mb = int(os.getenv("MAX_UPLOAD_MB", "50"))
    job_id = uuid.uuid4().hex
    root = Path(tempfile.gettempdir()) / "pdftoword" / job_id
    root.mkdir(parents=True, exist_ok=True)
    pdf = root / "input.pdf"
    docx = root / "output.docx"

    try:
        with pdf.open("wb") as target:
            size = 0
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > max_mb * 1024 * 1024:
                    raise HTTPException(413, f"文件不能超过 {max_mb} MB")
                target.write(chunk)
    except Exception:
        shutil.rmtree(root, ignore_errors=True)
        raise

    JOBS[job_id] = {
        "status": "running",
        "stage": "queued",
        "percent": 0,
        "message": "V5 高保真任务已创建",
    }
    asyncio.create_task(asyncio.to_thread(run_job, job_id, pdf, docx))
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
    path = Path(tempfile.gettempdir()) / "pdftoword" / job_id / "output.docx"
    if not path.exists():
        raise HTTPException(404, "结果文件不存在")
    return FileResponse(
        path,
        filename="converted-v5.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
