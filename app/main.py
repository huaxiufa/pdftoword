from __future__ import annotations

import asyncio
import os
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

from .pipeline import pdf_to_docx

app = FastAPI(title="PDF to Word")
WEB = Path(__file__).resolve().parent.parent / "web" / "index.html"
JOBS: dict[str, dict] = {}


def run_job(job_id: str, pdf: Path, docx: Path):
    def progress(data):
        JOBS[job_id].update(data)
    try:
        pdf_to_docx(pdf, docx, progress)
        JOBS[job_id]["status"] = "done"
    except Exception as exc:
        JOBS[job_id].update(status="error", stage="error", message=str(exc), percent=100)
    finally:
        pdf.unlink(missing_ok=True)


@app.get("/", response_class=HTMLResponse)
def index():
    return WEB.read_text(encoding="utf-8")


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "请上传 PDF 文件")
    max_mb = int(os.getenv("MAX_UPLOAD_MB", "50"))
    job_id = uuid.uuid4().hex
    root = Path(tempfile.gettempdir()) / "pdftoword" / job_id
    root.mkdir(parents=True, exist_ok=True)
    pdf = root / "input.pdf"
    docx = root / "output.docx"
    with pdf.open("wb") as f:
        size = 0
        while chunk := await file.read(1024 * 1024):
            size += len(chunk)
            if size > max_mb * 1024 * 1024:
                shutil.rmtree(root, ignore_errors=True)
                raise HTTPException(413, f"文件不能超过 {max_mb} MB")
            f.write(chunk)
    JOBS[job_id] = {"status":"running","stage":"queued","percent":0,"message":"任务已创建","current_page":0,"total_pages":0}
    asyncio.create_task(asyncio.to_thread(run_job, job_id, pdf, docx))
    return {"task_id": job_id}


@app.get("/progress/{job_id}")
def progress(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(404, "任务不存在")
    return JOBS[job_id]


@app.get("/result/{job_id}")
def result(job_id: str):
    item = JOBS.get(job_id)
    if not item or item.get("status") != "done":
        raise HTTPException(404, "结果尚未生成")
    root = Path(tempfile.gettempdir()) / "pdftoword" / job_id
    path = root / "output.docx"
    if not path.exists():
        raise HTTPException(404, "结果文件不存在")
    return FileResponse(path, filename="converted.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
