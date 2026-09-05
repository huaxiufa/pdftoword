import asyncio
import os
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from .pipeline import pdf_to_docx

app = FastAPI(title="PDF to Word", version="2.1.0")
MAX_MB = int(os.getenv("MAX_UPLOAD_MB", "50"))
JOBS: dict[str, dict] = {}


@app.get("/", response_class=FileResponse)
def index():
    return FileResponse(Path(__file__).parent.parent / "web" / "index.html")


@app.get("/health")
def health():
    return {"ok": True, "engine": "PaddleOCR PP-StructureV3 + coordinate DOCX"}


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Only PDF files are supported")

    job_id = uuid.uuid4().hex
    work = Path(tempfile.mkdtemp(prefix=f"pdftoword-{job_id}-"))
    pdf_path = work / "input.pdf"
    out_path = work / f"{Path(file.filename).stem}.docx"
    try:
        with pdf_path.open("wb") as f:
            size = 0
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_MB * 1024 * 1024:
                    shutil.rmtree(work, ignore_errors=True)
                    raise HTTPException(413, f"File is larger than {MAX_MB} MB")
                f.write(chunk)
    except Exception:
        shutil.rmtree(work, ignore_errors=True)
        raise

    JOBS[job_id] = {
        "status": "queued",
        "percent": 0,
        "message": "任务已创建，等待 PP-StructureV3…",
        "filename": f"{Path(file.filename).stem}.docx",
        "work": str(work),
        "output": str(out_path),
    }
    asyncio.create_task(_run_job(job_id, pdf_path, out_path))
    return {"task_id": job_id}


@app.get("/progress/{job_id}")
def progress(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Task not found")
    return {
        "task_id": job_id,
        "status": job["status"],
        "percent": job["percent"],
        "message": job["message"],
    }


@app.get("/result/{job_id}")
def result(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Task not found")
    if job["status"] != "done":
        raise HTTPException(409, "Conversion is not complete")
    output = Path(job["output"])
    if not output.exists():
        raise HTTPException(404, "Result file not found")
    return FileResponse(
        output,
        filename=job["filename"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


async def _run_job(job_id: str, pdf_path: Path, out_path: Path):
    job = JOBS[job_id]

    def update(value):
        job.update(value)
        job["status"] = value.get("stage", job["status"])

    try:
        await asyncio.to_thread(pdf_to_docx, pdf_path, out_path, update)
        job.update({"status": "done", "percent": 100, "message": "转换完成"})
    except Exception as exc:
        job.update({
            "status": "error",
            "percent": job.get("percent", 0),
            "message": f"转换失败：{exc}",
        })
