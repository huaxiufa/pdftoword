from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.pipeline import convert_pdf_to_docx

BASE = Path("/app/data")
UPLOADS = BASE / "uploads"
OUTPUT = BASE / "output"
UPLOADS.mkdir(parents=True, exist_ok=True)
OUTPUT.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="PDF to Word Engine", version="0.1.0")


@app.get("/health")
def health():
    return {"status": "ok", "service": "pdftoword"}


@app.post("/api/v1/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are supported")
    task_id = uuid4().hex
    pdf_path = UPLOADS / f"{task_id}.pdf"
    docx_path = OUTPUT / f"{task_id}.docx"
    pdf_path.write_bytes(await file.read())
    try:
        stats = convert_pdf_to_docx(pdf_path, docx_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}") from exc
    finally:
        pdf_path.unlink(missing_ok=True)
    return {
        "task_id": task_id,
        "status": "completed",
        **stats,
        "download_url": f"/api/v1/files/{task_id}",
    }


@app.get("/api/v1/files/{task_id}")
def download(task_id: str):
    path = OUTPUT / f"{task_id}.docx"
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{task_id}.docx",
    )
