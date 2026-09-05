import os
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from .pipeline import pdf_to_docx

app = FastAPI(title="PDF to Word", version="2.0.0")

MAX_MB = int(os.getenv("MAX_UPLOAD_MB", "50"))

@app.get("/")
def index():
    return {"service": "pdf-to-word", "engine": "PaddleOCR PP-StructureV3 + coordinate DOCX"}

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Only PDF files are supported")

    work = Path(tempfile.mkdtemp(prefix="pdftoword-"))
    pdf_path = work / "input.pdf"
    out_path = work / f"{Path(file.filename).stem}.docx"
    try:
        with pdf_path.open("wb") as f:
            size = 0
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_MB * 1024 * 1024:
                    raise HTTPException(413, f"File is larger than {MAX_MB} MB")
                f.write(chunk)
        pdf_to_docx(pdf_path, out_path)
        download_name = f"{Path(file.filename).stem}.docx"
        return FileResponse(out_path, filename=download_name,
                            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            background=None)
    except HTTPException:
        shutil.rmtree(work, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(work, ignore_errors=True)
        raise HTTPException(500, f"Conversion failed: {exc}") from exc

@app.get("/health")
def health():
    return {"ok": True}
