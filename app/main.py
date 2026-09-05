import os
import shutil
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from .pipeline import pdf_to_docx

app = FastAPI(title="PDF to Word", version="2.0.0")
MAX_MB = int(os.getenv("MAX_UPLOAD_MB", "50"))

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
        return FileResponse(out_path, filename=f"{Path(file.filename).stem}.docx",
                            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except HTTPException:
        shutil.rmtree(work, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(work, ignore_errors=True)
        raise HTTPException(500, f"Conversion failed: {exc}") from exc
