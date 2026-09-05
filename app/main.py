from pathlib import Path
from uuid import uuid4
import os
import shutil
import json

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.gemini import pdf_layout_analysis, pdf_to_structured
from app.pdf_tools import (
    extract_pages, images_to_pdf, merge_pdfs, pdf_to_images, rotate_pdf,
    split_pdf, compress_pdf, structured_to_xlsx,
)
from app.page_image_renderer import render_editable_pdf

BASE = Path(os.getenv("DATA_DIR", "/app/data")); OUTPUT = BASE / "output"
UPLOADS = BASE / "uploads"; OUTPUT.mkdir(parents=True, exist_ok=True); UPLOADS.mkdir(parents=True, exist_ok=True)
app = FastAPI(title="PDF Toolbox", version="1.0.0")

TOOLS = {"merge","split","extract","rotate","compress","pdf-to-word","pdf-to-excel","pdf-to-images","images-to-pdf"}

@app.get("/health")
def health():
    return {
        "status":"ok",
        "gemini": bool(os.getenv("GEMINI_API_KEY")),
        "model": os.getenv("GEMINI_MODEL", "gemini-3.7-flash"),
        "fallback_model": os.getenv("GEMINI_FALLBACK_MODEL", "gemini-3.6-flash"),
    }

@app.post("/api/v1/tools/{tool}")
async def tool(tool: str, files: list[UploadFile] = File(...), pages: str = "", angle: int = 90):
    if tool not in TOOLS: raise HTTPException(404, "Unknown tool")
    if not files: raise HTTPException(400, "No files")
    task = uuid4().hex; work = BASE / "jobs" / task; work.mkdir(parents=True)
    saved=[]
    try:
        for i,f in enumerate(files):
            if not f.filename: continue
            ext=Path(f.filename).suffix.lower()
            if tool not in {"images-to-pdf"} and ext != ".pdf": raise HTTPException(400,"PDF required")
            if tool == "images-to-pdf" and ext not in {".png",".jpg",".jpeg",".webp"}: raise HTTPException(400,"Image required")
            p=work/f"input-{i}{ext}"; p.write_bytes(await f.read()); saved.append(p)
        if tool == "merge":
            out=OUTPUT/f"{task}.pdf"; merge_pdfs(saved,out)
        elif tool == "split":
            targets=split_pdf(saved[0], work); shutil.make_archive(str(OUTPUT/task), "zip", work); return {"download_url":f"/api/v1/files/{task}.zip","count":len(targets)}
        elif tool == "extract":
            nums=[]
            for part in pages.split(","):
                if "-" in part:
                    a,b=map(int,part.split("-")); nums.extend(range(a,b+1))
                elif part.strip(): nums.append(int(part))
            if not nums: raise HTTPException(400,"pages is required, e.g. 1,3-5")
            out=OUTPUT/f"{task}.pdf"; extract_pages(saved[0],nums,out)
        elif tool == "rotate":
            if angle not in {90,180,270}: raise HTTPException(400,"angle must be 90, 180 or 270")
            out=OUTPUT/f"{task}.pdf"; rotate_pdf(saved[0],angle,out)
        elif tool == "compress":
            out=OUTPUT/f"{task}.pdf"; compress_pdf(saved[0],out)
        elif tool == "pdf-to-images":
            targets=pdf_to_images(saved[0],work,"png"); shutil.make_archive(str(OUTPUT/task),"zip",work); return {"download_url":f"/api/v1/files/{task}.zip","count":len(targets)}
        elif tool == "images-to-pdf":
            out=OUTPUT/f"{task}.pdf"; images_to_pdf(saved,out)
        elif tool == "pdf-to-word":
            try:
                layout = pdf_layout_analysis(saved[0])
            except json.JSONDecodeError as exc:
                raise HTTPException(502, f"Gemini returned invalid layout JSON: {exc.msg}") from exc
            out=OUTPUT/f"{task}.docx"
            render_editable_pdf(saved[0], layout, out)
        elif tool == "pdf-to-excel":
            prompt='''Extract all tables from this PDF. Return ONLY valid JSON: {"rows":[["cell1","cell2"]]}. Include column headers when present. Preserve values exactly; if there are multiple tables, append them separated by a blank row. Do not invent data.'''
            raw=pdf_to_structured(saved[0],prompt).strip(); cleaned=raw.removeprefix("```json").removesuffix("```").strip()
            try:
                data=json.loads(cleaned)
            except json.JSONDecodeError as exc:
                raise HTTPException(502, f"Gemini returned invalid JSON: {exc.msg}")
            out=OUTPUT/f"{task}.xlsx"; structured_to_xlsx(data,out)
        else: raise HTTPException(400,"Unsupported tool")
        return {"download_url":f"/api/v1/files/{out.name}","filename":out.name}
    except HTTPException:
        raise
    except Exception as exc:
        print(f"Tool {tool} failed: {type(exc).__name__}: {exc}", flush=True)
        if tool in {"pdf-to-word", "pdf-to-excel"}:
            message = str(exc)
            if "503" in message or "UNAVAILABLE" in message or "429" in message:
                raise HTTPException(503, f"Gemini temporarily unavailable. Please retry shortly. Details: {message}") from exc
            raise HTTPException(502, f"Document conversion failed: {message}") from exc
        raise HTTPException(500, f"Processing failed: {str(exc)}") from exc
    finally:
        shutil.rmtree(work, ignore_errors=True)

@app.get("/api/v1/files/{filename}")
def download(filename: str):
    safe=Path(filename).name; path=OUTPUT/safe
    if not path.exists(): raise HTTPException(404,"File not found")
    media="application/octet-stream"
    if path.suffix==".pdf": media="application/pdf"
    elif path.suffix==".docx": media="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif path.suffix==".xlsx": media="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif path.suffix==".zip": media="application/zip"
    return FileResponse(path,media_type=media,filename=safe)

app.mount("/", StaticFiles(directory="/app/web", html=True), name="web")
