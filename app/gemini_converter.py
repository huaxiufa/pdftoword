from __future__ import annotations

import html
import os
import re
import subprocess
import tempfile
from pathlib import Path
from urllib.parse import quote

import fitz
from google import genai

IMAGE_TOKEN_RE = re.compile(r"\[\[PDF_IMAGE_(\d+)\]\]")

PROMPT = """
Convert this PDF into one self-contained HTML document suitable for conversion to Microsoft Word.

Requirements:
1. Preserve the visual layout as closely as possible: page order, headings, paragraphs, tables, alignment, spacing, columns, and page breaks.
2. Make text real selectable HTML text, not screenshots.
3. Use simple Word/LibreOffice-friendly HTML and CSS. Do not use JavaScript, SVG, canvas, or external assets.
4. The source PDF contains {image_count} displayed images. You MUST insert every image placeholder exactly once, in visual reading order, at the exact place where that image belongs:
   [[PDF_IMAGE_1]], [[PDF_IMAGE_2]], ... [[PDF_IMAGE_{image_count}]]
5. Do not create, redraw, describe, or omit images. The placeholders will be replaced by the original PDF image files after your response.
6. Keep each placeholder inside the correct surrounding block (paragraph/table/cell/container) so the image remains near its original position.
7. If the PDF has multiple pages, create a page-break container between pages using CSS page-break-after: always.
8. Return ONLY complete HTML beginning with <!DOCTYPE html>. No Markdown fences and no explanation.
"""


def _image_blob(doc: fitz.Document, info: dict, image_map: dict[int, tuple[int, bytes]]) -> bytes | None:
    xref = int(info.get("xref") or 0)
    if xref and xref in image_map:
        return image_map[xref][1]
    if not xref:
        return None
    try:
        pix = fitz.Pixmap(doc, xref)
        if pix.alpha:
            pix = fitz.Pixmap(fitz.csRGB, pix)
        return pix.tobytes("png")
    except Exception:
        return None


def extract_displayed_images(pdf_path: Path, image_dir: Path) -> list[Path]:
    """Extract every displayed image occurrence in page/reading order."""
    image_dir.mkdir(parents=True, exist_ok=True)
    out: list[Path] = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            infos = []
            try:
                infos = page.get_image_info(xrefs=True)
            except Exception:
                pass

            # Fallback for older/odd PDFs.
            if not infos:
                for xref, _smask, _width, _height, *_ in page.get_images(full=True):
                    for rect in page.get_image_rects(xref, transform=True):
                        infos.append({"xref": xref, "bbox": tuple(rect[:4])})

            infos.sort(key=lambda item: (
                round(float(item.get("bbox", (0, 0, 0, 0))[1]), 2),
                round(float(item.get("bbox", (0, 0, 0, 0))[0]), 2),
            ))

            # xref -> (smask, bytes) cache for repeated images.
            cache: dict[int, tuple[int, bytes]] = {}
            for info in infos:
                xref = int(info.get("xref") or 0)
                if xref and xref not in cache:
                    blob = _image_blob(doc, info, {})
                    if blob:
                        cache[xref] = (0, blob)
                blob = _image_blob(doc, info, cache)
                if not blob:
                    continue
                path = image_dir / f"image-{len(out) + 1:04d}.png"
                path.write_bytes(blob)
                out.append(path)
    return out


def _clean_html(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:html)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)
    start = text.lower().find("<!doctype html>")
    if start >= 0:
        text = text[start:]
    return text.strip()


def _inject_images(document_html: str, image_paths: list[Path]) -> str:
    html_text = document_html
    for index, path in enumerate(image_paths, start=1):
        token = f"[[PDF_IMAGE_{index}]]"
        src = path.resolve().as_uri()
        tag = (
            f'<img src="{html.escape(src, quote=True)}" '
            'style="max-width:100%; height:auto; display:block;" />'
        )
        html_text = html_text.replace(token, tag)

    # If Gemini unexpectedly omitted a token, do not lose the original image.
    missing = [i for i in range(1, len(image_paths) + 1) if f"[[PDF_IMAGE_{i}]]" in html_text]
    if missing:
        extra = "".join(
            f'<p><img src="{html.escape(image_paths[i-1].resolve().as_uri(), quote=True)}" '
            'style="max-width:100%; height:auto;" /></p>'
            for i in missing
        )
        html_text = html_text.replace("</body>", extra + "</body>")
    return html_text


def convert_with_gemini(pdf_path: Path, out_docx: Path, progress=None) -> None:
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("未配置 GEMINI_API_KEY")

    model = os.getenv("GEMINI_MODEL", "gemini-3.7-flash")
    with tempfile.TemporaryDirectory(prefix="gemini-pdf-") as tmp:
        tmp_dir = Path(tmp)
        image_dir = tmp_dir / "images"
        if progress:
            progress({"stage": "images", "percent": 12, "message": "正在提取 PDF 原始图片…"})
        images = extract_displayed_images(pdf_path, image_dir)

        if progress:
            progress({"stage": "gemini_upload", "percent": 20, "message": "正在上传 PDF 到 Gemini…"})
        client = genai.Client(api_key=api_key)
        uploaded = client.files.upload(file=str(pdf_path))

        prompt = PROMPT.format(image_count=len(images))
        if progress:
            progress({"stage": "gemini", "percent": 35, "message": f"Gemini 正在重建排版（{model}）…"})
        response = client.models.generate_content(
            model=model,
            contents=[prompt, uploaded],
        )
        document_html = _clean_html(response.text or "")
        if not document_html.lower().startswith("<!doctype html>"):
            raise RuntimeError("Gemini 没有返回有效 HTML")

        document_html = _inject_images(document_html, images)
        html_path = tmp_dir / "document.html"
        html_path.write_text(document_html, encoding="utf-8")

        if progress:
            progress({"stage": "docx", "percent": 75, "message": "正在用 LibreOffice 生成 Word…"})
        out_docx.parent.mkdir(parents=True, exist_ok=True)
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to",
            "docx:Office Open XML Text",
            "--outdir",
            str(out_docx.parent),
            str(html_path),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=900)
        generated = out_docx.parent / "document.docx"
        if proc.returncode != 0 or not generated.exists():
            raise RuntimeError(f"LibreOffice 转换失败: {proc.stderr or proc.stdout}")
        if out_docx.exists():
            out_docx.unlink()
        generated.replace(out_docx)
