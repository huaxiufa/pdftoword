from __future__ import annotations

import html
import os
import re
import subprocess
import tempfile
from pathlib import Path

import fitz
from google import genai

PROMPT = """
Convert this PDF into one complete HTML document suitable for Microsoft Word conversion.

Requirements:
1. Preserve the visual layout as closely as possible: page order, headings, paragraphs, tables, alignment, spacing, columns, and page breaks.
2. Make text real selectable HTML text, not screenshots.
3. Use simple Word/LibreOffice-friendly HTML and CSS. Do not use JavaScript, SVG, canvas, or external assets.
4. The source PDF contains {image_count} displayed images. Insert every placeholder exactly once, in visual reading order, at the location where that image belongs:
   [[PDF_IMAGE_1]], [[PDF_IMAGE_2]], ... [[PDF_IMAGE_{image_count}]]
5. Never redraw, describe, replace, or omit an image. The placeholders are replaced later by the original PDF image files.
6. Keep each placeholder inside the correct paragraph, table cell, or layout container.
7. For multiple pages, preserve page boundaries with CSS page-break-after: always.
8. Return ONLY complete HTML beginning with <!DOCTYPE html>. No Markdown fences and no explanation.
"""


def _pixmap_png(doc: fitz.Document, xref: int) -> bytes | None:
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
    """Extract image occurrences, including inline images that have no xref."""
    image_dir.mkdir(parents=True, exist_ok=True)
    out: list[Path] = []

    with fitz.open(pdf_path) as doc:
        for page in doc:
            occurrences: list[tuple[float, float, bytes]] = []
            seen_bbox: set[tuple[float, float, float, float]] = set()

            # TextPage image blocks contain the actual bytes and also cover inline images.
            try:
                blocks = page.get_text("dict").get("blocks", [])
            except Exception:
                blocks = []
            for block in blocks:
                if block.get("type") != 1 or not block.get("image"):
                    continue
                bbox = block.get("bbox")
                if not bbox:
                    continue
                key = tuple(round(float(v), 2) for v in bbox)
                if key in seen_bbox:
                    continue
                seen_bbox.add(key)
                occurrences.append((float(bbox[1]), float(bbox[0]), block["image"]))

            # get_image_info catches displayed XObjects that may not expose bytes in text blocks.
            try:
                infos = page.get_image_info(xrefs=True)
            except Exception:
                infos = []
            if not infos:
                for xref, _smask, _width, _height, *_ in page.get_images(full=True):
                    for rect in page.get_image_rects(xref, transform=True):
                        infos.append({"xref": xref, "bbox": tuple(rect[:4])})

            for info in infos:
                bbox = info.get("bbox")
                if not bbox:
                    continue
                key = tuple(round(float(v), 2) for v in bbox)
                if key in seen_bbox:
                    continue
                blob = _pixmap_png(doc, int(info.get("xref") or 0))
                if not blob:
                    continue
                seen_bbox.add(key)
                occurrences.append((float(bbox[1]), float(bbox[0]), blob))

            occurrences.sort(key=lambda item: (round(item[0], 2), round(item[1], 2)))
            for _, _, blob in occurrences:
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
    present: set[int] = set()
    for index, path in enumerate(image_paths, start=1):
        token = f"[[PDF_IMAGE_{index}]]"
        if token not in html_text:
            continue
        present.add(index)
        src = path.resolve().as_uri()
        tag = (
            f'<img src="{html.escape(src, quote=True)}" '
            'style="max-width:100%; height:auto; display:block;" />'
        )
        html_text = html_text.replace(token, tag, 1)

    # Safety net: if Gemini omitted any placeholder, append the original image instead of losing it.
    missing = [i for i in range(1, len(image_paths) + 1) if i not in present]
    if missing:
        extra = "".join(
            f'<p><img src="{html.escape(image_paths[i-1].resolve().as_uri(), quote=True)}" '
            'style="max-width:100%; height:auto; display:block;" /></p>'
            for i in missing
        )
        if "</body>" in html_text.lower():
            html_text = re.sub(r"</body>", extra + "</body>", html_text, count=1, flags=re.I)
        else:
            html_text += extra
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
        response = client.models.generate_content(model=model, contents=[prompt, uploaded])
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
            "libreoffice", "--headless", "--convert-to",
            "docx:Office Open XML Text", "--outdir", str(out_docx.parent), str(html_path)
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=900)
        generated = out_docx.parent / "document.docx"
        if proc.returncode != 0 or not generated.exists():
            raise RuntimeError(f"LibreOffice 转换失败: {proc.stderr or proc.stdout}")
        if out_docx.exists():
            out_docx.unlink()
        generated.replace(out_docx)
