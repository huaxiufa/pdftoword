# PDF → Word V4

V4 abandons the Gemini-generated HTML renderer and rebuilds the DOCX from the PDF's own objects.

## Architecture

**PDF → PyMuPDF → Text/Image objects → DOCX V4**

PyMuPDF provides page geometry, text blocks / lines / spans, image bytes and image bounding boxes. The renderer maps those objects into Word sections, paragraphs and floating image anchors. This keeps the PDF as the source of truth for page size, object position and image placement.

PyMuPDF documents that `Page.get_text("dict")` exposes text/image blocks and image bounding boxes, while `Page.get_image_rects()` provides accurate displayed image rectangles. Coordinates are based on the unrotated page coordinate system, so V4 keeps the PDF geometry explicit instead of asking an LLM to recreate it. citeturn0search0turn0search2turn0search6

## What V4 handles

- Exact PDF page width / height per Word section.
- Text blocks, lines and spans with font name, size, color, bold and italic when available.
- Original displayed images, including inline PDF images that do not appear in `Page.get_images()`.
- Image bounding boxes and basic rotation handling.
- Image clipping when an image extends outside the visible page.
- Floating Word image anchors using page-relative coordinates.
- No Gemini dependency and no LibreOffice dependency for conversion.

## Limitations

PDF and DOCX have different layout models. V4 is designed to preserve geometry and editable text, but Word's text metrics, font availability, wrapping and floating-object behavior can still differ from the PDF. Multi-column reading order and complex vector graphics require additional heuristics. PyMuPDF itself notes that PDF text extraction order depends on how the source PDF was created. citeturn0search7

Scanned PDFs without a text layer are not yet OCR-first in this V4 baseline. OCR can be added later using PyMuPDF's OCR text page support. citeturn0search4

## Run

Copy `.env.example` to `.env` if needed and keep `MAX_UPLOAD_MB` at the desired limit.

```bash
docker compose up -d --build
```

Open:

```text
http://localhost:8000
```

## Update an existing installation

```bash
git pull origin main
docker compose up -d --build
```
