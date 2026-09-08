# PDF → Word V5

V5 is the final visual-fidelity renderer.

## Architecture

**PDF → PyMuPDF page rendering → exact-size DOCX page image**

Instead of rebuilding the PDF's text, images and drawing objects independently, V5 asks MuPDF to render each complete PDF page exactly as a PDF viewer would see it, then places that rendered page at the exact PDF page dimensions inside a matching Word section.

This is intentional: PDF and DOCX use different layout engines, so independently rebuilding every font metric, transparency mask, clipping path, avatar crop, vector shape and image transform can introduce visible differences. PyMuPDF's page renderer supports DPI, colorspace, transparency and other rendering controls, while `get_image_info()` is available when object-level inspection is needed.

## What V5 prioritizes

- Visual fidelity over editable-text fidelity.
- Transparent PDF images and soft masks are composited by MuPDF before insertion into Word.
- Processed avatars, clipping, masks, rotation and image transforms are preserved as they appear in the PDF.
- Fonts, glyph shapes, line spacing, kerning, vector graphics and other PDF drawing operations are preserved visually by page rendering.
- Exact PDF page width / height is applied to each Word section.
- No Gemini dependency.
- No LibreOffice dependency for conversion.

## Important trade-off

The page content is embedded as a high-resolution page image, so the resulting Word document is **visually faithful but not equivalent to a natively editable Word document**. This is the deliberate final trade-off for the case where preserving the original PDF appearance is more important than reflowable/editable text.

Default rendering is 220 DPI. You can change it with:

```env
PDF_RENDER_DPI=220
```

The renderer clamps this setting to 120–300 DPI.

## Run

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
