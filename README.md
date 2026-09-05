# PDF → Word

A PDF-to-DOCX service built around **PaddleOCR PP-StructureV3** and PDF-native geometry.

## Design

`PDF → PP-StructureV3 → layout JSON + page-image size → coordinate DOCX`

- Text: OCR output becomes editable Word text boxes positioned from layout coordinates.
- Images: embedded PDF images are extracted with PyMuPDF and inserted as floating Word images at their PDF coordinates.
- Tables: PP-StructureV3 table HTML is rebuilt as editable Word tables inside positioned containers.
- Scanned PDFs: the OCR/layout path works without `pdf2docx`.
- Full-page scan images are not inserted as a background image, so OCR text remains editable.
- PaddleOCR runs in a dedicated subprocess to avoid PDX global-initialization conflicts in the API process.

This is a fidelity-oriented reconstruction, not a promise of pixel-perfect Word output. Word's layout engine and editable-object model are different from PDF, so complex typography, overlapping objects, fonts, and unusual tables can still differ.

## Run

```bash
docker compose up --build
```

Open `http://localhost:8000`, choose a PDF, and start conversion.

## API

`POST /convert` with multipart field `file` returns a `.docx`.

`GET /health` returns service status.

## Configuration

- `PADDLE_DEVICE=cpu` (set to `gpu` when the container has a compatible GPU runtime)
- `PADDLEOCR_TIMEOUT=1800`
- `MAX_UPLOAD_MB=50`

## Why this route

PP-StructureV3 exposes layout blocks, OCR results, table recognition and structured JSON, and its result object also supports DOCX export. This project deliberately uses the structured JSON rather than relying on the built-in Word exporter, because the latter does not expose enough control over PDF-native coordinates for this fidelity target.
