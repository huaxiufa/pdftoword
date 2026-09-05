# PDF to Word — PP-OCRv5 + DocLayout-YOLO

Fast OCR/layout pipeline for editable DOCX.

## Pipeline

`PDF → page image → DocLayout-YOLO → PP-OCRv5 → coordinate JSON → editable DOCX`

- No `pdf2docx`
- No native PDF text reconstruction
- Original PDF images are extracted with PyMuPDF and placed at their page coordinates
- OCR text remains editable in Word
- CPU-friendly defaults and model caches

## Run

```bash
docker compose up --build
```

Open `http://localhost:8000`.

Models are cached in Docker volumes. Set `PADDLE_DEVICE=cuda:0` and install the matching PaddlePaddle GPU wheel for GPU deployment.

## Notes

DocLayout-YOLO is used only for page-level semantic regions. PP-OCRv5 supplies editable text boxes. Table regions are OCRed as editable text in this first implementation; the next iteration can add cell-level table reconstruction without changing the pipeline contract.
