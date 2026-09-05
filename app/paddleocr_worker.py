from pathlib import Path
import shutil
import sys
import tempfile


def _merge_docx_parts(parts, output):
    if not parts:
        raise RuntimeError("PaddleOCR did not produce any Word pages")
    if len(parts) == 1:
        shutil.copyfile(parts[0], output)
        return

    from docxcompose.composer import Composer
    from docx import Document

    master = Document(str(parts[0]))
    composer = Composer(master)
    for part in parts[1:]:
        composer.append(Document(str(part)))
    composer.save(str(output))


def convert(source_pdf: Path, output: Path):
    # Keep all Paddle/PaddleX imports and initialization inside this process.
    # The FastAPI process must never import PaddleOCR, avoiding PDX global-state
    # collisions with other libraries or repeated request handling.
    from paddleocr import PaddleOCRVL

    pipeline = PaddleOCRVL(pipeline_version="v1.6")
    temp_root = Path(tempfile.mkdtemp(prefix="paddleocr-word-", dir=output.parent))
    try:
        pages = list(pipeline.predict(input=str(source_pdf)))
        if not pages:
            raise RuntimeError("PaddleOCR returned no page results")

        page_docs = []
        for index, result in enumerate(pages, start=1):
            page_dir = temp_root / f"page-{index}"
            page_dir.mkdir(parents=True, exist_ok=True)
            result.save_to_word(save_path=str(page_dir))
            docs = sorted(page_dir.glob("*.docx"))
            if not docs:
                raise RuntimeError(f"PaddleOCR did not export Word for page {index}")
            page_docs.append(docs[0])

        _merge_docx_parts(page_docs, output)
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: python -m app.paddleocr_worker INPUT.pdf OUTPUT.docx")
    convert(Path(sys.argv[1]), Path(sys.argv[2]))
