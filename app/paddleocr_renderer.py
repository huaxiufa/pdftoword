from pathlib import Path
import shutil
import tempfile


_pipeline = None


def _build_pipeline():
    global _pipeline
    if _pipeline is None:
        # Use the official PaddleOCR-VL Python entry point. Calling
        # paddlex.create_pipeline() directly can trigger PaddleX's global
        # initialization more than once in a long-running API process.
        from paddleocr import PaddleOCRVL

        _pipeline = PaddleOCRVL(pipeline_version="v1.6")
    return _pipeline


def _merge_docx_parts(parts, output):
    if not parts:
        raise RuntimeError("PaddleOCR did not produce any Word pages")
    if len(parts) == 1:
        shutil.copyfile(parts[0], output)
        return

    try:
        from docxcompose.composer import Composer
        from docx import Document
    except ImportError as exc:
        raise RuntimeError("docxcompose is required to merge multi-page PaddleOCR output") from exc

    master = Document(str(parts[0]))
    composer = Composer(master)
    for part in parts[1:]:
        composer.append(Document(str(part)))
    composer.save(str(output))


def render_editable_pdf(source_pdf, output):
    """Convert a PDF to editable DOCX using the official PaddleOCR-VL pipeline.

    PaddleOCR-VL performs layout detection, reading-order handling, document
    element recognition and has an official DOCX exporter. PDF pages are
    processed individually by PaddleOCR, so the generated page documents are
    merged back into one DOCX while preserving each page's Word layout.
    """
    source_pdf = Path(source_pdf)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)

    pipeline = _build_pipeline()
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
