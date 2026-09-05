from pathlib import Path
import shutil
import sys
import tempfile


def _merge_docx_parts(parts, output):
    if not parts:
        raise RuntimeError("No Word pages were produced")
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
    from paddleocr import PPStructureV3

    pipeline = PPStructureV3(
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        engine="paddle",
    )

    temp_root = Path(tempfile.mkdtemp(prefix="paddleocr-word-", dir=output.parent))
    try:
        page_docs = []
        for index, result in enumerate(
            pipeline.predict_iter(input=str(source_pdf)), start=1
        ):
            print(f"PaddleOCR recovery: processing page {index}", flush=True)
            page_dir = temp_root / f"page-{index}"
            page_dir.mkdir(parents=True, exist_ok=True)
            result.save_to_word(save_path=str(page_dir))
            docs = sorted(page_dir.glob("*.docx"))
            if not docs:
                raise RuntimeError(f"No Word output for page {index}")
            page_docs.append(docs[0])
        _merge_docx_parts(page_docs, output)
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: python -m app.paddleocr_worker INPUT.pdf OUTPUT.docx")
    convert(Path(sys.argv[1]), Path(sys.argv[2]))
