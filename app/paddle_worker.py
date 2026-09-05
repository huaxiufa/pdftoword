import json
import os
import sys
from pathlib import Path

import fitz
from paddleocr import PPStructureV3


def _write_progress(path: Path, stage: str, percent: int, message: str, **extra) -> None:
    payload = {"stage": stage, "percent": percent, "message": message}
    payload.update(extra)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def main() -> int:
    if len(sys.argv) not in (3, 4):
        print("usage: python -m app.paddle_worker input.pdf output_dir [progress.json]", file=sys.stderr)
        return 2
    pdf_path = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])
    progress_path = Path(sys.argv[3]) if len(sys.argv) == 4 else None
    output_dir.mkdir(parents=True, exist_ok=True)

    total_pages = 0
    try:
        with fitz.open(pdf_path) as doc:
            total_pages = len(doc)
    except Exception:
        pass

    if progress_path:
        _write_progress(
            progress_path,
            "loading",
            3,
            "正在加载 PP-StructureV3 模型…",
            current_page=0,
            total_pages=total_pages,
        )

    pipeline = PPStructureV3(
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=True,
        device=os.getenv("PADDLE_DEVICE", "cpu"),
    )

    if progress_path:
        _write_progress(
            progress_path,
            "ocr",
            8,
            "PP-StructureV3 已启动，开始识别…",
            current_page=0,
            total_pages=total_pages,
        )

    results = pipeline.predict(input=str(pdf_path))
    for page_index, res in enumerate(results):
        payload = getattr(res, "json", None)
        if callable(payload):
            payload = payload()
        if not isinstance(payload, dict):
            temp = output_dir / f"page-{page_index:05d}"
            temp.mkdir(exist_ok=True)
            res.save_to_json(save_path=str(temp))
            candidates = sorted(temp.glob("*.json"))
            if not candidates:
                raise RuntimeError(f"No JSON result for page {page_index}")
            payload = json.loads(candidates[0].read_text(encoding="utf-8"))

        img_map = getattr(res, "img", None)
        if callable(img_map):
            img_map = img_map()
        if isinstance(img_map, dict):
            layout_img = img_map.get("layout_det_res")
            if layout_img is not None:
                layout_img.save(output_dir / f"page-{page_index:05d}.png")

        (output_dir / f"page-{page_index:05d}.json").write_text(
            json.dumps(payload, ensure_ascii=False, default=lambda x: x.tolist()),
            encoding="utf-8",
        )

        current = page_index + 1
        percent = 8 + int(current / max(total_pages, current) * 86)
        if progress_path:
            _write_progress(
                progress_path,
                "ocr",
                min(percent, 94),
                f"正在识别第 {current} / {total_pages or '?'} 页…",
                current_page=current,
                total_pages=total_pages,
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
