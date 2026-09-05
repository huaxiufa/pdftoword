import json
import os
import sys
from pathlib import Path

from paddleocr import PPStructureV3


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: python -m app.paddle_worker input.pdf output_dir", file=sys.stderr)
        return 2
    pdf_path = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])
    output_dir.mkdir(parents=True, exist_ok=True)

    # One Paddle process owns the pipeline. This avoids PDX re-initialization.
    pipeline = PPStructureV3(
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=True,
        device=os.getenv("PADDLE_DEVICE", "cpu"),
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

        # layout_det_res coordinates are in the processed page-image coordinate
        # system. Save the exact visualization image so the renderer can scale
        # those coordinates back to PDF points instead of guessing a DPI.
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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
