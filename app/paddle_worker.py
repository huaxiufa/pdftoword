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

    # One Paddle process owns the pipeline. This avoids PDX re-initialization
    # when the web server handles multiple requests.
    pipeline = PPStructureV3(
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=True,
        device=os.getenv("PADDLE_DEVICE", "cpu"),
    )
    results = pipeline.predict(input=str(pdf_path))
    for page_index, res in enumerate(results):
        payload = _to_jsonable(getattr(res, "json", None))
        if payload is None:
            temp = output_dir / f"page-{page_index:05d}"
            temp.mkdir(exist_ok=True)
            res.save_to_json(save_path=str(temp))
            candidates = sorted(temp.glob("*.json"))
            if not candidates:
                raise RuntimeError(f"No JSON result for page {page_index}")
            payload = json.loads(candidates[0].read_text(encoding="utf-8"))
        (output_dir / f"page-{page_index:05d}.json").write_text(
            json.dumps(payload, ensure_ascii=False), encoding="utf-8"
        )
    return 0


def _to_jsonable(value):
    if value is None:
        return None
    if callable(value):
        try:
            value = value()
        except TypeError:
            return None
    if isinstance(value, dict):
        return value
    return None


if __name__ == "__main__":
    raise SystemExit(main())
