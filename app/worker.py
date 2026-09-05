from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

import fitz
from doclayout_yolo import YOLOv10
from paddleocr import PaddleOCR


def save(path: Path, obj):
    path.write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")


def as_dict(result):
    data = getattr(result, "json", None)
    if callable(data):
        data = data()
    if isinstance(data, dict):
        return data.get("res", data)
    if isinstance(result, dict):
        return result.get("res", result)
    return {}


def layout_boxes(result):
    # DocLayout-YOLO follows the Ultralytics-style result API.
    boxes = getattr(result, "boxes", None)
    if boxes is None:
        return []
    xyxy = boxes.xyxy.cpu().tolist()
    conf = boxes.conf.cpu().tolist() if getattr(boxes, "conf", None) is not None else [1.0] * len(xyxy)
    cls = boxes.cls.cpu().tolist() if getattr(boxes, "cls", None) is not None else [0] * len(xyxy)
    names = getattr(result, "names", {}) or {}
    out = []
    for b, s, c in zip(xyxy, conf, cls):
        ci = int(c)
        out.append({"bbox": [float(x) for x in b], "score": float(s), "label": str(names.get(ci, ci))})
    return out


def ocr_lines(result):
    data = as_dict(result)
    texts = data.get("rec_texts") or []
    boxes = data.get("rec_boxes") or []
    scores = data.get("rec_scores") or []
    out = []
    for i, text in enumerate(texts):
        text = str(text).strip()
        if not text:
            continue
        b = boxes[i] if i < len(boxes) else None
        if b is None or len(b) < 4:
            continue
        out.append({"text": text, "score": float(scores[i]) if i < len(scores) else 1.0, "bbox": [float(x) for x in b[:4]]})
    return out


def center_in(box, line):
    x0,y0,x1,y1 = box
    a,b,c,d = line["bbox"]
    cx, cy = (a+c)/2, (b+d)/2
    return x0 <= cx <= x1 and y0 <= cy <= y1


def main():
    if len(sys.argv) != 4:
        print("usage: python -m app.worker input.pdf output_dir progress.json", file=sys.stderr)
        return 2
    pdf_path, out_dir, progress_path = map(Path, sys.argv[1:])
    out_dir.mkdir(parents=True, exist_ok=True)
    with fitz.open(pdf_path) as pdf:
        total = len(pdf)
        pages = []
        for i, page in enumerate(pdf):
            pix = page.get_pixmap(matrix=fitz.Matrix(1.75, 1.75), alpha=False)
            image_path = out_dir / f"page-{i:05d}.png"
            pix.save(image_path)
            pages.append((i, page.rect.width, page.rect.height, image_path))

    device = os.getenv("PADDLE_DEVICE", "cpu")
    threads = int(os.getenv("OCR_CPU_THREADS", os.getenv("OMP_NUM_THREADS", "4")))
    os.environ.setdefault("OMP_NUM_THREADS", str(threads))
    os.environ.setdefault("MKL_NUM_THREADS", str(threads))

    started = time.monotonic()
    layout_model = YOLOv10.from_pretrained(os.getenv("DOCLAYOUT_MODEL", "juliozhao/DocLayout-YOLO-DocStructBench"))
    ocr = PaddleOCR(
        ocr_version="PP-OCRv5",
        lang=os.getenv("OCR_LANG", "ch"),
        device=device,
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        cpu_threads=threads,
    )
    init = int(time.monotonic() - started)
    save(progress_path, {"stage":"ocr","percent":5,"message":f"PP-OCRv5 + DocLayout-YOLO 已加载（{init}s）","current_page":0,"total_pages":total})

    for n, (page_index, page_w, page_h, image_path) in enumerate(pages, 1):
        layout_result = layout_model.predict(str(image_path), imgsz=int(os.getenv("DOCLAYOUT_IMGSZ","1024")), conf=float(os.getenv("DOCLAYOUT_CONF","0.20")), device=device, verbose=False)[0]
        regions = layout_boxes(layout_result)
        ocr_result = next(iter(ocr.predict(str(image_path))))
        lines = ocr_lines(ocr_result)
        iw, ih = fitz.Pixmap(image_path).width, fitz.Pixmap(image_path).height
        sx, sy = page_w / iw, page_h / ih
        for r in regions:
            r["bbox"] = [r["bbox"][0]*sx, r["bbox"][1]*sy, r["bbox"][2]*sx, r["bbox"][3]*sy]
            r["lines"] = []
        for line in lines:
            line["bbox"] = [line["bbox"][0]*sx, line["bbox"][1]*sy, line["bbox"][2]*sx, line["bbox"][3]*sy]
            matches = [r for r in regions if center_in(r["bbox"], line)]
            target = min(matches, key=lambda r: (r["bbox"][2]-r["bbox"][0])*(r["bbox"][3]-r["bbox"][1])) if matches else None
            if target is not None:
                target["lines"].append(line)
        regions.sort(key=lambda r: (r["bbox"][1], r["bbox"][0]))
        for r in regions:
            r["lines"].sort(key=lambda x: (x["bbox"][1], x["bbox"][0]))
        save(out_dir / f"page-{page_index:05d}.json", {"page_index":page_index,"page_width":page_w,"page_height":page_h,"image_width":iw,"image_height":ih,"regions":regions})
        save(progress_path, {"stage":"ocr","percent":5+int(n/max(total,1)*88),"message":f"正在处理第 {n} / {total} 页…","current_page":n,"total_pages":total})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
