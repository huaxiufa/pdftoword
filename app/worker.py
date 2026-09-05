from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

import easyocr
import fitz
from doclayout_yolo import YOLOv10
from huggingface_hub import hf_hub_download


def save(path: Path, obj):
    path.write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")


def layout_boxes(result):
    boxes = getattr(result, "boxes", None)
    if boxes is None:
        return []
    xyxy = boxes.xyxy.cpu().tolist()
    conf = boxes.conf.cpu().tolist() if getattr(boxes, "conf", None) is not None else [1.0] * len(xyxy)
    cls = boxes.cls.cpu().tolist() if getattr(boxes, "cls", None) is not None else [0] * len(xyxy)
    names = getattr(result, "names", {}) or {}
    out = []
    for b, score, cls_id in zip(xyxy, conf, cls):
        ci = int(cls_id)
        out.append({
            "bbox": [float(x) for x in b],
            "score": float(score),
            "label": str(names.get(ci, ci)),
            "lines": [],
        })
    return out


def ocr_lines(reader: easyocr.Reader, image_path: Path):
    results = reader.readtext(
        str(image_path),
        detail=1,
        paragraph=False,
        width_ths=float(os.getenv("OCR_WIDTH_THS", "0.65")),
        link_threshold=float(os.getenv("OCR_LINK_THRESHOLD", "0.4")),
        low_text=float(os.getenv("OCR_LOW_TEXT", "0.3")),
        mag_ratio=float(os.getenv("OCR_MAG_RATIO", "1.0")),
    )
    out = []
    for polygon, text, score in results:
        text = str(text).strip()
        if not text or len(polygon) < 4:
            continue
        xs = [float(p[0]) for p in polygon]
        ys = [float(p[1]) for p in polygon]
        out.append({
            "text": text,
            "score": float(score),
            "bbox": [min(xs), min(ys), max(xs), max(ys)],
        })
    return out


def center_in(box, line):
    x0, y0, x1, y1 = box
    a, b, c, d = line["bbox"]
    cx, cy = (a + c) / 2, (b + d) / 2
    return x0 <= cx <= x1 and y0 <= cy <= y1


def load_layout_model():
    repo_id = os.getenv("DOCLAYOUT_MODEL", "juliozhao/DocLayout-YOLO-DocStructBench")
    filename = os.getenv("DOCLAYOUT_MODEL_FILE", "doclayout_yolo_docstructbench_imgsz1024.pt")
    model_path = hf_hub_download(repo_id=repo_id, filename=filename)
    return YOLOv10(model_path)


def load_ocr_reader():
    languages = [x.strip() for x in os.getenv("OCR_LANG", "ch_sim,en").split(",") if x.strip()]
    return easyocr.Reader(
        languages,
        gpu=False,
        model_storage_directory=os.getenv("EASYOCR_MODEL_DIR", "/root/.EasyOCR/model"),
        download_enabled=True,
        verbose=True,
    )


def main():
    if len(sys.argv) != 4:
        print("usage: python -m app.worker input.pdf output_dir progress.json", file=sys.stderr)
        return 2

    pdf_path, out_dir, progress_path = map(Path, sys.argv[1:])
    out_dir.mkdir(parents=True, exist_ok=True)

    with fitz.open(pdf_path) as pdf:
        total = len(pdf)
        pages = []
        scale = float(os.getenv("PDF_RENDER_SCALE", "1.75"))
        for i, page in enumerate(pdf):
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            image_path = out_dir / f"page-{i:05d}.png"
            pix.save(image_path)
            pages.append((i, page.rect.width, page.rect.height, image_path))

    threads = max(1, int(os.getenv("OCR_CPU_THREADS", "1")))
    os.environ["OMP_NUM_THREADS"] = str(threads)
    os.environ["MKL_NUM_THREADS"] = str(threads)
    os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")

    started = time.monotonic()
    save(progress_path, {
        "stage": "models",
        "percent": 2,
        "message": "正在加载 DocLayout-YOLO…",
        "current_page": 0,
        "total_pages": total,
    })
    layout_model = load_layout_model()

    save(progress_path, {
        "stage": "models",
        "percent": 4,
        "message": "正在加载 EasyOCR…",
        "current_page": 0,
        "total_pages": total,
    })
    reader = load_ocr_reader()
    init = int(time.monotonic() - started)
    save(progress_path, {
        "stage": "ocr",
        "percent": 5,
        "message": f"DocLayout-YOLO + EasyOCR 已加载（{init}s）",
        "current_page": 0,
        "total_pages": total,
    })

    for n, (page_index, page_w, page_h, image_path) in enumerate(pages, 1):
        layout_result = layout_model.predict(
            str(image_path),
            imgsz=int(os.getenv("DOCLAYOUT_IMGSZ", "1024")),
            conf=float(os.getenv("DOCLAYOUT_CONF", "0.20")),
            device="cpu",
            verbose=False,
        )[0]
        regions = layout_boxes(layout_result)
        lines = ocr_lines(reader, image_path)

        pix = fitz.Pixmap(image_path)
        iw, ih = pix.width, pix.height
        sx, sy = page_w / iw, page_h / ih

        for region in regions:
            b = region["bbox"]
            region["bbox"] = [b[0] * sx, b[1] * sy, b[2] * sx, b[3] * sy]

        orphan_lines = []
        for line in lines:
            b = line["bbox"]
            line["bbox"] = [b[0] * sx, b[1] * sy, b[2] * sx, b[3] * sy]
            matches = [r for r in regions if center_in(r["bbox"], line)]
            if matches:
                target = min(
                    matches,
                    key=lambda r: (r["bbox"][2] - r["bbox"][0]) * (r["bbox"][3] - r["bbox"][1]),
                )
                target["lines"].append(line)
            else:
                orphan_lines.append(line)

        if orphan_lines:
            regions.append({
                "bbox": [0.0, 0.0, page_w, page_h],
                "score": 1.0,
                "label": "text",
                "lines": orphan_lines,
                "synthetic": True,
            })

        regions.sort(key=lambda r: (r["bbox"][1], r["bbox"][0]))
        for region in regions:
            region["lines"].sort(key=lambda x: (x["bbox"][1], x["bbox"][0]))

        save(out_dir / f"page-{page_index:05d}.json", {
            "page_index": page_index,
            "page_width": page_w,
            "page_height": page_h,
            "image_width": iw,
            "image_height": ih,
            "regions": regions,
        })
        save(progress_path, {
            "stage": "ocr",
            "percent": 5 + int(n / max(total, 1) * 88),
            "message": f"正在处理第 {n} / {total} 页…",
            "current_page": n,
            "total_pages": total,
        })

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
