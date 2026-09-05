import json
import os
import sys
import threading
import time
from pathlib import Path

import fitz
from paddleocr import PPStructureV3

MODEL_NAMES = (
    "PP-DocBlockLayout", "PP-DocLayout_plus-L", "PP-LCNet_x1_0_textline_ori",
    "PP-OCRv5_server_det", "PP-OCRv5_server_rec", "PP-LCNet_x1_0_table_cls",
    "SLANeXt_wired", "SLANet_plus", "RT-DETR-L_wired_table_cell_det",
    "RT-DETR-L_wireless_table_cell_det",
)

def _write_progress(path: Path, stage: str, percent: int, message: str, **extra) -> None:
    payload = {"stage": stage, "percent": percent, "message": message}
    payload.update(extra)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

def _cached_models() -> list[str]:
    root = Path(os.getenv("PADDLEX_HOME", "/root/.paddlex")) / "official_models"
    return [n for n in MODEL_NAMES if (root / n).is_dir()] if root.exists() else []

def _start_heartbeat(progress_path: Path | None, total_pages: int):
    if progress_path is None:
        return None, None
    stop = threading.Event(); started = time.monotonic(); cached = _cached_models()
    def run():
        while not stop.wait(5):
            elapsed = int(time.monotonic() - started)
            current = _cached_models()
            if len(current) > len(cached): cached[:] = current
            msg = (f"模型已缓存，正在初始化 PP-StructureV3（{elapsed} 秒）…"
                   if len(cached) == len(MODEL_NAMES)
                   else f"正在加载模型（已缓存 {len(cached)}/{len(MODEL_NAMES)}，{elapsed} 秒）…")
            _write_progress(progress_path, "loading", min(4 + elapsed // 30, 7), msg,
                            current_page=0, total_pages=total_pages, init_seconds=elapsed,
                            cached_models=len(cached), total_models=len(MODEL_NAMES))
    threading.Thread(target=run, name="paddle-init-progress", daemon=True).start()
    return stop, started

def main() -> int:
    if len(sys.argv) not in (3, 4):
        print("usage: python -m app.paddle_worker input.pdf output_dir [progress.json]", file=sys.stderr)
        return 2
    pdf_path, output_dir = Path(sys.argv[1]), Path(sys.argv[2])
    progress_path = Path(sys.argv[3]) if len(sys.argv) == 4 else None
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        with fitz.open(pdf_path) as doc: total_pages = len(doc)
    except Exception: total_pages = 0

    cached = _cached_models()
    if progress_path:
        _write_progress(progress_path, "loading", 3, f"正在检查模型缓存（{len(cached)}/{len(MODEL_NAMES)}）…",
                        current_page=0, total_pages=total_pages, cached_models=len(cached), total_models=len(MODEL_NAMES))
    os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
    threads = os.getenv("PADDLE_CPU_THREADS", "4")
    os.environ.setdefault("OMP_NUM_THREADS", threads)
    os.environ.setdefault("MKL_NUM_THREADS", threads)
    if progress_path:
        _write_progress(progress_path, "loading", 4, "模型缓存已就绪，正在初始化 PP-StructureV3…",
                        current_page=0, total_pages=total_pages, cached_models=len(cached), total_models=len(MODEL_NAMES))
    heartbeat_stop, init_started = _start_heartbeat(progress_path, total_pages)
    try:
        pipeline = PPStructureV3(
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            device=os.getenv("PADDLE_DEVICE", "cpu"),
        )
    finally:
        if heartbeat_stop: heartbeat_stop.set()
    init_seconds = int(time.monotonic() - init_started) if init_started else 0
    if progress_path:
        _write_progress(progress_path, "ocr", 8, f"PP-StructureV3 已启动（模型初始化 {init_seconds} 秒），开始识别…",
                        current_page=0, total_pages=total_pages, init_seconds=init_seconds)

    results = pipeline.predict(input=str(pdf_path))
    for page_index, res in enumerate(results):
        payload = getattr(res, "json", None)
        if callable(payload): payload = payload()
        if not isinstance(payload, dict):
            temp = output_dir / f"page-{page_index:05d}"; temp.mkdir(exist_ok=True)
            res.save_to_json(save_path=str(temp))
            candidates = sorted(temp.glob("*.json"))
            if not candidates: raise RuntimeError(f"No JSON result for page {page_index}")
            payload = json.loads(candidates[0].read_text(encoding="utf-8"))
        img_map = getattr(res, "img", None)
        if callable(img_map): img_map = img_map()
        if isinstance(img_map, dict) and img_map.get("layout_det_res") is not None:
            img_map["layout_det_res"].save(output_dir / f"page-{page_index:05d}.png")
        (output_dir / f"page-{page_index:05d}.json").write_text(
            json.dumps(payload, ensure_ascii=False, default=lambda x: x.tolist()), encoding="utf-8")
        current = page_index + 1
        if progress_path:
            _write_progress(progress_path, "ocr", min(8 + int(current / max(total_pages, current) * 86), 94),
                            f"正在识别第 {current} / {total_pages or '?'} 页…", current_page=current, total_pages=total_pages)
    return 0

if __name__ == "__main__": raise SystemExit(main())
