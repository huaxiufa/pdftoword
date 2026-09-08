from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

from PIL import Image, ImageOps

SUPPORTED = {".jpg", ".jpeg", ".png", ".webp"}
Image.MAX_IMAGE_PIXELS = 100_000_000


def _safe_name(name: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).stem).strip("._") or "image"
    return stem[:100]


def _has_alpha(image: Image.Image) -> bool:
    return "A" in image.getbands() or "transparency" in image.info


def _resize(image: Image.Image, max_width: int | None, max_height: int | None) -> Image.Image:
    if not max_width and not max_height:
        return image
    w, h = image.size
    scale = 1.0
    if max_width and w > max_width:
        scale = min(scale, max_width / w)
    if max_height and h > max_height:
        scale = min(scale, max_height / h)
    if scale >= 1:
        return image
    size = (max(1, round(w * scale)), max(1, round(h * scale)))
    return image.resize(size, Image.Resampling.LANCZOS)


def _encode(image: Image.Image, fmt: str, quality: int, strip_metadata: bool) -> tuple[bytes, str]:
    out = io.BytesIO()
    save_kwargs: dict = {"optimize": True}
    if fmt == "JPEG":
        image = image.convert("RGB")
        save_kwargs.update(format="JPEG", quality=quality, progressive=True)
        ext = ".jpg"
    elif fmt == "WEBP":
        image = image.convert("RGBA" if _has_alpha(image) else "RGB")
        save_kwargs.update(format="WEBP", quality=quality, method=6)
        ext = ".webp"
    else:
        save_kwargs.update(format="PNG", compress_level=9)
        ext = ".png"
    if not strip_metadata:
        exif = image.getexif()
        if exif:
            save_kwargs["exif"] = exif.tobytes()
    image.save(out, **save_kwargs)
    return out.getvalue(), ext


def _compress_one(source: Path, options: dict) -> tuple[bytes, str, dict]:
    with Image.open(source) as opened:
        image = ImageOps.exif_transpose(opened)
        original_size = source.stat().st_size
        original_w, original_h = image.size
        image.load()
        fmt = str(options.get("format", "auto")).lower()
        if fmt == "auto":
            fmt = "png" if _has_alpha(image) and source.suffix.lower() == ".png" else "webp"
        if fmt == "jpg":
            output_fmt = "JPEG"
        elif fmt == "png":
            output_fmt = "PNG"
        else:
            output_fmt = "WEBP"
        quality = max(10, min(95, int(options.get("quality", 82))))
        max_width = int(options.get("max_width") or 0) or None
        max_height = int(options.get("max_height") or 0) or None
        image = _resize(image, max_width, max_height)
        target_kb = int(options.get("target_kb") or 0) or None
        data, ext = _encode(image, output_fmt, quality, bool(options.get("strip_metadata", True)))
        if target_kb and output_fmt in {"JPEG", "WEBP"} and len(data) > target_kb * 1024:
            q = quality
            while q > 25 and len(data) > target_kb * 1024:
                q -= 5
                data, ext = _encode(image, output_fmt, q, bool(options.get("strip_metadata", True)))
        ratio = max(0.0, 1 - len(data) / original_size) if original_size else 0.0
        return data, ext, {
            "name": source.name,
            "original_bytes": original_size,
            "output_bytes": len(data),
            "saved_percent": round(ratio * 100, 1),
            "original_dimensions": f"{original_w}×{original_h}",
            "output_dimensions": f"{image.width}×{image.height}",
        }


def compress_files(files: list[Path], output_zip: Path, options: dict, progress) -> dict:
    results: list[dict] = []
    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        total = max(1, len(files))
        used: set[str] = set()
        for index, source in enumerate(files, 1):
            progress(round((index - 1) / total * 90), f"正在压缩 {source.name}…")
            data, ext, stats = _compress_one(source, options)
            base = _safe_name(source.name)
            filename = base + ext
            counter = 2
            while filename.lower() in used:
                filename = f"{base}-{counter}{ext}"
                counter += 1
            used.add(filename.lower())
            archive.writestr(filename, data)
            results.append(stats)
    progress(100, "全部图片压缩完成")
    original = sum(x["original_bytes"] for x in results)
    output = sum(x["output_bytes"] for x in results)
    return {
        "count": len(results),
        "original_bytes": original,
        "output_bytes": output,
        "saved_percent": round((1 - output / original) * 100, 1) if original else 0,
        "results": results,
    }
