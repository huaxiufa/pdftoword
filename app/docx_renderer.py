from __future__ import annotations

import math

import fitz

from .docx_renderer_v3 import V3Renderer


def _rotate_image_bytes(blob: bytes, angle: int) -> bytes:
    """Rotate extracted image bytes to match its PDF display transform."""
    if angle % 360 == 0:
        return blob
    try:
        pix = fitz.Pixmap(blob)
        # PyMuPDF Pixmap.flip_rotate uses numeric modes:
        # 1=90 CCW, 2=270 CCW, 3=180.
        mode = {90: 1, 180: 3, 270: 2}.get(angle % 360)
        if mode is None:
            return blob
        rotated = pix.flip_rotate(mode)
        return rotated.tobytes("png")
    except Exception:
        return blob


def _render_images_fixed(self, doc, page):
    """Render each displayed image occurrence using its real PDF transform.

    PyMuPDF's image blocks provide both the displayed bbox and the transformation
    matrix used to map the source image into that bbox. We use the bbox for the
    Word page position and rotate the extracted bytes from that matrix, rather
    than guessing from the raw XObject. This preserves repeated / transformed
    occurrences of the same image independently.
    """
    seen = set()
    page_rot = int(page.rotation or 0) % 360

    try:
        blocks = page.get_text("dict").get("blocks", [])
    except Exception:
        blocks = []

    for block in blocks:
        if block.get("type") != 1:
            continue

        bbox = block.get("bbox")
        blob = block.get("image")
        transform = block.get("transform")
        if not bbox or not blob:
            continue

        try:
            rect = fitz.Rect(bbox)
            if rect.width <= 1 or rect.height <= 1:
                continue

            # Image-block bboxes are in the page's unrotated coordinate system.
            # The Word section uses page.rect, so convert coordinates exactly once.
            placed = rect * page.rotation_matrix if page_rot else rect

            # The transform describes the source image's orientation inside its
            # displayed bbox. Word receives the extracted source bytes, so apply
            # that orientation to the bytes before placing them.
            angle = 0
            if transform is not None:
                try:
                    m = fitz.Matrix(transform)
                    angle = round(math.degrees(math.atan2(m.b, m.a))) % 360
                    # PDF image transforms can encode a reflected axis; for the
                    # normal PDF image cases handled here, normalize to quarter turns.
                    angle = min((0, 90, 180, 270), key=lambda a: abs(((angle - a + 180) % 360) - 180))
                except Exception:
                    angle = 0

            placed_blob = _rotate_image_bytes(blob, angle)

            key = (
                round(placed.x0, 2), round(placed.y0, 2),
                round(placed.x1, 2), round(placed.y1, 2),
                len(placed_blob), angle,
            )
            if key in seen:
                continue
            seen.add(key)

            area_ratio = (placed.width * placed.height) / max(
                1.0, page.rect.width * page.rect.height
            )
            self._add_positioned_image(
                doc,
                placed_blob,
                placed.x0,
                placed.y0,
                placed.width,
                placed.height,
                behind=area_ratio >= 0.80,
            )
        except Exception:
            continue


V3Renderer._render_images = _render_images_fixed
CoordinateDocxRenderer = V3Renderer

__all__ = ["CoordinateDocxRenderer"]
