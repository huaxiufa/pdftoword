from __future__ import annotations

import fitz

from .docx_renderer_v3 import V3Renderer


def _render_images_fixed(self, doc, page):
    """Render displayed PDF image blocks without reconstructing transforms."""
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
        if not bbox or not blob:
            continue

        try:
            rect = fitz.Rect(bbox)
            if rect.width <= 1 or rect.height <= 1:
                continue

            # PyMuPDF returns extraction coordinates in unrotated page space.
            # The renderer creates the Word section using page.rect, which
            # reflects the visible rotation, so apply rotation exactly once.
            placed = rect * page.rotation_matrix if page_rot else rect

            key = (
                round(placed.x0, 2), round(placed.y0, 2),
                round(placed.x1, 2), round(placed.y1, 2),
                len(blob),
            )
            if key in seen:
                continue
            seen.add(key)

            area_ratio = (placed.width * placed.height) / max(
                1.0, page.rect.width * page.rect.height
            )
            self._add_positioned_image(
                doc,
                blob,
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
