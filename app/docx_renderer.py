from __future__ import annotations

import math

import fitz

from .docx_renderer_v3 import V3Renderer


def _quarter_turn_from_transform(transform) -> int:
    """Return the clockwise quarter-turn encoded by a PDF image transform.

    PyMuPDF matrices use ``[a, b, c, d, e, f]``.  In image transforms,
    ``Matrix(0, -s, s, 0, ...)`` is a 90-degree *clockwise* image rotation.
    Therefore the displayed clockwise angle is ``-atan2(b, a)`` rather than
    ``atan2(b, a)``.  The previous implementation used the opposite sign,
    which is why rotated PDF images could appear rotated the wrong way in Word.
    """
    if transform is None:
        return 0
    try:
        m = fitz.Matrix(transform)
        raw = math.degrees(math.atan2(m.b, m.a))
        clockwise = (-raw) % 360
        return min(
            (0, 90, 180, 270),
            key=lambda a: abs(((clockwise - a + 180) % 360) - 180),
        )
    except Exception:
        return 0


def _rotate_image_bytes(blob: bytes, clockwise_angle: int) -> bytes:
    """Bake the PDF image rotation into the image bytes.

    ``Pixmap.flip_rotate`` uses numeric modes where 1 is 90 CCW and 2 is
    270 CCW (90 clockwise).  The PDF transform angle passed here is clockwise,
    so 90 and 270 intentionally map to the opposite PyMuPDF modes.
    """
    angle = clockwise_angle % 360
    if angle == 0:
        return blob
    try:
        pix = fitz.Pixmap(blob)
        mode = {90: 2, 180: 3, 270: 1}.get(angle)
        if mode is None:
            return blob
        rotated = pix.flip_rotate(mode)
        return rotated.tobytes("png")
    except Exception:
        return blob


def _render_images_fixed(self, doc, page):
    """Render every displayed image at its exact PDF bbox and orientation.

    ``Page.get_text("dict")`` is deliberately used instead of ``get_images``:
    PyMuPDF reports one image block for every displayed occurrence, including
    repeated XObjects, and each block contains both the displayed ``bbox`` and
    the image transformation matrix.  Coordinates returned by extraction are in
    unrotated page space; ``page.rotation_matrix`` converts them to the visible
    page space used by the Word section.  This applies page rotation exactly
    once and image rotation separately.
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

            # PyMuPDF extraction coordinates are unrotated.  Word's section
            # dimensions come from page.rect, which reflects page rotation.
            # Convert the bbox exactly once into that visible coordinate space.
            placed = rect * page.rotation_matrix if page_rot else rect

            # Bake the image's own PDF transform into its bytes.  Do NOT rotate
            # the bbox again: the bbox already describes the displayed footprint.
            image_angle = _quarter_turn_from_transform(transform)
            placed_blob = _rotate_image_bytes(blob, image_angle)

            # Keep distinct occurrences.  The same source image can legitimately
            # appear multiple times at different locations / transforms.
            key = (
                round(placed.x0, 2), round(placed.y0, 2),
                round(placed.x1, 2), round(placed.y1, 2),
                image_angle,
                len(placed_blob),
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
