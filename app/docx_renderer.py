from __future__ import annotations

import math

import fitz

from .docx_renderer_v3 import V3Renderer


# PyMuPDF image coordinates are reported in unrotated page space, while
# Page.rect reflects page rotation. The V3 renderer previously used the raw
# image rectangle and raw image bytes, which makes rotated / transformed PDF
# images appear displaced or rotated in Word. Patch only the image layer here
# so the rest of the V3 text/table renderer remains unchanged.
#
# Use the documented integer transformation modes rather than relying on
# Pixmap.ROTATE_* attributes: some installed PyMuPDF builds expose the modes
# only as documented numeric values.
_ROTATE_90 = 1
_ROTATE_270 = 2
_ROTATE_180 = 3


def _render_images_fixed(self, doc, page):
    seen = set()
    page_rot = int(page.rotation or 0) % 360
    rot_modes = {
        90: _ROTATE_90,
        180: _ROTATE_180,
        270: _ROTATE_270,
    }

    for image in page.get_images(full=True):
        xref = image[0]
        try:
            occurrences = page.get_image_rects(xref, transform=True)
        except Exception:
            occurrences = [(r, None) for r in page.get_image_rects(xref)]

        for rect, matrix in occurrences:
            if rect.width <= 1 or rect.height <= 1:
                continue

            # Coordinates from get_image_rects() are unrotated. Convert them
            # into the same page coordinate system used by Word's page-sized
            # section when the PDF itself has a rotation flag.
            if page_rot:
                try:
                    placed = rect * page.rotation_matrix
                except Exception:
                    placed = rect
            else:
                placed = rect

            key = (
                xref,
                round(placed.x0, 2), round(placed.y0, 2),
                round(placed.x1, 2), round(placed.y1, 2),
            )
            if key in seen:
                continue
            seen.add(key)

            try:
                pix = fitz.Pixmap(self.pdf, xref)
                if pix.alpha:
                    pix = fitz.Pixmap(fitz.csRGB, pix)

                # Preserve the image transformation matrix. PDF images are
                # often rotated/scaled at placement time; extracting the raw
                # xref alone loses that transformation.
                if matrix is not None:
                    try:
                        angle = round(math.degrees(math.atan2(matrix.b, matrix.a))) % 360
                        if angle in rot_modes:
                            pix = pix.flip_rotate(rot_modes[angle])
                    except Exception:
                        pass

                # A page rotation also rotates the visual content.
                if page_rot in rot_modes:
                    try:
                        pix = pix.flip_rotate(rot_modes[page_rot])
                    except Exception:
                        pass

                blob = pix.tobytes("png")
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
