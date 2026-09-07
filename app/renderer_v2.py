from __future__ import annotations

from pathlib import Path

class CoordinateDocxRenderer:
    """Placeholder for the new DrawingML renderer."""
    def __init__(self, pdf_path: Path):
        self.pdf_path = pdf_path

    def render(self, pages, out_path: Path) -> None:
        raise NotImplementedError("renderer v2 is being installed")
