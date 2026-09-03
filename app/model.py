from dataclasses import dataclass, field


@dataclass
class Rect:
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def width(self):
        return self.x1 - self.x0

    @property
    def height(self):
        return self.y1 - self.y0


@dataclass
class TextSpan:
    text: str
    bbox: Rect
    font: str
    size: float
    flags: int = 0

    @property
    def bold(self):
        return bool(self.flags & 16)

    @property
    def italic(self):
        return bool(self.flags & 2)


@dataclass
class TextLine:
    spans: list[TextSpan] = field(default_factory=list)

    @property
    def text(self):
        return "".join(s.text for s in self.spans).strip()

    @property
    def bbox(self):
        return Rect(
            min(s.bbox.x0 for s in self.spans),
            min(s.bbox.y0 for s in self.spans),
            max(s.bbox.x1 for s in self.spans),
            max(s.bbox.y1 for s in self.spans),
        )


@dataclass
class Block:
    kind: str
    bbox: Rect
    text: str = ""
    lines: list[TextLine] = field(default_factory=list)
    image: bytes | None = None


@dataclass
class PageModel:
    width: float
    height: float
    blocks: list[Block] = field(default_factory=list)


@dataclass
class DocumentModel:
    pages: list[PageModel] = field(default_factory=list)
