import re

from app.model import DocumentModel

BULLET_RE = re.compile(r"^(?:•|·|▪|‣|◦|●|-|–|—)\s*")


def _font_size(block):
    values = [s.size for line in block.lines for s in line.spans]
    return sum(values) / len(values) if values else 0


def _bold_ratio(block):
    values = [s.bold for line in block.lines for s in line.spans]
    return sum(values) / len(values) if values else 0


def analyze_layout(model: DocumentModel) -> DocumentModel:
    for page in model.pages:
        text_blocks = [b for b in page.blocks if b.kind == "text"]
        avg = sum(_font_size(b) for b in text_blocks) / max(len(text_blocks), 1)
        for block in text_blocks:
            text = block.text.strip()
            if not text:
                continue
            first = text.splitlines()[0].strip()
            if BULLET_RE.match(first):
                block.kind = "list"
                continue
            short = len(text.replace("\n", " ")) <= 90
            larger = avg and _font_size(block) >= avg * 1.12
            bold = _bold_ratio(block) >= 0.5
            if short and (larger or bold):
                block.kind = "heading"
    return model
