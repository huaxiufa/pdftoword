import fitz

from app.model import Block, DocumentModel, PageModel, Rect, TextLine, TextSpan


def parse_pdf(path) -> DocumentModel:
    pdf = fitz.open(path)
    model = DocumentModel()
    try:
        for page in pdf:
            page_model = PageModel(page.rect.width, page.rect.height)
            raw = page.get_text("dict")
            for raw_block in raw.get("blocks", []):
                if raw_block.get("type") == 0:
                    lines = []
                    for raw_line in raw_block.get("lines", []):
                        spans = []
                        for s in raw_line.get("spans", []):
                            if not s.get("text"):
                                continue
                            spans.append(TextSpan(
                                text=s["text"],
                                bbox=Rect(*s["bbox"]),
                                font=s.get("font", ""),
                                size=float(s.get("size", 0)),
                                flags=int(s.get("flags", 0)),
                            ))
                        if spans:
                            lines.append(TextLine(spans))
                    if lines:
                        page_model.blocks.append(Block(
                            kind="text",
                            bbox=Rect(*raw_block["bbox"]),
                            text="\n".join(x.text for x in lines),
                            lines=lines,
                        ))
                elif raw_block.get("type") == 1 and raw_block.get("image"):
                    page_model.blocks.append(Block(
                        kind="image",
                        bbox=Rect(*raw_block["bbox"]),
                        image=raw_block["image"],
                    ))
            page_model.blocks.sort(key=lambda b: (b.bbox.y0, b.bbox.x0))
            model.pages.append(page_model)
    finally:
        pdf.close()
    return model
