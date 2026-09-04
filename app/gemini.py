import json
import os
import re
import time
from pathlib import Path

from google import genai

MODEL = os.getenv("GEMINI_MODEL", "gemini-3.7-flash")
_client = None


def client():
    global _client
    if _client is None:
        key = os.getenv("GEMINI_API_KEY")
        if not key:
            raise RuntimeError("GEMINI_API_KEY is not configured")
        _client = genai.Client(api_key=key)
    return _client


def _is_retryable(exc: Exception) -> bool:
    status = getattr(exc, "status_code", None)
    return status in {429, 500, 502, 503, 504}


def pdf_to_structured(path: Path, prompt: str):
    c = client()
    uploaded = c.files.upload(file=str(path), config={"mime_type": "application/pdf"})
    try:
        last_error = None
        for attempt in range(4):
            try:
                response = c.models.generate_content(model=MODEL, contents=[uploaded, prompt])
                text = response.text or ""
                if not text.strip():
                    raise RuntimeError("Gemini returned an empty response")
                return text
            except Exception as exc:
                last_error = exc
                if not _is_retryable(exc) or attempt == 3:
                    raise
                time.sleep(2 ** attempt)
        raise last_error or RuntimeError("Gemini request failed")
    finally:
        try:
            c.files.delete(name=uploaded.name)
        except Exception:
            pass


def pdf_layout_analysis(path: Path):
    """Ask Gemini to visually understand every PDF page and return layout JSON."""
    prompt = r'''
You are the layout reconstruction engine for a PDF-to-Word converter.
Inspect the uploaded PDF visually, including text, photos, logos, tables, lines,
columns, spacing and page structure. Do NOT summarize the document.
Return ONLY valid JSON with this schema:
{
  "pages": [
    {
      "page": 1,
      "width_ratio": 1,
      "height_ratio": 1,
      "background": "white|other",
      "columns": [
        {
          "x": 0, "y": 0, "w": 1, "h": 1,
          "elements": [
            {"type":"text|image|table|line|shape", "x":0, "y":0, "w":1, "h":1,
             "text":"", "font_size":11, "bold":false, "italic":false,
             "align":"left|center|right", "image_index":0,
             "rows":[["cell"]]}
          ]
        }
      ]
    }
  ]
}
Coordinates must be normalized 0..1 relative to each page. Preserve reading order.
For tables, preserve every visible cell and row. For images, identify their position
and set image_index in top-to-bottom, left-to-right order on that page.
Do not omit decorative lines/shapes that materially affect the layout.
'''
    raw = pdf_to_structured(path, prompt)
    match = re.search(r"```(?:json)?\s*(.*?)\s*```", raw, re.S)
    cleaned = match.group(1) if match else raw.strip()
    return json.loads(cleaned)
