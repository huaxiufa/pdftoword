import json
import os
import re
import time
from pathlib import Path

from google import genai

MODEL = os.getenv("GEMINI_MODEL", "gemini-3.7-flash")
FALLBACK_MODEL = os.getenv("GEMINI_FALLBACK_MODEL", "gemini-3.6-flash")
_client = None


def client():
    global _client
    if _client is None:
        key = os.getenv("GEMINI_API_KEY")
        if not key:
            raise RuntimeError("GEMINI_API_KEY is not configured")
        _client = genai.Client(api_key=key)
    return _client


def _status_code(exc: Exception):
    """Extract an HTTP/API status code across google-genai exception variants."""
    for value in (
        getattr(exc, "status_code", None),
        getattr(exc, "code", None),
        getattr(getattr(exc, "response", None), "status_code", None),
    ):
        if isinstance(value, int):
            return value
        if isinstance(value, str):
            match = re.search(r"\b(429|500|502|503|504)\b", value)
            if match:
                return int(match.group(1))

    match = re.search(r"\b(429|500|502|503|504)\b", str(exc))
    return int(match.group(1)) if match else None


def _is_retryable(exc: Exception) -> bool:
    return _status_code(exc) in {429, 500, 502, 503, 504}


def _model_candidates():
    """Return configured models in priority order, without duplicates."""
    configured = os.getenv("GEMINI_MODELS", "")
    values = [item.strip() for item in configured.split(",") if item.strip()]
    if not values:
        values = [MODEL, FALLBACK_MODEL]
    elif MODEL:
        values.insert(0, MODEL)
        if FALLBACK_MODEL:
            values.append(FALLBACK_MODEL)

    result = []
    for value in values:
        if value and value not in result:
            result.append(value)
    return result


def _generate_with_fallback(c, uploaded, prompt):
    models = _model_candidates()
    last_error = None

    for model_index, model in enumerate(models):
        # A busy Gemini model should not block the conversion for a long time.
        # Two quick retries are enough before moving to the next configured model.
        max_attempts = 2
        for attempt in range(max_attempts):
            try:
                response = c.models.generate_content(
                    model=model,
                    contents=[uploaded, prompt],
                )
                text = response.text or ""
                if not text.strip():
                    raise RuntimeError("Gemini returned an empty response")
                return text
            except Exception as exc:
                last_error = exc
                status = _status_code(exc)
                if not _is_retryable(exc):
                    raise

                # 503 = temporary model overload. Switch immediately after one
                # short retry instead of waiting through repeated failures.
                # Other transient errors get the same bounded retry policy.
                if attempt + 1 < max_attempts:
                    time.sleep(1.5 * (attempt + 1))
                elif model_index + 1 < len(models):
                    break

    raise last_error or RuntimeError("Gemini request failed")


def _upload_with_retry(c, path: Path):
    last_error = None
    for attempt in range(3):
        try:
            return c.files.upload(
                file=str(path),
                config={"mime_type": "application/pdf"},
            )
        except Exception as exc:
            last_error = exc
            if not _is_retryable(exc) or attempt == 2:
                raise
            time.sleep(1.5 * (attempt + 1))
    raise last_error or RuntimeError("Gemini PDF upload failed")


def pdf_to_structured(path: Path, prompt: str):
    c = client()
    uploaded = _upload_with_retry(c, path)
    try:
        return _generate_with_fallback(c, uploaded, prompt)
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
