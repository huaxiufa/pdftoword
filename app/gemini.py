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
    for value in (
        getattr(exc, "status_code", None),
        getattr(exc, "code", None),
        getattr(getattr(exc, "response", None), "status_code", None),
    ):
        if isinstance(value, int):
            return value
        if isinstance(value, str):
            match = re.search(r"\b(400|401|403|404|408|409|429|500|502|503|504)\b", value)
            if match:
                return int(match.group(1))
    match = re.search(r"\b(400|401|403|404|408|409|429|500|502|503|504)\b", str(exc))
    return int(match.group(1)) if match else None


def _is_retryable(exc: Exception) -> bool:
    return _status_code(exc) in {408, 429, 500, 502, 503, 504}


def _model_candidates():
    """Configured models in priority order. GEMINI_MODELS can override defaults."""
    configured = os.getenv("GEMINI_MODELS", "")
    if configured.strip():
        raw = [x.strip() for x in configured.split(",") if x.strip()]
    else:
        raw = [MODEL, FALLBACK_MODEL]
    result = []
    for value in raw:
        if value and value not in result:
            result.append(value)
    return result


def _generate_with_fallback(c, uploaded, prompt):
    models = _model_candidates()
    if not models:
        raise RuntimeError("No Gemini models configured")

    errors = []
    for model_index, model in enumerate(models):
        for attempt in range(2):
            try:
                print(
                    f"Gemini PDF conversion: model={model} attempt={attempt + 1}/2",
                    flush=True,
                )
                response = c.models.generate_content(
                    model=model,
                    contents=[uploaded, prompt],
                )
                text = response.text or ""
                if not text.strip():
                    raise RuntimeError(f"Gemini model {model} returned an empty response")
                print(f"Gemini PDF conversion succeeded: model={model}", flush=True)
                return text
            except Exception as exc:
                status = _status_code(exc)
                errors.append(f"{model}: {status or type(exc).__name__}: {exc}")
                print(f"Gemini model failed: model={model} status={status} error={exc}", flush=True)

                # 404/400/401/403 are not transient: do not waste time retrying.
                if status in {400, 401, 403, 404}:
                    break
                if not _is_retryable(exc):
                    raise
                if attempt == 0:
                    time.sleep(1.5)

    raise RuntimeError("All configured Gemini models failed. " + " | ".join(errors))


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
