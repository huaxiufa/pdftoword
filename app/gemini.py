import os
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
    """Upload a PDF to Gemini and return the model text.

    Retry transient Gemini demand/rate-limit/server errors with exponential
    backoff while reusing the same uploaded file.
    """
    c = client()
    uploaded = c.files.upload(
        file=str(path), config={"mime_type": "application/pdf"}
    )
    try:
        last_error = None
        for attempt in range(4):
            try:
                response = c.models.generate_content(
                    model=MODEL,
                    contents=[uploaded, prompt],
                )
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
