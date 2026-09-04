import os
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


def pdf_to_structured(path: Path, prompt: str):
    uploaded = client().files.upload(file=str(path), config={"mime_type": "application/pdf"})
    try:
        response = client().models.generate_content(model=MODEL, contents=[uploaded, prompt])
        return response.text or ""
    finally:
        try:
            client().files.delete(name=uploaded.name)
        except Exception:
            pass
