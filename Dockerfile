# syntax=docker/dockerfile:1.7
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    DOCLAYOUT_MODEL=juliozhao/DocLayout-YOLO-DocStructBench \
    DOCLAYOUT_MODEL_FILE=doclayout_yolo_docstructbench_imgsz1024.pt \
    DOCLAYOUT_CONF=0.20 \
    DOCLAYOUT_IMGSZ=1024 \
    OCR_LANG=ch_sim,en \
    OCR_CPU_THREADS=1 \
    OMP_NUM_THREADS=1 \
    MKL_NUM_THREADS=1 \
    TOKENIZERS_PARALLELISM=false

RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 libglib2.0-0 libsm6 libxext6 libxrender1 libgomp1 \
    libreoffice-writer fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt ./
RUN --mount=type=cache,target=/root/.cache/pip \
    pip install -U pip && \
    pip install -r requirements.txt

COPY app ./app
COPY web ./web

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=3)" || exit 1

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
