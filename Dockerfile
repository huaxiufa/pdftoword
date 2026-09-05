FROM python:3.11-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 \
    PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK=True \
    PADDLE_DEVICE=cpu \
    OCR_LANG=ch \
    DOCLAYOUT_MODEL=juliozhao/DocLayout-YOLO-DocStructBench \
    DOCLAYOUT_CONF=0.20 \
    DOCLAYOUT_IMGSZ=1024 \
    OMP_NUM_THREADS=4 \
    MKL_NUM_THREADS=4
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 libglib2.0-0 libsm6 libxext6 libxrender1 libgomp1 \
    libreoffice-writer fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -U pip && pip install --no-cache-dir -r requirements.txt
COPY app ./app
COPY web ./web
EXPOSE 8000
CMD ["uvicorn","app.main:app","--host","0.0.0.0","--port","8000"]
