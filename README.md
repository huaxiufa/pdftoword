# PDF to Word — DocLayout-YOLO + EasyOCR

基于 **DocLayout-YOLO + EasyOCR + python-docx** 的 PDF → 可编辑 Word 转换器。

## Pipeline

`PDF → page image → DocLayout-YOLO → EasyOCR → coordinate JSON → editable DOCX`

- 不使用 `pdf2docx`
- 不依赖 PaddlePaddle / PaddleOCR
- 不做原生 PDF 文本重建
- PDF 中的图片尽量提取为真实 Word 图片，并保留页面坐标
- OCR 文字以可编辑 Word 文本框输出
- 支持中文、英文
- Docker 默认 CPU 单线程，适合 Apple Silicon / M1
- Hugging Face 与 EasyOCR 模型使用 Docker volume 缓存

## Run

```bash
docker compose up --build
```

打开 `http://localhost:8000`。

## Model cache

Docker Compose 会持久化：

- `/root/.cache/huggingface`：DocLayout-YOLO 模型
- `/root/.EasyOCR`：EasyOCR 检测与识别模型

首次启动会下载模型，后续启动直接复用缓存。

## Notes

DocLayout-YOLO 用于识别页面版面区域，EasyOCR 提供文字检测与识别，最终根据坐标生成可编辑 DOCX。复杂表格、艺术字、特殊字体和非常规排版仍可能需要进一步后处理；目标是“版面尽量还原 + 内容可编辑”，而不是承诺像素级复制 PDF。
