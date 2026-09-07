# PDF → Word · Gemini

这是一次全新重写：不再使用原来的 OCR / DocLayout-YOLO / DrawingML 渲染链路。

新流程：

**PDF → Gemini → HTML → LibreOffice → DOCX**

同时使用 PyMuPDF 提取 PDF 中实际显示的原始图片，再按 Gemini 返回的 `[[PDF_IMAGE_n]]` 占位符回填，因此图片不会依赖 Gemini 重新生成。

## 启动

复制 `.env.example` 为 `.env`，填入：

```env
GEMINI_API_KEY=你的_key
GEMINI_MODEL=gemini-3.7-flash
```

然后：

```bash
docker compose up --build
```

打开：`http://localhost:8000`

## 说明

- 默认上传限制 50 MB。
- Gemini 负责理解文字、表格、页面结构和图片位置。
- 原始 PDF 图片由 PyMuPDF 提取，不让模型重绘图片。
- LibreOffice 在容器内负责最终 HTML → DOCX。
- 如果 Gemini 返回异常，当前版本会直接报错，不会偷偷切回旧 OCR 引擎。

Gemini 官方 API 支持直接上传 PDF 并让模型进行文档理解；本项目使用 Python `google-genai` SDK 的 Files API。
