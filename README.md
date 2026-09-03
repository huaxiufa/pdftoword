# PDF to Word Engine

一个面向高保真 PDF → DOCX 的可扩展转换引擎第一版。

## 当前能力

- PDF 文本层解析
- 字体、字号、粗体、斜体保留
- 基于版面特征的标题识别
- 项目符号列表识别
- PDF 图片提取并嵌入 DOCX
- 多页 PDF
- FastAPI 上传与下载接口
- Docker / Docker Compose 一键运行

## 本地运行

需要 Docker Desktop。

```bash
docker compose up --build
```

打开：

```text
http://localhost:8000/docs
```

## API

### POST /api/v1/convert

上传 PDF：

```bash
curl -F "file=@sample.pdf" http://localhost:8000/api/v1/convert
```

返回 `task_id` 和下载地址。

### GET /api/v1/files/{task_id}

下载转换后的 DOCX。

## 架构

```text
PDF
 ↓
PyMuPDF Parser
 ↓
Layout Analyzer
 ↓
Document IR
 ↓
DOCX Renderer
 ↓
DOCX
```

## 下一阶段

1. 多栏阅读顺序识别
2. 表格检测与合并单元格
3. 页眉/页脚识别
4. 浮动图片位置还原
5. 扫描 PDF OCR
6. 更精确的 Word 分页和段落间距
7. 回归测试与 PDF/Word 视觉差异评估
