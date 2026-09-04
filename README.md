# PDF to Word Engine

一个面向高保真 PDF → DOCX 的可扩展转换引擎。

## v0.2.0 — Optimization Framework

本版本把转换流程升级为可量化的闭环优化系统：

```text
PDF
 ↓
Parser
 ↓
Document Model
 ↓
Layout Engine
 ↓
DOCX Generator
 ↓
LibreOffice
 ↓
Rendered PDF
 ↓
Visual Comparator
 ├─ visual score
 ├─ page score
 ├─ overlay PNG
 ├─ diff PNG
 └─ debug PDF
 ↓
Optimization Engine
 ↓
最佳参数 / 最佳 DOCX
```

### 当前能力

- PDF 文本层解析
- 字体、字号、粗体、斜体保留
- 标题与项目符号识别
- PDF 图片提取并嵌入 DOCX
- 页面尺寸、零页边距与近似坐标还原
- East Asian 字体 OOXML 命名空间安全处理
- LibreOffice 渲染回 PDF
- 逐页视觉差异评分
- 自动搜索 DOCX 字号/垂直密度参数
- 自动停止：达到高分或连续多轮收敛
- 自动保存优化参数历史
- 自动生成 `debug.pdf`
- 自动生成逐页 `overlay-page-N.png`
- 自动生成逐页 `diff-page-N.png`
- JSON 对比与优化报告
- FastAPI 上传、下载、报告和诊断资源接口
- Docker / Docker Compose 一键运行

## 本地运行

需要 Docker Desktop。

```bash
docker compose down
docker compose build --no-cache
docker compose up
```

打开：

```text
http://localhost:8000/docs
```

## API

### POST /api/v1/convert

上传 PDF 后自动完成完整优化闭环：

```text
PDF → 解析 → 版面分析 → 候选 DOCX → LibreOffice → 视觉评分 → 参数搜索 → 最佳 DOCX
```

返回内容包含：

- `task_id`
- `download_url`
- `report_url`
- `debug_url`
- `comparison.overall_score`
- `comparison.page_scores`
- `comparison.issues`
- `comparison.optimization`

### GET /api/v1/reports/{task_id}

获取最终 `comparison.json`。

### GET /api/v1/reports/{task_id}/debug.pdf

获取原 PDF 与最终渲染结果的叠加诊断 PDF。

### GET /api/v1/reports/{task_id}/diagnostics/overlay-page-N.png

查看指定页面的叠加图。

### GET /api/v1/reports/{task_id}/diagnostics/diff-page-N.png

查看指定页面的差异图。

## 优化历史

每次转换都会在任务目录保存：

```text
data/reports/<task_id>/
├── comparison.json
├── optimization-history.json
├── debug.pdf
├── candidates/
│   ├── candidate-00.docx
│   └── ...
├── iteration-01/
├── iteration-02/
└── final/
    ├── debug.pdf
    └── diagnostics/
        ├── overlay-page-1.png
        └── diff-page-1.png
```

优化器优先选择页数完全一致的候选，再比较视觉分数；达到 0.985 或连续三轮变化小于 0.001 时提前停止，避免无意义地继续渲染。

## 版本路线

- **v0.2.0**：Optimization Framework
- **v0.3.0**：更强诊断、结构差异指标、Golden PDF 回归测试
- **v0.4.0**：更智能的参数搜索与多栏/表格/页眉页脚优化
- **v1.0.0**：稳定生产版本
