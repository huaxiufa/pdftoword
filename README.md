# PDF to Word Engine

一个面向高保真 PDF → DOCX 的可扩展转换引擎。

## 当前能力

- PDF 文本层解析
- 字体、字号、粗体、斜体保留
- 基于版面特征的标题识别
- 项目符号列表识别
- PDF 图片提取并嵌入 DOCX
- PDF 页面尺寸、零页边距与近似坐标还原
- 自动把生成的 DOCX 用 LibreOffice 渲染回 PDF
- 原 PDF vs 输出 PDF 自动逐页视觉差异评分
- 自动生成 JSON 对比报告
- FastAPI 上传、下载、报告接口
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

上传 PDF 后自动完成：

```text
PDF → 解析 → 版面分析 → DOCX → LibreOffice PDF → 视觉对比 → JSON 报告
```

示例：

```bash
curl -F "file=@sample.pdf" http://localhost:8000/api/v1/convert
```

返回内容包含：

- `task_id`
- `download_url`：DOCX 下载地址
- `report_url`：视觉对比报告
- `comparison.overall_score`：总体相似度
- `comparison.page_scores`：逐页相似度
- `comparison.issues`：低分页面与页数异常

### GET /api/v1/files/{task_id}

下载转换后的 DOCX。

### GET /api/v1/reports/{task_id}

查看自动生成的视觉差异 JSON 报告。

## 自动优化闭环

当前版本已经建立第一条可量化的优化闭环：

```text
上传 PDF
   ↓
解析 PDF 文本 / 图片 / 坐标
   ↓
版面分析
   ↓
生成可编辑 DOCX
   ↓
LibreOffice 重新排版为 PDF
   ↓
逐页像素差异评分
   ↓
定位低分页面
   ↓
调整布局算法
   ↓
再次转换与评分
```

评分不是“是否生成成功”，而是用于驱动版面算法迭代。当前指标为 100 DPI 下的归一化像素平均绝对误差，并对页数不一致施加惩罚。

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
 ↓
LibreOffice
 ↓
Rendered PDF
 ↓
Visual Comparator
 ↓
comparison.json
```

## 下一阶段：高保真优化

1. 多栏区域自动聚类与正确阅读顺序
2. 表格线 + 文本对齐联合检测，支持合并单元格
3. 页眉/页脚检测与 Word header/footer
4. 图片浮动定位、缩放和文字环绕
5. 扫描 PDF OCR
6. 字体映射与缺失字体回退策略
7. 根据视觉差异报告自动调整字号、缩进、行距、块间距
8. Golden PDF 回归测试集与 CI 阈值
9. 结构差异指标：文本框、图片、表格、页面尺寸，而不仅是像素差异
10. 对低分页面生成差异图，作为算法调参依据
