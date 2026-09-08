# Image Compressor

一个直接可用的本地 Web 图片压缩器。支持批量压缩 JPG、PNG、WebP，调整质量、缩放尺寸、转换格式，并把所有结果一次打包成 ZIP 下载。

## 功能

- 批量上传 JPG / JPEG / PNG / WebP
- 自动压缩并生成 ZIP
- 输出格式：自动、WebP、JPG、PNG
- 质量滑块：40–95
- 最大宽度 / 最大高度缩放
- 可设置目标文件大小（JPG / WebP 会自动降低质量）
- 自动处理 EXIF 旋转
- 默认删除 EXIF 等元数据
- PNG / WebP 保留透明背景
- 自动避免同名输出文件覆盖
- 显示原始大小、压缩后大小和节省比例
- 原始本地文件不会被修改

2026 年的常见 Web 图片工作流通常会优先考虑 WebP；AVIF 压缩率更高但编码更慢、兼容性略低，因此这个版本先把 WebP 作为默认现代格式。citeturn0search1turn0search7

## 运行

```bash
docker compose up -d --build
```

打开：

```text
http://localhost:8000
```

## 更新

```bash
git pull origin main
docker compose up -d --build
```

## 推荐参数

### 商品图 / 网站图片

- 输出：WebP
- 质量：75–85
- 最大宽度：1920 或 1600
- 删除 EXIF：开启

### 手机照片

- 输出：WebP
- 质量：75–82
- 最大宽度：1600–1920
- 如果必须兼容旧系统：改成 JPG

### 截图 / 透明 PNG

- 输出：PNG 或 WebP
- 如果必须保留透明背景，不要转 JPG

## Docker

默认最大上传总大小为 50 MB，可通过环境变量调整：

```yaml
services:
  app:
    environment:
      - MAX_UPLOAD_MB=100
```

## 商业化方向

这个版本可以直接继续做成一个付费微工具：

- 免费：每天少量图片 / 每次最多 10 张
- Pro：批量 100 张、最大文件限制更高
- 后续可增加图片尺寸检测、自动生成 WebP/AVIF 双版本、网站图片 SEO 优化等功能
