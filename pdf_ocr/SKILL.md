---
name: pdf_ocr
description: Use whenever the user asks to OCR PDF files or remove watermarks from PDFs (PDF OCR、扫描件识别、扫描页提取、去水印、水印去除、remove watermark、watermark removal、水印检测). Detects and removes PDF watermarks (two known variants: content-stream /Artifact blocks and per-page /Watermark annotations) with PaddleOCR verification, and extracts text from scanned pages with PaddleOCR. When asked to remove watermarks, follow the confirmation workflow: ask the user for the watermark text, test on ONE pdf first, show the result, get user confirmation, then ask whether to batch-process the rest. The doc_summary skill delegates its OCR and watermark steps to this skill.
---

# pdf_ocr — PDF OCR 与去水印 (PaddleOCR)

## 环境 (已部署, 见 D:\pdf-summary-ai\README.md)

| 组件 | 路径 / 版本 |
|------|-------------|
| Python | `D:\pdf-summary-ai\.venv\Scripts\python.exe` (3.11.9) |
| PaddleOCR | 3.7.0 + paddlepaddle-gpu 3.3.1 (cu126, CUDNN 9.9) |
| 模型缓存 | `C:\Users\glenn\.paddlex\official_models\` (首次自动下载) |

## 环境陷阱 (必须遵守)

1. **代理**: 系统 HTTP 代理 `127.0.0.1:10809` 会拦截 localhost。PaddleOCR 离线可用（模型已缓存）；若需调 Ollama 用 `curl.exe --noproxy "*"` 或 `requests` + `trust_env=False`。
2. **CUDNN**: 勿升级 `nvidia-cudnn-cu12`（必须 ==9.9.0.52）。
3. **首次调用 PaddleOCR**: 自动下载模型，仅首次。
4. **GPU 降级**: `use_gpu=False` 可强制 CPU；默认 GPU 不可用时自动降级。
5. **qwen 不可用于图像**: 本地 Ollama `qwen2.5:7b` 为纯文本模型，不支持图像输入（报 `Multimodal data provided, but model does not support multimodal requests`）。水印/图像的 OCR 确认只能用 PaddleOCR 或 tesseract。

## 工作流

### A. 去水印 (可选，必须与用户确认)

1. **询问用户是否去水印**。是 → 继续；否 → 直接做 OCR 提取 (B)。
2. **提示用户输入水印文字内容**（例如 `THIS DOCUMENT PROVIDED BY THE ABBOTT AEROSPACE TECHNICAL LIBRARY ABBOTTAEROSPACE.COM`）。
3. **选一个 PDF 做测试**，运行 `scripts/detect_watermark.py <pdf>`：
   - 检测水印载体类型：
     - `content-stream`：页面内容流中的 `/Artifact <</Subtype/Watermark>> BDC ... EMC` 块（NASA 系列，见 F:\abbotspace\downloads\remove_watermark.md）
     - `annotation`：每页一个 `/Watermark` 注解（顶部 rect，AIAA 系列）
     - `none`：无这两种载体
   - 脚本会用 PaddleOCR 识别第 1 页顶部条带，**把识别文字与用户输入比对**，确认确实是该水印。
4. **移除测试 PDF**：`scripts/remove_watermark.py <pdf>` → 输出 `<pdf名>_nowm.pdf`（保留原件）。
5. **验证**：渲染 `_nowm.pdf` 顶部条带 + PaddleOCR，水印文字消失、正文和文本层完好。把验证结果展示给用户。
6. **用户确认后**，再询问是否批量处理目录内其余 PDF；批量时逐个跳过已处理文件、`none` 文件、损坏文件（`format error: non-page object in page tree`）。
7. 注意：若水印烧入扫描背景图（无内容流块、无注解），脚本无法移除，需告知用户只能图像修复。

### B. OCR 文本提取 (供 doc_summary 调用)

- **只读前 30 页** (`--pages 30` 默认): 文档不足 30 页时读取整个文档。
- 文本层页直接用 pymupdf 提取；扫描页（文本 < 20 字符）用 PaddleOCR（dpi=200）。
- 运行 `scripts/ocr_pdf.py <pdf|目录> [--pages 30] [--lang ch|en] [--out out.json]`。
- 输出 JSON：`{"文件路径": {"页码": {"text": "...", "source": "text"|"ocr"}}}`。

## 脚本

| 脚本 | 用途 |
|------|------|
| `scripts/detect_watermark.py` | 检测水印载体类型 + 顶部 OCR 确认水印文字 |
| `scripts/remove_watermark.py` | 移除水印（content-stream / annotation 两种变体），默认输出 `_nowm.pdf`；`--inplace` 可就地覆盖原文件（不另存副本） |
| `scripts/ocr_pdf.py` | 页面文本提取 + PaddleOCR 扫描页识别，输出 JSON |

## 快捷命令

```powershell
# 环境验证 (GPU)
D:\pdf-summary-ai\.venv\Scripts\python.exe -c "import paddle; print(paddle.device.is_compiled_with_cuda(), paddle.device.cuda.device_count())"

# OCR 冒烟测试
D:\pdf-summary-ai\.venv\Scripts\python.exe -c "from paddleocr import PaddleOCR; o=PaddleOCR(lang='ch'); print(o.predict(r'D:\pdf-summary-ai\temp\ocr_test.png')[0]['rec_texts'])"

# 水印检测
D:\pdf-summary-ai\.venv\Scripts\python.exe <skill>\scripts\detect_watermark.py <pdf>
```