---
name: doc_summary
description: Use whenever the user asks to summarize PDF documents (批量PDF摘要、PDF 摘要生成、提取PDF摘要、summary the PDFs、summarize a document、生成中文摘要). The skill: (1) optionally delegates OCR and watermark removal to the pdf_ocr skill when the PDFs contain scanned pages or the user asks to remove watermarks, (2) extracts text from the first 30 pages of each PDF with pymupdf, (3) generates Chinese summaries locally with Ollama qwen2.5:7b via HTTP API. Environment-specific pitfalls are baked in: HTTP proxy must be bypassed for localhost (NO_PROXY / curl --noproxy), and Ollama CLI is broken by the proxy so use the REST API directly.
---

# doc_summary — PDF 批量摘要 (pymupdf + pdf_ocr + qwen2.5:7b)

## 环境 (已部署, 见 D:\pdf-summary-ai\README.md)

| 组件 | 路径 / 版本 |
|------|-------------|
| 项目根目录 | `D:\pdf-summary-ai` |
| Python | `D:\pdf-summary-ai\.venv\Scripts\python.exe` (3.11.9) |
| 配置 | `D:\pdf-summary-ai\config.yaml` |
| Ollama | 0.32.14, API `http://localhost:11434` |
| LLM | `qwen2.5:7b` (Q4, 4.36GB, 全 GPU ~86 tok/s) |

OCR 与去水印环境（PaddleOCR 等）见 **pdf_ocr** skill。

## 环境陷阱 (必须遵守)

1. **代理**: 系统 HTTP 代理 `127.0.0.1:10809` 会拦截 localhost 请求。
   - Ollama CLI 报 `something went wrong` 时: 先设 `$env:NO_PROXY="localhost,127.0.0.1,::1"`
   - 调 Ollama API: 一律用 `curl.exe --noproxy "*"` 或 Python `requests` + `trust_env=False`
   - Python 脚本内: `session = requests.Session(); session.trust_env = False`

## 工作流

### 1. 输入
- PDF 放入 `D:\pdf-summary-ai\input\` 或由用户指定目录
- 配置文件 `config.yaml` 控制: 读取前 30 页 (`pdf.read_first_pages`)、批量大小、模型名、温度
- **不读全文**: 只读取前 30 页内容; 文档不足 30 页时读取整个文档

### 2. 调用 pdf_ocr skill (OCR / 去水印)
处理 PDF 文件时先加载 **pdf_ocr** skill，按其工作流执行：
- **去水印**: 若用户要求或检测到水印 → 按 pdf_ocr 流程 (输入水印文字 → 测试 1 个 PDF → 用户确认 → 询问是否批量)。
- **OCR**: 含扫描页的 PDF → 用 pdf_ocr 的 `scripts/ocr_pdf.py` 提取 (文本层 + PaddleOCR)。
- 摘要脚本 `scripts/summarize_pdf.py` 会在检测到扫描页时自动调用 pdf_ocr 的 OCR 脚本，无需手动干预。

### 3. 文本提取 (pymupdf)
- **只读前 30 页** (`min(30, 页数)`): 不足 30 页才读整个文档
- 每页用 `pymupdf` (import as `pymupdf`, 不用已废弃的 `fitz`) 提取文本层
- 判断扫描页: 文本长度 < 阈值 (如 20 字符) → 交给 pdf_ocr 做 OCR

### 4. 摘要生成 (Ollama qwen2.5:7b)
- 通过 HTTP API, 不用 CLI: `POST http://localhost:11434/api/generate`
- `requests` 会话必须 `trust_env = False` (绕过代理)
- 模型名: `qwen2.5:7b` (8GB 显存够用; 显存不足时回退 `qwen2.5:3b`)
- 参数建议: `stream=False`, `options: {"temperature": 0.3, "num_ctx": 4096}`
- **两阶段摘要** (长文档): 分块 ~3000 字符 → 阶段1逐块提取要点(保留术语/数值/要求) → 阶段2用下方"航空技术文档专家模板"对全部要点做一次 ~1000 字结构化中文摘要。短文档(单块)直接走阶段2。
- 航空技术文档专家摘要模板 (默认格式, 见下方提示词): 输出 8 节结构 (基本信息 / 背景目的 / 核心内容 / 关键参数 / 要求与注意 / 技术价值 / 关键词 / 一句话总结), 中文约 1000 字。

```
你是资深航空技术文档分析专家，具有航空工程、适航法规、技术标准和科研论文分析经验。
阅读以下英文技术文档内容，生成高质量中文摘要。
文档类型可能包括: 航空技术报告、科研论文、适航法规、适航标准、咨询通告(AC)、技术规范、工程手册、专业书籍章节。

【重要限制】
1. 只能根据提供的文本内容进行总结。
2. 不得使用外部知识补充。
3. 不得猜测作者没有说明的内容。
4. 不得改变技术含义。
5. 对法规、标准类文件，必须保留原文中的强制性要求。
6. 如果信息不存在，请明确说明"文档未提供相关信息"。

【摘要目标】生成约1000字中文专业摘要。摘要不是简单翻译，而是理解英文技术内容、提取核心技术信息、保留工程意义、转换为中文专业表达。

【输出结构】
# 1. 文档基本信息
文件名称、发布机构/作者(如果存在)、发布时间(如果存在)、文档类型、适用领域
# 2. 文档背景与目的
为什么发布该文件、解决什么问题、应用场景
# 3. 核心内容摘要
技术原理、方法流程、系统组成、关键技术、主要章节内容
- 论文: 研究问题、方法、实验设计、数据来源、主要结果
- 技术报告: 技术目标、工程方案、测试方法、验证过程、结果
- 法规/咨询通告: 适用对象、管理范围、主要要求、限制条件、合规措施、实施要求
# 4. 关键技术参数和数据
数值、性能指标、时间要求、重量、温度、速度、可靠性指标、测试结果。没有数据不要创造。
# 5. 关键要求与注意事项
必须满足的条件、禁止事项、风险点、工程注意事项、合规要求
# 6. 技术价值和应用意义
对工程人员、研发/测试/认证工作、实际应用的价值与意义
# 7. 专业关键词 (10个以内)
英文关键词 + 中文对应
# 8. 一句话总结
这份文件主要解决什么问题

【写作要求】使用航空工程专业中文；保留英文缩写(FAA、EASA、AC、FAR、CS、DO-178C、DO-254等)且首次出现时给出中文解释；不删除关键技术细节；不使用营销语言；不写"本文非常重要"等空泛句子；保持客观。
【输出语言】中文，约1000字。

文档内容:
{notes}
```

### 5. 输出
- 默认在 PDF 所在目录生成与 PDF 同名的 `<pdf名>.md` (UTF-8); 用 `--out-dir` 可指定输出目录
- 每篇摘要包含: 文件名、页数、OCR 页数、模型、耗时、摘要正文

## 快捷命令

```powershell
# 环境验证 (GPU)
D:\pdf-summary-ai\.venv\Scripts\python.exe -c "import paddle; print(paddle.device.is_compiled_with_cuda(), paddle.device.cuda.device_count())"

# Ollama API 验证
curl.exe --noproxy "*" http://localhost:11434/api/tags
```

## 参考脚本

`scripts/summarize_pdf.py` — 完整流水线参考实现 (提取→OCR(pdf_ocr)→分块→摘要→输出)。
用法: `.venv\Scripts\python.exe <skill>\scripts\summarize_pdf.py [pdf路径或input目录] [--out-dir 目录] [--ocr-script PATH]`
- 自动定位同级 pdf_ocr 的 `scripts/ocr_pdf.py` (也可用 `--ocr-script` 覆盖路径)
- OCR 相关脚本与去水印流程见 **pdf_ocr** skill