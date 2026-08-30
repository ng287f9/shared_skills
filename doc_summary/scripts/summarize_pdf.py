"""doc_summary 参考实现: PDF -> 文本/OCR(pdf_ocr skill) -> qwen2.5:7b 中文摘要

用法:
    python summarize_pdf.py <pdf路径或目录> [--pages 30] [--model qwen2.5:7b] [--out-dir DIR] [--ocr-script PATH]

OCR 步骤委托给 pdf_ocr skill 的 scripts/ocr_pdf.py (默认自动定位同级 pdf_ocr 目录,
可用 --ocr-script 覆盖)。检测到扫描页(<20字符)时自动调用。
输出: 默认在 PDF 所在目录生成与 PDF 同名的 <pdf名>.md (可用 --out-dir 指定目录)
依赖: D:\\pdf-summary-ai\\.venv (pymupdf, requests)
"""
import argparse
import json
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import pymupdf
import requests

OLLAMA_URL = "http://localhost:11434/api/generate"
TEXT_THRESHOLD = 20
DEFAULT_OCR_SCRIPT = str(Path(__file__).resolve().parent.parent.parent / "pdf_ocr" / "scripts" / "ocr_pdf.py")


def get_session() -> requests.Session:
    s = requests.Session()
    s.trust_env = False  # 绕过系统代理, 直接访问 localhost
    return s


def extract_text(pdf_path: Path, max_pages: int) -> tuple[list[str], list[int]]:
    """返回 (每页文本列表, 需要OCR的页码列表)"""
    doc = pymupdf.open(str(pdf_path))
    pages_text, ocr_pages = [], []
    for i in range(min(max_pages, len(doc))):
        text = doc[i].get_text().strip()
        pages_text.append(text)
        if len(text) < TEXT_THRESHOLD:
            ocr_pages.append(i)
    return pages_text, ocr_pages


def run_ocr_script(pdf_path: Path, max_pages: int, ocr_script: str, lang: str = "ch") -> dict:
    """调用 pdf_ocr skill 的 ocr_pdf.py, 返回 {页码: {"text":..., "source": "text"|"ocr"}}"""
    tmp_dir = Path(tempfile.mkdtemp(prefix="ocr_"))
    tmp = tmp_dir / "ocr.json"
    try:
        subprocess.run(
            [sys.executable, ocr_script, str(pdf_path), "--pages", str(max_pages),
             "--lang", lang, "--out", str(tmp)],
            check=True, capture_output=True,
        )
        data = json.loads(tmp.read_text(encoding="utf-8"))
    finally:
        tmp.unlink(missing_ok=True)
        try:
            tmp_dir.rmdir()
        except OSError:
            pass
    return data.get(str(pdf_path), {})


def chunk_text(text: str, size: int = 3000) -> list[str]:
    return [text[i : i + size] for i in range(0, len(text), size)] if text.strip() else []


def _call_llm(session: requests.Session, prompt: str, model: str) -> str:
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": 0.3, "num_ctx": 4096},
    }
    r = session.post(OLLAMA_URL, json=payload, timeout=1800)
    r.raise_for_status()
    return r.json()["response"].strip()


def summarize_chunk(session: requests.Session, chunk: str, model: str) -> str:
    """阶段1: 逐块提取要点, 保留术语/数值/要求, 不丢失细节"""
    prompt = (
        "你是专业的航空技术文档分析师。请用中文提取以下英文文档片段的要点，"
        "保留所有关键术语、英文缩写、数值、参数、时间要求和强制性要求。"
        "只提取，不总结，不遗漏，不加评述。\n\n"
        f"文档内容:\n{chunk}"
    )
    return _call_llm(session, prompt, model)


AVIATION_EXPERT_TEMPLATE = """你是资深航空技术文档分析专家，具有航空工程、适航法规、技术标准和科研论文分析经验。
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
{notes}"""


def final_summary(session: requests.Session, notes: str, model: str) -> str:
    """阶段2: 用航空技术文档专家模板生成 ~1000 字结构化中文摘要"""
    if len(notes) > 8000:
        notes = notes[:8000] + "\n[...内容过长, 已截断...]"
    return _call_llm(session, AVIATION_EXPERT_TEMPLATE.format(notes=notes), model)


def process_one(pdf_path: Path, session: requests.Session, args) -> Path:
    t0 = time.time()
    pages_text, ocr_pages = extract_text(pdf_path, args.pages)
    ocr_text = {}
    if ocr_pages:
        ocr_text = run_ocr_script(pdf_path, args.pages, args.ocr_script)

    full_text = []
    for i, t in enumerate(pages_text):
        entry = ocr_text.get(str(i))
        if entry and entry["source"] == "ocr":
            full_text.append(f"[OCR页{i+1}]\n{entry['text']}")
        elif t:
            full_text.append(f"[文本页{i+1}]\n{t}")
    full = "\n".join(full_text)

    chunks = chunk_text(full)
    if not chunks:
        summary = "未能提取到文本内容。"
    else:
        notes = [summarize_chunk(session, c, args.model) for c in chunks]
        summary = final_summary(session, "\n\n".join(notes), args.model)

    out_dir = Path(args.out_dir) if args.out_dir else pdf_path.parent
    out = out_dir / f"{pdf_path.stem}.md"
    out.write_text(
        f"# {pdf_path.name}\n\n"
        f"- 处理页数: {min(args.pages, len(pages_text))}\n"
        f"- OCR 页数: {len(ocr_pages)}\n"
        f"- 模型: {args.model}\n"
        f"- 耗时: {time.time()-t0:.1f}s\n\n"
        f"## 摘要\n\n{summary}\n",
        encoding="utf-8",
    )
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("target", help="PDF 文件或目录")
    ap.add_argument("--pages", type=int, default=30)
    ap.add_argument("--model", default="qwen2.5:7b")
    ap.add_argument("--out-dir", default=None, help="摘要输出目录 (默认: PDF 所在目录)")
    ap.add_argument("--ocr-script", default=DEFAULT_OCR_SCRIPT, help="pdf_ocr 的 ocr_pdf.py 路径")
    args = ap.parse_args()

    if not Path(args.ocr_script).exists():
        print(f"警告: 未找到 OCR 脚本 {args.ocr_script} (扫描页将跳过 OCR)", file=sys.stderr)

    session = get_session()
    try:
        r = session.get("http://localhost:11434/api/tags", timeout=10)
        r.raise_for_status()
    except Exception as e:
        print(f"错误: Ollama 不可用 ({e})。请先启动服务。")
        sys.exit(1)

    target = Path(args.target)
    files = sorted(target.glob("*.pdf")) if target.is_dir() else [target]
    for f in files:
        try:
            out = process_one(f, session, args)
            print(f"[OK] {f.name} -> {out}")
        except Exception as e:
            print(f"[FAIL] {f.name}: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()