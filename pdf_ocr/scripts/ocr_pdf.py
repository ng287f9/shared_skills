"""pdf_ocr: 页面文本提取 + PaddleOCR 扫描页识别

用法:
    python ocr_pdf.py <pdf|目录> [--pages 30] [--lang ch|en] [--out out.json]

输出 JSON: {"文件路径": {"页码": {"text": "...", "source": "text"|"ocr"}}}
也可作为模块导入: from ocr_pdf import extract_pages
"""
import argparse
import io
import json
import sys
from pathlib import Path

import pymupdf
from PIL import Image

import numpy as np

TEXT_THRESHOLD = 20


def extract_pages(pdf_path, max_pages=30, lang="ch", ocr=None) -> dict:
    """返回 {页码: {"text": str, "source": "text"|"ocr"}}"""
    doc = pymupdf.open(str(pdf_path))
    out = {}
    for i in range(min(max_pages, len(doc))):
        t = doc[i].get_text().strip()
        if len(t) >= TEXT_THRESHOLD:
            out[i] = {"text": t, "source": "text"}
            continue
        if ocr is None:
            from paddleocr import PaddleOCR

            ocr = PaddleOCR(lang=lang)
        pix = doc[i].get_pixmap(dpi=200)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        res = ocr.predict(np.array(img))
        out[i] = {"text": "\n".join(res[0]["rec_texts"]), "source": "ocr"}
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("target", help="PDF 文件或目录")
    ap.add_argument("--pages", type=int, default=30)
    ap.add_argument("--lang", default="ch")
    ap.add_argument("--out", default=None, help="JSON 输出路径 (默认 stdout)")
    args = ap.parse_args()

    target = Path(args.target)
    files = sorted(target.glob("*.pdf")) if target.is_dir() else [target]
    all_res = {}
    for f in files:
        try:
            all_res[str(f)] = extract_pages(f, args.pages, args.lang)
        except Exception as e:
            print(f"[ERROR] {f.name}: {e}", file=sys.stderr)
    text = json.dumps(all_res, ensure_ascii=False, indent=1)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
    else:
        sys.stdout.write(text)


if __name__ == "__main__":
    main()