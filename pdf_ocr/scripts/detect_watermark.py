"""pdf_ocr: 检测 PDF 水印载体类型与内容

用法:
    python detect_watermark.py <pdf|目录> [--lang en|ch]

输出(每文件): 载体类型 content-stream / annotation / none, 受影响页数, 顶部 OCR 文字
"""
import argparse
import io
import re
import sys
from pathlib import Path

import pymupdf
from PIL import Image

ARTIFACT_PAT = re.compile(r"/Artifact\s*<</Subtype\s*/Watermark[^>]*>>\s*BDC.*?EMC", re.S)
ANNOT_XREF_RE = re.compile(r"(\d+) 0 R")


def page_annot_xrefs(doc, page_xref: int) -> list[int]:
    """返回页面注解 xref 列表 (解析 /Annots 间接引用)"""
    key = doc.xref_get_key(page_xref, "Annots")
    if key[0] == "null":
        return []
    refs = key[1]
    if key[0] == "xref":
        m = re.match(r"(\d+) 0 R", refs)
        if not m:
            return []
        refs = doc.xref_object(int(m.group(1)), compressed=True)
    return [int(m.group(1)) for m in ANNOT_XREF_RE.finditer(refs)]


def detect(pdf_path: Path, lang: str = "en") -> dict:
    doc = pymupdf.open(str(pdf_path))
    pages = doc.page_count
    cs_pages, annot_pages, sample_rect = 0, 0, None
    for i in range(pages):
        page = doc[i]
        for c in page.get_contents():
            if ARTIFACT_PAT.search(doc.xref_stream(c).decode("latin-1")):
                cs_pages += 1
                break
        key = doc.xref_get_key(page.xref, "Annots")
        for xr in page_annot_xrefs(doc, page.xref):
            if doc.xref_get_key(xr, "Subtype")[1].strip() == "/Watermark":
                annot_pages += 1
                if sample_rect is None:
                    sample_rect = doc.xref_get_key(xr, "Rect")[1]
                break

    if cs_pages:
        variant = "content-stream"
    elif annot_pages:
        variant = "annotation"
    else:
        variant = "none"

    top_text = []
    if variant != "none":
        from paddleocr import PaddleOCR

        import numpy as np

        ocr = PaddleOCR(lang=lang)
        clip = pymupdf.Rect(0, 0, doc[0].rect.width, 100)
        pix = doc[0].get_pixmap(clip=clip, dpi=150)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        top_text = ocr.predict(np.array(img))[0]["rec_texts"]
    doc.close()
    return {
        "file": pdf_path.name,
        "pages": pages,
        "variant": variant,
        "content_stream_pages": cs_pages,
        "annotation_pages": annot_pages,
        "sample_rect": sample_rect,
        "top_ocr_text": top_text,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("target", help="PDF 文件或目录")
    ap.add_argument("--lang", default="en", help="顶部 OCR 语言 (默认 en)")
    args = ap.parse_args()

    target = Path(args.target)
    files = sorted(target.glob("*.pdf")) if target.is_dir() else [target]
    for f in files:
        try:
            r = detect(f, args.lang)
            print(f"[{r['variant']}] {r['file']} (pages={r['pages']}, "
                  f"content-stream={r['content_stream_pages']}, annotation={r['annotation_pages']}, "
                  f"rect={r['sample_rect']})")
            if r["top_ocr_text"]:
                print(f"    top OCR: {' | '.join(r['top_ocr_text'])}")
        except Exception as e:
            print(f"[ERROR] {f.name}: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()