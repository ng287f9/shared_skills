"""pdf_ocr: 移除 PDF 水印 (content-stream Artifact 块 / Watermark 注解)

用法:
    python remove_watermark.py <pdf> [--variant auto|content-stream|annotation] [--out 输出路径]

输出: 默认 <pdf名>_nowm.pdf (保留原件), 与 remove_watermark.md 流程一致
"""
import argparse
import re
import sys
from pathlib import Path

import pymupdf

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


def page_has_watermark_annot(doc, page_xref: int) -> bool:
    for xr in page_annot_xrefs(doc, page_xref):
        if doc.xref_get_key(xr, "Subtype")[1].strip() == "/Watermark":
            return True
    return False


def detect_variant(doc) -> str:
    for i in range(doc.page_count):
        for c in doc[i].get_contents():
            if ARTIFACT_PAT.search(doc.xref_stream(c).decode("latin-1")):
                return "content-stream"
    for i in range(doc.page_count):
        if page_has_watermark_annot(doc, doc[i].xref):
            return "annotation"
    return "none"


def remove(pdf_path, out_path=None, variant="auto", inplace=False):
    doc = pymupdf.open(str(pdf_path))
    v = detect_variant(doc) if variant == "auto" else variant
    removed = 0
    if v == "content-stream":
        for i in range(doc.page_count):
            for c in doc[i].get_contents():
                s = doc.xref_stream(c).decode("latin-1")
                s2 = ARTIFACT_PAT.sub("", s)
                if s2 != s:
                    doc.update_stream(c, s2.encode("latin-1"), compress=True)
                    removed += 1
    elif v == "annotation":
        for i in range(doc.page_count):
            if page_has_watermark_annot(doc, doc[i].xref):
                doc.xref_set_key(doc[i].xref, "Annots", "[]")
                removed += 1
    if removed == 0:
        doc.close()
        print(f"[SKIP] {Path(pdf_path).name}: 未检测到可移除水印 (variant={v})")
        return 0, None
    if inplace:
        # 就地覆盖原文件: 先存临时文件再替换 (避免直接覆盖已打开的文档)
        import os

        tmp = str(Path(pdf_path).with_name(Path(pdf_path).stem + "_tmp_nowm.pdf"))
        doc.save(tmp, garbage=4, deflate=True, clean=True)
        doc.close()
        os.replace(tmp, str(pdf_path))
        print(f"[OK] {Path(pdf_path).name} 已就地去除水印 ({v}, 修改 {removed} 处)")
        return removed, str(pdf_path)
    dst = out_path or str(Path(pdf_path).with_name(Path(pdf_path).stem + "_nowm.pdf"))
    doc.save(dst, garbage=4, deflate=True, clean=True)
    doc.close()
    print(f"[OK] {Path(pdf_path).name} -> {dst} ({v}, 修改 {removed} 处)")
    return removed, dst


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="PDF 文件")
    ap.add_argument("--variant", default="auto", choices=["auto", "content-stream", "annotation"])
    ap.add_argument("--out", default=None, help="输出路径 (默认 <pdf名>_nowm.pdf)")
    ap.add_argument("--inplace", action="store_true", help="就地覆盖原文件, 不另存副本")
    args = ap.parse_args()

    try:
        remove(Path(args.pdf), args.out, args.variant, args.inplace)
    except Exception as e:
        print(f"[ERROR] {args.pdf}: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()