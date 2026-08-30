#!/usr/bin/env python3
"""Batch-remove image-based Abbott watermarks (multiprocess, in-place).

Targets images whose content digest matches known watermark stamps AND that
are placed in the top strip of a page. Progress log compatible with
telegram progress_report.py.
"""

import argparse
import hashlib
import json
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

import fitz

DIGESTS = {
    "c047e87532cb",  # THIS DOCUMENT PROVIDED BY THE ABBOTT ... (672x136)
    "15a3a97a6f15",  # TECHNICAL LIBRARY crop
    "c1a720cceff7",  # TECHNICAL LIBRARY crop
    "f49c183f8329",  # THIS DOCUMENT PROVIDED BY THE ABBOTT ... variant
    "6fb7f5120c88",  # blue TECHNICAL LIBRARY stamp variant
}

WM_KEYWORDS = ("ABBOTT", "AEROSPACE", "TECHNICAL", "LIBRARY")


def ocr_is_watermark(data: bytes) -> bool:
    """OCR fallback for unknown digests: match Abbott watermark keywords."""
    import io
    import pytesseract
    from PIL import Image
    try:
        im = Image.open(io.BytesIO(data))
        if im.mode not in ("L", "RGB"):
            im = im.convert("RGB")
        txt = pytesseract.image_to_string(im).upper()
        return sum(k in txt for k in WM_KEYWORDS) >= 2
    except Exception:
        return False


def worker(path: str):
    try:
        doc = fitz.open(path)
        removed = 0
        for page in doc:
            xrefs = set()
            for img in page.get_images(full=True):
                xref = img[0]
                try:
                    rects = page.get_image_rects(xref)
                    if not rects:
                        continue
                    r = rects[0]
                    if not (r.y0 < 65 and r.width < 450 and r.height < 60):
                        continue
                    info = doc.extract_image(xref)
                    dg = hashlib.md5(info["image"]).hexdigest()[:12]
                except Exception:
                    continue
                if dg in DIGESTS or ocr_is_watermark(info["image"]):
                    xrefs.add(xref)
            for xref in xrefs:
                try:
                    page.delete_image(xref)
                    removed += 1
                except Exception:
                    pass
        if removed == 0:
            doc.close()
            return path, "skip-nowm", 0
        tmp = path + ".wm_tmp.pdf"
        doc.save(tmp, garbage=4, deflate=True, clean=True)
        doc.close()
        os.replace(tmp, path)
        return path, "ok", removed
    except Exception as exc:  # noqa: BLE001
        return path, "error", str(exc)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--workers", type=int, default=8)
    args = ap.parse_args()

    pdfs = []
    for root, _dirs, files in os.walk(args.root):
        for f in files:
            low = f.lower()
            if low.endswith(".pdf") and not f.endswith(
                    (".ocr_tmp.pdf", ".wm_tmp.pdf")):
                pdfs.append(os.path.join(root, f))
    total = len(pdfs)
    print(f"BATCH START: {total} PDFs, engine=watermark-image-remove",
          flush=True)

    ok = skip = err = 0
    t0 = time.time()
    with ProcessPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(worker, p): p for p in pdfs}
        for i, fut in enumerate(as_completed(futs), 1):
            path, status, info = fut.result()
            if status == "ok":
                ok += 1
            elif status == "error":
                err += 1
                print(f"[{i}/{total}] ERROR: {os.path.basename(path)}: {info}",
                      flush=True)
                continue
            else:
                skip += 1
            print(f"[{i}/{total}] {status}: {os.path.basename(path)} "
                  f"(removed={info})", flush=True)

    mins = round((time.time() - t0) / 60, 1)
    summary = {"total": total, "cleaned": ok, "no_watermark": skip,
               "errors": err, "elapsed_min": mins}
    print("BATCH DONE " + json.dumps(summary, ensure_ascii=False), flush=True)

    msg = (
        "🧹 *图片型水印移除完成*\n"
        "━━━━━━━━━━━━━━━━━━\n"
        f"📁 处理总数：*{total}* 个\n"
        f"✅ 成功去除：{ok} 个\n"
        f"⏭ 无匹配水印：{skip} 个\n"
        f"❌ 出错：{err} 个\n"
        f"🕐 耗时：{mins} 分钟\n\n"
        "💡 剩余未清除的文件可能为烧入扫描背景的水印，无法无损移除"
    )
    sys.path.insert(0, r"C:\Users\glenn\.config\opencode\skills\telegram_notifier\scripts")
    from send_telegram import send_message
    send_message(msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
