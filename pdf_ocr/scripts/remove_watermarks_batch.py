#!/usr/bin/env python3
"""Batch-remove Abbott watermarks from all PDFs (multiprocess, in-place).

Variants handled per remove_watermark.md:
  - content-stream Artifact blocks (/Subtype/Watermark BDC..EMC)
  - per-page /Watermark annotations

Progress lines "[i/total] ok|skip-nowm|error: name" are written to stdout so
the Telegram progress reporter can watch them. Sends a Telegram summary when
finished.
"""

import argparse
import json
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

sys.path.insert(0, r"C:\Users\glenn\.claude\skills\pdf_ocr\scripts")
sys.path.insert(0, r"C:\Users\glenn\.config\opencode\skills\telegram_notifier\scripts")

from remove_watermark import remove  # noqa: E402
from send_telegram import send_message  # noqa: E402


def worker(path: str):
    try:
        removed, out = remove(Path(path), None, "auto", True)
        if out:
            return path, "ok", removed
        return path, "skip-nowm", 0
    except Exception as exc:  # noqa: BLE001
        return path, "error", str(exc)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("root", help="directory to process recursively")
    ap.add_argument("--workers", type=int, default=8)
    args = ap.parse_args()

    pdfs = []
    for root, _dirs, files in os.walk(args.root):
        for f in files:
            if f.lower().endswith(".pdf") and not f.endswith(".ocr_tmp.pdf"):
                pdfs.append(os.path.join(root, f))
    total = len(pdfs)
    print(f"BATCH START: {total} PDFs, engine=watermark-remove", flush=True)

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
    summary = {"total": total, "removed": ok, "no_watermark": skip,
               "errors": err, "elapsed_min": mins}
    print("BATCH DONE " + json.dumps(summary, ensure_ascii=False), flush=True)

    msg = (
        "🧹 *PDF 水印移除完成*\n"
        "━━━━━━━━━━━━━━━━━━\n"
        f"📁 处理总数：*{total}* 个\n"
        f"✅ 成功去除水印：{ok} 个\n"
        f"⏭ 无水印跳过：{skip} 个\n"
        f"❌ 出错：{err} 个\n"
        f"🕐 耗时：{mins} 分钟\n\n"
        "即将重新扫描并恢复双工人 OCR 处理"
    )
    send_message(msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
