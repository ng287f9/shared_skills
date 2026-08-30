#!/usr/bin/env python3
"""Scan a PDF tree: classify files into has-text-layer vs needs-OCR.

Outputs:
  --out LIST_FILE      one pending (needs-OCR) path per line
  --json SUMMARY_JSON  counts + errors
Sends a Telegram summary when finished.
"""

import argparse
import json
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

import fitz

sys.path.insert(0, r"C:\Users\glenn\.config\opencode\skills\telegram_notifier\scripts")
from send_telegram import send_message  # noqa: E402

MIN_TEXT_CHARS = 20


def check_pdf(path: str):
    try:
        doc = fitz.open(path)
        missing = 0
        for p in range(len(doc)):
            if len((doc[p].get_text() or "").strip()) < MIN_TEXT_CHARS:
                missing += 1
        n = len(doc)
        doc.close()
        return path, n, missing, None
    except Exception as exc:  # noqa: BLE001
        return path, 0, 0, str(exc)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--out", required=True)
    ap.add_argument("--json", dest="summary_json", required=True)
    ap.add_argument("--workers", type=int, default=8)
    args = ap.parse_args()

    pdfs = []
    for root, _dirs, files in os.walk(args.root):
        for f in files:
            if f.lower().endswith(".pdf") and not f.endswith(".ocr_tmp.pdf"):
                pdfs.append(os.path.join(root, f))
    total = len(pdfs)
    print(f"SCAN START: {total} PDFs", flush=True)

    pending, has_text, errors = [], 0, []
    pages_missing = pages_total = 0
    t0 = time.time()
    with ProcessPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(check_pdf, p): p for p in pdfs}
        for i, fut in enumerate(as_completed(futs), 1):
            path, npages, missing, err = fut.result()
            if err:
                errors.append({"path": path, "error": err})
            elif missing > 0:
                pending.append(path)
                pages_missing += missing
                pages_total += npages
            else:
                has_text += 1
            if i % 200 == 0 or i == total:
                rate = i / max(time.time() - t0, 1)
                print(f"[{i}/{total}] scanned, pending={len(pending)} "
                      f"({rate:.1f} files/s)", flush=True)

    with open(args.out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(pending))
    summary = {
        "total": total,
        "has_text_layer": has_text,
        "needs_ocr": len(pending),
        "scan_errors": len(errors),
        "pending_pages_total": pages_total,
        "pending_pages_missing": pages_missing,
        "elapsed_min": round((time.time() - t0) / 60, 1),
    }
    with open(args.summary_json, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, ensure_ascii=False, indent=2)
    if errors:
        err_path = args.summary_json + ".errors.json"
        with open(err_path, "w", encoding="utf-8") as fh:
            json.dump(errors, fh, ensure_ascii=False, indent=2)
    print("SCAN DONE " + json.dumps(summary, ensure_ascii=False), flush=True)

    msg = (
        "🔍 *PDF 文字层扫描完成*\n"
        "━━━━━━━━━━━━━━━━━━\n"
        f"📁 扫描总数：*{total}* 个\n"
        f"⏭ 已有文字层：{has_text} 个\n"
        f"🆘 待 OCR 处理：*{len(pending)}* 个\n"
        f"　└ 待处理页数：约 {pages_missing} 页\n"
        f"❌ 损坏/无法打开：{len(errors)} 个\n"
        f"🕐 耗时：{summary['elapsed_min']} 分钟\n\n"
        "即将启动双工人处理：🟢 GPU(Paddle) + 🔵 CPU(Tesseract)"
    )
    send_message(msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
