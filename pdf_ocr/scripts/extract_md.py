#!/usr/bin/env python3
"""Extract first-30-page text from each OCR'd PDF into a same-name .md file
in the same directory. Multiprocess, read-only w.r.t. PDFs.

Progress log compatible with telegram progress_report.py.
"""

import argparse
import json
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

import fitz

PAGES = 30


def worker(args_tuple):
    path, retries = args_tuple
    for attempt in range(retries + 1):
        try:
            doc = fitz.open(path)
            parts = []
            n = min(PAGES, doc.page_count)
            for i in range(n):
                txt = (doc[i].get_text() or "").strip()
                if txt:
                    parts.append(f"## Page {i + 1}\n\n{txt}")
            doc.close()
            if not parts:
                return path, "skip-empty", 0
            md_path = os.path.splitext(path)[0] + ".md"
            tmp = md_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fh:
                fh.write("\n\n---\n\n".join(parts))
            os.replace(tmp, md_path)
            return path, "ok", len(parts)
        except Exception as exc:  # noqa: BLE001
            if attempt >= retries:
                return path, "error", str(exc)
            time.sleep(2 * (attempt + 1))  # backoff (possible file lock)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--pages", type=int, default=PAGES)
    args = ap.parse_args()

    pdfs = []
    for root, _dirs, files in os.walk(args.root):
        for f in files:
            low = f.lower()
            if low.endswith(".pdf") and not f.endswith(
                    (".ocr_tmp.pdf", ".wm_tmp.pdf")):
                pdfs.append(os.path.join(root, f))
    total = len(pdfs)
    print(f"BATCH START: {total} PDFs, engine=extract-md", flush=True)

    ok = skip = err = 0
    t0 = time.time()
    tasks = [(p, 2) for p in pdfs]
    with ProcessPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(worker, t): t[0] for t in tasks}
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
                  f"(pages={info})", flush=True)

    mins = round((time.time() - t0) / 60, 1)
    summary = {"total": total, "md_written": ok, "empty": skip,
               "errors": err, "elapsed_min": mins}
    print("BATCH DONE " + json.dumps(summary, ensure_ascii=False), flush=True)

    msg = (
        "📝 *Markdown 文本提取完成*\n"
        "━━━━━━━━━━━━━━━━━━\n"
        f"📁 处理总数：*{total}* 个 PDF\n"
        f"✅ 已生成 MD：{ok} 个（前 {args.pages} 页）\n"
        f"⏭ 无文本跳过：{skip} 个\n"
        f"❌ 出错：{err} 个\n"
        f"🕐 耗时：{mins} 分钟"
    )
    sys.path.insert(0, r"C:\Users\glenn\.config\opencode\skills\telegram_notifier\scripts")
    from send_telegram import send_message
    send_message(msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
