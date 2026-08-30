#!/usr/bin/env python3
"""Batch-OCR scanned PDFs into searchable PDFs (text layer added, images kept).

- Skips pages that already have a text layer (< 20 chars => scanned page).
- Writes result to a temp file first, then replaces the original in place.
- Logs progress lines "[i/total] ..." suitable for log-watch summarization.
"""

import argparse
import json
import os
import sys
import time
import traceback

os.environ.setdefault("HIP_VISIBLE_DEVICES", "")
import fitz  # pymupdf

DPI = 200
ZOOM = DPI / 72.0
MIN_TEXT_CHARS = 20


def get_ocr(lang: str):
    from paddleocr import PaddleOCR
    return PaddleOCR(lang=lang, use_doc_orientation_classify=False,
                     use_doc_unwarping=False, use_textline_orientation=False)


def ocr_page_lines(ocr, pix) -> list[tuple[list[float], str]]:
    """Run PaddleOCR on a pixmap; return [(pdf_coords_rect, text), ...]."""
    import numpy as np
    mode = "RGB" if pix.n < 4 else "CMYK"
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = fitz.Pixmap(fitz.csRGB, pix)
        img = np.frombuffer(img.samples, dtype=np.uint8).reshape(
            img.height, img.width, 3)
    res = ocr.predict(img)
    lines = []
    if not res:
        return lines
    r = res[0]
    texts = r.get("rec_texts") or []
    scores = r.get("rec_scores") or [1.0] * len(texts)
    polys = r.get("rec_polys")
    if polys is None:
        polys = r.get("dt_polys")
    for i, text in enumerate(texts):
        if not text.strip():
            continue
        score = float(scores[i]) if i < len(scores) else 0.0
        if score < 0.5:
            continue
        poly = polys[i] if polys is not None and i < len(polys) else None
        if poly is None:
            continue
        pts = [[float(x) / ZOOM, float(y) / ZOOM] for x, y in poly]
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        rect = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
        lines.append((rect, text))
    return lines


def tesseract_page_lines(pix) -> list[tuple[list[float], str]]:
    """Run Tesseract on a pixmap; return [(pdf_coords_rect, text), ...]."""
    import numpy as np
    import pytesseract
    from PIL import Image
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, pix.n)
    if pix.n == 4:
        rgb = fitz.Pixmap(fitz.csRGB, pix)
        img = np.frombuffer(rgb.samples, dtype=np.uint8).reshape(
            rgb.height, rgb.width, 3)
    im = Image.fromarray(img)
    data = pytesseract.image_to_data(im, lang="eng",
                                     output_type=pytesseract.Output.DICT)
    groups: dict[tuple, list] = {}
    n = len(data["text"])
    for i in range(n):
        word = (data["text"][i] or "").strip()
        conf = float(data["conf"][i]) if i < len(data["conf"]) else -1
        if not word or conf < 40:
            continue
        key = (data["page_num"][i], data["block_num"][i],
               data["par_num"][i], data["line_num"][i])
        groups.setdefault(key, []).append(
            (data["left"][i], data["top"][i],
             data["width"][i], data["height"][i], word))
    lines = []
    for _key, words in groups.items():
        x0 = min(w[0] for w in words) / ZOOM
        y0 = min(w[1] for w in words) / ZOOM
        x1 = max(w[0] + w[2] for w in words) / ZOOM
        y1 = max(w[1] + w[3] for w in words) / ZOOM
        lines.append((fitz.Rect(x0, y0, x1, y1),
                      " ".join(w[4] for w in words)))
    return lines


def add_text_layer(page, lines) -> int:
    """Insert invisible text into page. Returns number of lines inserted."""
    count = 0
    for rect, text in lines:
        h = max(rect.height, 2.0)
        fontsize = max(4.0, min(h * 0.85, 24.0))
        origin = fitz.Point(rect.x0, rect.y1 - h * 0.2)
        try:
            page.insert_text(origin, text[:255], fontsize=fontsize,
                             fontname="helv", render_mode=3)
            count += 1
        except Exception:
            continue
    return count


def process_pdf(path: str, ocr, lang: str, engine: str = "paddle") -> dict:
    doc = fitz.open(path)
    pages_ocrd = 0
    words = 0
    changed = False
    for pno in range(len(doc)):
        page = doc[pno]
        existing = (page.get_text() or "").strip()
        if len(existing) >= MIN_TEXT_CHARS:
            continue
        pix = page.get_pixmap(dpi=DPI)
        if engine == "tesseract":
            lines = tesseract_page_lines(pix)
        else:
            lines = ocr_page_lines(ocr, pix)
        n = add_text_layer(page, lines)
        pages_ocrd += 1
        words += n
        changed = True
        print(f"    page {pno + 1}/{len(doc)}: ocr'd, {n} lines", flush=True)

    if not changed:
        doc.close()
        return {"status": "skip-has-text", "pages_ocrd": 0}

    tmp = path + ".ocr_tmp.pdf"
    doc.save(tmp, garbage=3, deflate=True)
    doc.close()
    # sanity check the temp file opens
    chk = fitz.open(tmp)
    ok = len(chk) > 0 and any(
        len(chk[p].get_text()) >= MIN_TEXT_CHARS for p in range(len(chk))
    )
    chk.close()
    if ok:
        os.replace(tmp, path)
        return {"status": "ok", "pages_ocrd": pages_ocrd, "words": words}
    else:
        if os.path.exists(tmp):
            os.remove(tmp)
        return {"status": "verify-failed", "pages_ocrd": 0}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="*", help="PDF files or directories")
    ap.add_argument("--list", dest="list_file",
                    help="file with one PDF path per line")
    ap.add_argument("--lang", default="en", choices=["en", "ch"])
    ap.add_argument("--engine", default="paddle",
                    choices=["paddle", "tesseract"])
    args = ap.parse_args()

    pdfs = []
    if args.list_file:
        with open(args.list_file, "r", encoding="utf-8-sig") as fh:
            pdfs.extend(ln.strip() for ln in fh if ln.strip())
    for p in args.paths:
        if os.path.isdir(p):
            for root, _dirs, files in os.walk(p):
                pdfs.extend(os.path.join(root, f) for f in files
                            if f.lower().endswith(".pdf"))
        elif p.lower().endswith(".pdf"):
            pdfs.append(p)
    total = len(pdfs)
    print(f"BATCH START: {total} PDFs, lang={args.lang}, "
          f"engine={args.engine}", flush=True)

    t0 = time.time()
    ocr = None
    if args.engine == "paddle":
        ocr = get_ocr(args.lang)
    results = {"ok": 0, "skip-has-text": 0, "error": 0}
    for i, pdf in enumerate(sorted(pdfs), 1):
        name = os.path.basename(pdf)
        try:
            r = process_pdf(pdf, ocr, args.lang, args.engine)
            results[r["status"]] = results.get(r["status"], 0) + 1
            print(f"[{i}/{total}] {r['status']}: {name} "
                  f"(pages_ocrd={r['pages_ocrd']})", flush=True)
        except Exception as exc:  # noqa: BLE001
            results["error"] = results.get("error", 0) + 1
            print(f"[{i}/{total}] ERROR: {name}: {exc}", flush=True)
            traceback.print_exc()

    mins = (time.time() - t0) / 60
    summary = {
        "total": total, "elapsed_min": round(mins, 1),
        "searchable_created": results.get("ok", 0),
        "already_had_text": results.get("skip-has-text", 0),
        "errors": results.get("error", 0),
    }
    print("BATCH DONE " + json.dumps(summary, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
