#!/usr/bin/env python3
"""Extract the translatable lines of one calendar week from the flight-test
planning workbook (FT_Short_Term_Planning_*.xlsm style).

Usage:
    python extract_week.py <file.xlsm> <sheet: MSN1003|MSN1004> <week> <ctx_out.json>

Writes TWO things, split for token efficiency:

  * <ctx_out.json>  — MACHINE-ONLY context consumed by apply_translation.py:
        {"sheet","week","week_row","cells": {"B267": 123, ...}}
    (cells maps each translatable cell ref -> its shared-string index).
    Claude never needs to read this file.

  * stdout          — ONLY the ``to_translate`` dict:
        {"to_translate": {"<english line or merged FH key>": "", ...}}
    This is the small payload Claude reads and fills in.  Keeping the heavy
    per-cell inventory out of stdout (and out of Claude's context) is the main
    token/cache-hit optimisation of this skill.

``to_translate`` holds every unique line that NEEDS a Chinese translation
(deduplicated across cells), with skip rules already applied:
  - lines containing [ ... ] bracket tags  -> skipped
  - lines containing ':' or '：'           -> skipped
  - lines already containing CJK           -> skipped (idempotent re-runs)
  - empty lines                            -> skipped
Consecutive "<N>FH Inspection ..." lines are merged into ONE key like
"25/50/100/200FH Inspection (75%)".

Conditional-formatting colours are NOT extracted here — they are read from the
static ``references/cf_colors.json`` table by apply_translation.py.

Claude: fill every "" value in the printed ``to_translate`` with the Chinese
translation, save the result as mapping.json (same shape), then run
apply_translation.py.
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

FH_RE = re.compile(r"^\s*\d+\s*FH Inspection\b(.*)$", re.I)
# Inspection line: "<N> FH Inspection <suffix>" OR "<N> Month(s) Inspection <suffix>".
# group(1)=number, group(2)=unit (FH|Month|Months), group(3)=suffix (qualifier) if any.
INSP_RE = re.compile(r"^\s*(\d+)\s*(FH|Months?)\s*Inspection\b(.*)$", re.I)
CJK_RE = re.compile(r"[一-鿿]")

WEEK_ROW0 = 5      # row of week 1's "CW" cell
WEEK_STRIDE = 10   # rows per week block
BLOCK_ROWS = 9     # rows after the CW row that belong to the week
COL_MIN, COL_MAX = 2, 11  # columns B..K scanned for content


def sheet_xml_name(z: zipfile.ZipFile, sheet: str) -> str:
    wb = z.read("xl/workbook.xml").decode()
    rid = dict(re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb))[sheet]
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    tgt = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))[rid]
    return "xl/" + tgt.lstrip("/")


def si_lines(si_el) -> list:
    """Flatten an <si> (plain <t> or rich <r> runs) to text, split to lines."""
    text = "".join(t.text or "" for t in si_el.iter(f"{M}t"))
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def col_number(col_ref: str) -> int:
    n = 0
    for ch in col_ref:
        n = n * 26 + ord(ch.upper()) - 64
    return n


def skip_line(line: str) -> bool:
    s = line.strip()
    if not s:
        return True
    if "[" in s and "]" in s:
        return True
    if ":" in s or "：" in s:
        return True
    if CJK_RE.search(s):
        return True
    return False


def merge_inspections(lines):
    """Yield (kind, payload): ('line', text) or ('insp', merged_key, count).

    Consecutive lines matching "<N> FH/Month(s) Inspection <suffix>" are merged
    into ONE key only when the unit type AND the suffix (precision) are
    identical across the whole run — e.g. '25/50/100/200FH Inspection (75%)'.
    Lines whose precision differs are NOT merged, so '50FH Inspection' +
    '25FH Inspection' + '20FH Inspection (100%-completed)' yield the separate
    keys '50/25FH Inspection' and '20FH Inspection (100%-completed)'.
    """
    i, n = 0, len(lines)
    while i < n:
        m = INSP_RE.match(lines[i])
        if not m:
            yield ("line", lines[i], 1)
            i += 1
            continue
        unit0 = m.group(2)
        unit0_norm = unit0.lower().rstrip("s")
        nums = [m.group(1)]
        suffix = m.group(3).strip()
        j = i + 1
        while j < n:
            mm = INSP_RE.match(lines[j])
            if not mm:
                break
            if mm.group(2).lower().rstrip("s") != unit0_norm:
                break
            if mm.group(3).strip() != suffix:
                break
            nums.append(mm.group(1))
            j += 1
        key = "/".join(nums) + unit0 + " Inspection" + (f" {suffix}" if suffix else "")
        yield ("insp", key, j - i)
        i = j


def _detect_week_row(root, week):
    """Return the row number of the CW cell for *week* (formula first, then scan)."""
    week_row = WEEK_ROW0 + (week - 1) * WEEK_STRIDE
    a_val = None
    for row in root.iter(f"{M}row"):
        if int(row.get("r")) == week_row:
            for c in row.findall(f"{M}c"):
                if c.get("r") == f"A{week_row}":
                    v = c.find(f"{M}v")
                    if v is not None:
                        try:
                            a_val = int(v.text)
                        except (ValueError, TypeError):
                            pass
    if a_val != week:
        for row in root.iter(f"{M}row"):
            for c in row.findall(f"{M}c"):
                ref = c.get("r")
                if ref and ref.startswith("A"):
                    v = c.find(f"{M}v")
                    if v is not None:
                        try:
                            if int(v.text) == week:
                                week_row = int(row.get("r"))
                                a_val = week
                                break
                        except (ValueError, TypeError):
                            continue
            if a_val == week:
                break
    if a_val != week:
        sys.stderr.write(
            f"WARNING: Could not find week {week} in column A of the sheet. "
            f"Using computed row {week_row}.\n")
    return week_row


def main():
    path, sheet, week = sys.argv[1], sys.argv[2], int(sys.argv[3])
    ctx_out = sys.argv[4] if len(sys.argv) > 4 else f"{sheet}_CW{week}.ctx.json"

    z = zipfile.ZipFile(path)
    sx = z.read(sheet_xml_name(z, sheet)).decode()
    sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
    sis = list(sst.findall(f"{M}si"))
    root = ET.fromstring(sx)

    week_row = _detect_week_row(root, week)

    # --- Extract cells: ref -> shared-string index (slim; apply re-parses text) ---
    cells, to_translate = {}, {}
    for row in root.iter(f"{M}row"):
        r = int(row.get("r"))
        if not (week_row < r <= week_row + BLOCK_ROWS):
            continue
        for c in row.findall(f"{M}c"):
            if c.get("t") != "s":
                continue
            ref = c.get("r")
            coln = col_number(re.match(r"([A-Z]+)", ref).group(1))
            if not (COL_MIN <= coln <= COL_MAX):
                continue
            si_idx = int(c.find(f"{M}v").text)
            cells[ref] = si_idx
            for item in merge_inspections(si_lines(sis[si_idx])):
                if item[0] == "insp":
                    to_translate.setdefault(item[1], "")
                elif not skip_line(item[1]):
                    to_translate.setdefault(item[1].strip(), "")

    # --- Write machine-only context sidecar ---
    ctx = {"sheet": sheet, "week": week, "week_row": week_row, "cells": cells}
    with open(ctx_out, "w", encoding="utf-8") as fh:
        json.dump(ctx, fh, ensure_ascii=False)

    # --- Print ONLY the small to_translate payload for Claude ---
    sys.stderr.write(
        f"context -> {ctx_out}  ({len(cells)} cells, week_row={week_row}); "
        f"{len(to_translate)} lines to translate\n")
    payload = json.dumps({"to_translate": to_translate}, ensure_ascii=False, indent=1)
    sys.stdout.buffer.write(payload.encode("utf-8") + b"\n")


if __name__ == "__main__":
    main()
