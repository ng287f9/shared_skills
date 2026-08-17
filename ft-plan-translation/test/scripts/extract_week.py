#!/usr/bin/env python3
"""Extract the translatable lines of one calendar week from the flight-test
planning workbook (FT_Short_Term_Planning_*.xlsm style).

Usage:
    python extract_week.py <file.xlsm> <sheet: MSN1003|MSN1004> <week_number>

Prints a JSON object:
{
  "sheet": "MSN1003", "week": 27, "week_row": 265,
  "cells": { "B267": {"si": 123, "lines": [...], "col": 2, "row": 267, "text": "..."}, ... },
  "to_translate": { "<english line or merged FH key>": "" , ... },
  "cf_rules": [ ... ]
}

``to_translate`` holds every unique line that NEEDS a Chinese translation
(deduplicated across cells), with skip rules already applied:
  - lines containing [ ... ] bracket tags  -> skipped
  - lines containing ':' or '：'           -> skipped
  - lines already containing CJK           -> skipped (idempotent re-runs)
  - empty lines                            -> skipped
Consecutive "<N>FH Inspection ..." lines are merged into ONE key like
"25/50/100/200FH Inspection (75%)" — translate that single key; the apply
script inserts its translation once, after the LAST FH line of the group.

Claude: fill every "" value with the Chinese translation (regular font is
handled by apply_translation.py; just supply text), save as mapping.json,
then run apply_translation.py.
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
import datetime

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

FH_RE = re.compile(r"^\s*\d+\s*FH Inspection\b(.*)$", re.I)
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


def si_lines(si_el) -> list[str]:
    """Flatten an <si> (plain <t> or rich <r> runs) to text, split to lines."""
    text = "".join(t.text or "" for t in si_el.iter(f"{M}t"))
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def si_text(si_el) -> str:
    """Return the full raw text of an <si> element (no line splitting)."""
    return "".join(t.text or "" for t in si_el.iter(f"{M}t"))


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


def merge_fh(lines: list[str]):
    """Yield (kind, payload): ('line', text) or ('fh', merged_key, count)."""
    i = 0
    while i < len(lines):
        m = FH_RE.match(lines[i])
        if m:
            nums, suffix = [], None
            j = i
            while j < len(lines):
                mm = FH_RE.match(lines[j])
                if not mm:
                    break
                nums.append(re.match(r"\s*(\d+)", lines[j]).group(1))
                suffix = mm.group(1).strip()
                j += 1
            key = "/".join(nums) + "FH Inspection" + (f" {suffix}" if suffix else "")
            yield ("fh", key, j - i)
            i = j
        else:
            yield ("line", lines[i], 1)
            i += 1


def _col_ref(n: int) -> str:
    """Convert column number to letter(s)."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    path, sheet, week = sys.argv[1], sys.argv[2], int(sys.argv[3])
    z = zipfile.ZipFile(path)
    sx = z.read(sheet_xml_name(z, sheet)).decode()
    sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
    sis = list(sst.findall(f"{M}si"))

    # --- Determine week_row: try formula first, then scan column A ---
    week_row = WEEK_ROW0 + (week - 1) * WEEK_STRIDE
    root = ET.fromstring(sx)
    # Verify A{week_row} is the correct CW number
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
        # Scan column A for the matching week number
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
        sys.stderr.write(f"WARNING: Could not find week {week} in column A of {sheet}. "
                         f"Using computed row {week_row}.\n")

    # --- Extract cells ---
    out_cells, to_translate = {}, {}

    for row in root.iter(f"{M}row"):
        r = int(row.get("r"))
        if not (week_row < r <= week_row + BLOCK_ROWS):
            continue
        for c in row.findall(f"{M}c"):
            if c.get("t") != "s":
                continue
            ref = c.get("r")
            col = re.match(r"([A-Z]+)", ref).group(1)
            coln = col_number(col)
            if not (COL_MIN <= coln <= COL_MAX):
                continue
            v = c.find(f"{M}v")
            si_idx = int(v.text)
            lines = si_lines(sis[si_idx])
            full_text = si_text(sis[si_idx])
            out_cells[ref] = {
                "si": si_idx,
                "lines": lines,
                "col": coln,
                "row": r,
                "text": full_text,
            }
            for item in merge_fh(lines):
                if item[0] == "fh":
                    to_translate.setdefault(item[1], "")
                elif not skip_line(item[1]):
                    to_translate.setdefault(item[1].strip(), "")

    # --- Also extract CF rules for the apply script ---
    cf_blocks = re.findall(
        r'<conditionalFormatting\s+sqref="([^"]*)"\s*>(.*?)</conditionalFormatting>',
        sx, re.S)
    cf_rules = []
    for sqref, body in cf_blocks:
        rule_list = []
        for attrs_str, rule_body in re.findall(r"<cfRule(.*?)>(.*?)</cfRule>", body, re.S):
            a = dict(re.findall(r'(\w+)="([^"]*)"', attrs_str))
            if a.get("type") != "cellIs" or a.get("operator") != "equal":
                continue
            fm = re.search(r"<formula>(.*?)</formula>", rule_body)
            if not fm:
                continue
            fm_text = fm.group(1).strip()
            if fm_text.startswith('"') and fm_text.endswith('"'):
                rule_list.append({
                    "text": fm_text[1:-1],
                    "dxfId": int(a.get("dxfId", "0")),
                    "priority": int(a.get("priority", "0")),
                })
        if rule_list:
            cf_rules.append({"sqref": sqref.split(), "rules": rule_list})

    json_str = json.dumps({
        "sheet": sheet, "week": week, "week_row": week_row,
        "cells": out_cells, "to_translate": to_translate,
        "cf_rules": cf_rules,
    }, ensure_ascii=False, indent=1)
    sys.stdout.buffer.write(json_str.encode('utf-8') + b'\n')


if __name__ == "__main__":
    main()
