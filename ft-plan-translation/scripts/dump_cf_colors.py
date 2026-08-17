#!/usr/bin/env python3
"""One-time generator for ``references/cf_colors.json``.

Reads a real FT_Short_Term_Planning workbook, resolves every ``cellIs equal``
conditional-formatting rule on the MSN sheets down to concrete ARGB colours
(theme and indexed colours are looked up in ``theme1.xml`` / the indexed
palette), and writes a flat ``text -> {bg, font}`` table.

Run this ONLY when the workbook's conditional formatting changes.  At runtime
``apply_translation.py`` just reads the resulting JSON — it never re-parses CF
rules, which keeps each translation deterministic (better prompt-cache hits)
and cheap (no CF data in the extract JSON).

Usage:
    python dump_cf_colors.py <workbook.xlsm> [sheet ...] > ../references/cf_colors.json
    # default sheets: MSN1003 MSN1004
"""
import json
import re
import sys
import zipfile

# reuse the (already-tested) resolution helpers from the apply script
from apply_translation import (
    _sheet_xml_name, _parse_dxfs, _parse_cf_rules,
    _resolve_theme_colors, _resolve_indexed_colors, _build_fill_direct,
)


def _resolve_font_color(dxf, theme_colors, indexed_colors):
    """DXF font colour -> 'AARRGGBB' string (or None)."""
    if dxf["font_color_rgb"]:
        return dxf["font_color_rgb"]
    if dxf["font_theme"] is not None:
        rgb = theme_colors.get(int(dxf["font_theme"]))
        if rgb:
            return "FF" + rgb if len(rgb) == 6 else rgb
    if dxf["font_color_indexed"] is not None:
        rgb = indexed_colors.get(int(dxf["font_color_indexed"]))
        if rgb:
            return "FF" + rgb if len(rgb) == 6 else rgb
    return None


def _resolve_bg_fill(dxf):
    """DXF fill -> ready-to-use <patternFill> XML for a cell fill (or None).

    Returns the fill XML directly (not just an RGB) so theme+tint and indexed
    backgrounds — e.g. the gray 'Select Activity' fill — are preserved exactly.
    """
    if not dxf["fill_xml"]:
        return None
    fill_xml, _rgb = _build_fill_direct(dxf["fill_xml"])
    return fill_xml


def main():
    path = sys.argv[1]
    sheets = sys.argv[2:] or ["MSN1003", "MSN1004"]

    z = zipfile.ZipFile(path)
    styles_xml = z.read("xl/styles.xml").decode()
    dxfs = _parse_dxfs(styles_xml)
    theme_colors = _resolve_theme_colors(z)
    indexed_colors = _resolve_indexed_colors(styles_xml)

    # text -> (priority, {fill, font}); keep the highest-PRECEDENCE rule per
    # text.  Per OOXML, the LOWEST priority number wins (1 = highest precedence).
    best = {}
    for sheet in sheets:
        sheet_xml = z.read(_sheet_xml_name(z, sheet)).decode()
        for _sqref_parts, rule_list in _parse_cf_rules(sheet_xml):
            for rule in rule_list:
                text = rule["text"]
                dxf_id = rule["dxfId"]
                if dxf_id >= len(dxfs):
                    continue
                dxf = dxfs[dxf_id]
                entry = {
                    "fill": _resolve_bg_fill(dxf),
                    "font": _resolve_font_color(dxf, theme_colors, indexed_colors),
                }
                if entry["fill"] is None and entry["font"] is None:
                    continue
                prev = best.get(text)
                if prev is None or rule["priority"] < prev[0]:
                    best[text] = (rule["priority"], entry)

    table = {text: entry for text, (_pri, entry) in sorted(best.items())}
    out = {
        "_comment": ("Static activity-text -> conditional-format colour table, "
                     "baked from the workbook's CF rules by dump_cf_colors.py. "
                     "'fill' is a ready-to-use <patternFill> (preserves theme/tint "
                     "and indexed fills); 'font' is ARGB (8 hex) or null. Regenerate "
                     "only when the workbook's conditional formatting changes."),
        "map": table,
    }
    sys.stdout.buffer.write(
        json.dumps(out, ensure_ascii=False, indent=1).encode("utf-8") + b"\n")


if __name__ == "__main__":
    main()
