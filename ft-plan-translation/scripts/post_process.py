import sys, zipfile, re, math

SHEET_PATHS = ['xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml']
SHEET_NAMES = ['MSN1003', 'MSN1004']

# CW27 = rows 265-274, CW28 = rows 275-284, + headers 1-3
VISIBLE_ROWS = {1, 2, 3} | set(range(265, 285))


# ── CJK detection ─────────────────────────────────────────────────────
def _is_cjk(c):
    """Rough CJK + full-width punctuation detection."""
    o = ord(c)
    return (o >= 0x2e80) or (o >= 0xff00 and o <= 0xffef)


# ── column widths from sheet XML ──────────────────────────────────────
def _parse_col_widths(sheet_xml):
    """Return ``{col_number: width_in_excel_units}`` (1-indexed)."""
    cols_m = re.search(r'<cols>(.*?)</cols>', sheet_xml, re.S)
    if not cols_m:
        return {}
    result = {}
    for el in re.finditer(
        r'<col[^>]*?min="(\d+)"[^>]*?max="(\d+)"[^>]*?width="([^"]+)"',
        cols_m.group(1), re.S
    ):
        cmin, cmax = int(el.group(1)), int(el.group(2))
        w = float(el.group(3))
        for c in range(cmin, cmax + 1):
            result[c] = w
    return result


def _col_letter_to_number(col_ref):
    n = 0
    for ch in col_ref.upper():
        n = n * 26 + ord(ch) - 64
    return n


# ── SI text → segments (split by \r\n) ───────────────────────────────
def _build_si_segments(content_map):
    """Build ``{si_index: [segment_text, …]}``."""
    sst_xml = content_map.get('xl/sharedStrings.xml', b'').decode('utf-8')
    si_bodies = re.findall(r'<si>(.*?)</si>', sst_xml, re.S)
    segments = {}
    for i, si in enumerate(si_bodies):
        texts = re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)
        full = ''.join(texts)
        full = full.replace('&lt;', '<').replace('&gt;', '>') \
                   .replace('&amp;', '&').replace('&quot;', '"') \
                   .replace('&apos;', "'")
        full = full.replace('\r\n', '\n').replace('\r', '\n')
        segments[i] = [s for s in full.split('\n') if s] or ['']
    return segments


# ── visual-line count with wrapping simulation ────────────────────────
def _visual_lines_for_si(si_idx, col_width, si_segments):
    """Estimate displayed (wrapped) lines for a cell at given column width.

    Each ``\\r\\n`` segment is evaluated independently:
      - CJK / full-width chars count as 2 weight units
      - all other chars count as 1
      - visual lines per segment = ceil(weighted_length / column_width)
    """
    segs = si_segments.get(si_idx, [''])
    total = 0
    cw = max(col_width, 1)
    for seg in segs:
        weight = sum(2 if _is_cjk(c) else 1 for c in seg)
        total += max(1, math.ceil(weight / cw))
    return max(1, total)


def _max_visual_lines_in_row(part_text, si_segments, col_widths):
    """Maximum visual lines across all shared-string cells in a row."""
    max_l = 1
    for m in re.finditer(
        r'<c r="([A-Z]+)(\d+)"[^>]*\bt="s"[^>]*>.*?<v>(\d+)</v>',
        part_text, re.S
    ):
        col_letter = m.group(1)
        si_idx = int(m.group(3))
        col_num = _col_letter_to_number(col_letter)
        cw = col_widths.get(col_num, 30)
        n = _visual_lines_for_si(si_idx, cw, si_segments)
        if n > max_l:
            max_l = n
    return max_l


# ── main processing ───────────────────────────────────────────────────
def process_sheet_xml(xml_content, si_segments):
    """Re-write row elements with Excel-like auto-fit row heights.

    Algorithm:
      1. Parse column widths from ``<cols>``.
      2. For each visible row, sum visual lines per cell accounting
         for text wrapping at column width (CJK → 2× weight).
      3. Height = ``max(15, visual_lines × 15)`` pt.
    """
    col_widths = _parse_col_widths(xml_content)
    parts = xml_content.split('<row ')
    result = [parts[0]]

    for part in parts[1:]:
        if part.startswith('/>'):
            result.append('<row />')
            continue

        r_match = re.search(r'\br\s*=\s*"(\d+)"', part)
        if not r_match:
            result.append('<row ' + part)
            continue

        row_num = int(r_match.group(1))
        is_visible = row_num in VISIBLE_ROWS

        gt_pos = part.find('>')
        if gt_pos == -1:
            result.append('<row ' + part)
            continue

        attrs = part[:gt_pos]
        if attrs.rstrip().endswith('/'):
            attrs_clean = attrs.rstrip()[:-1].rstrip()
            tag_close = '/>'
            rest = part[gt_pos+1:]
        else:
            attrs_clean = attrs
            tag_close = '>'
            rest = part[gt_pos+1:]

        # Strip stale formatting attributes
        for attr in ('hidden', 'customHeight', 'ht'):
            for q in ('"', "'"):
                pat = rf'\s*{attr}\s*=\s*{q}[^{q}]*{q}'
                attrs_clean = re.sub(pat, '', attrs_clean)

        if is_visible:
            vis = _max_visual_lines_in_row(part, si_segments, col_widths)
            h = round(max(vis * 15, 15), 1)
            if h > 15:
                attrs_clean += f' ht="{h}" customHeight="1"'
        else:
            attrs_clean += ' hidden="1"'

        result.append(f'<row {attrs_clean}{tag_close}{rest}')

    return ''.join(result)


def post_process(input_path, output_path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        file_list = zin.namelist()
        content_map = {}
        for f in file_list:
            content_map[f] = zin.read(f)

    si_segments = _build_si_segments(content_map)
    print(f'  Shared strings parsed: {len(si_segments)} entries')

    for sheet_path, sheet_name in zip(SHEET_PATHS, SHEET_NAMES):
        if sheet_path not in content_map:
            print(f'  WARNING: {sheet_path} not found')
            continue

        xml_content = content_map[sheet_path].decode('utf-8')
        processed = process_sheet_xml(xml_content, si_segments)
        content_map[sheet_path] = processed.encode('utf-8')
        print(f'  {sheet_name}: processed')

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fname in file_list:
            if fname in content_map:
                zout.writestr(fname, content_map[fname])

    print(f'Post-process complete: {output_path}')
    print(f'  Visible: header(1-3) + CW27(265-274) + CW28(275-284)')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: post_process.py <input.xlsm> <output.xlsm>')
        sys.exit(1)
    post_process(sys.argv[1], sys.argv[2])
