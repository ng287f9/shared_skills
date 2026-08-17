---
name: ft-plan-translation
description: "Use this skill whenever the user asks to translate a specific calendar week (CW number) of the flight test planning Excel workbook (FT_Short_Term_Planning_*.xlsm style, Dornierseawings/Seastar program) into bilingual English+Chinese — e.g. '翻译27周', 'translate CW 27 in the flight test plan', '把SN1003的第28周翻译成中英文'. The workbook has MSN1003 and MSN1004 sheets, column A holds CW week numbers, and each week is a 10-row block. Translation is done in place in the original .xlsm (macros preserved): each English line in a cell gets a Chinese line below it, same color/size, regular (non-bold) font; lines with [brackets] or colons are skipped; consecutive 'N FH Inspection' lines get one merged translation. Also trigger when the user asks to fix or re-run a previous week translation of this workbook."
---

# Flight Test Plan Week Translation (xlsm, EN → EN+CN in place)

Translates ONE calendar week of the flight-test planning workbook into bilingual form, editing the original `.xlsm` in place via direct XML surgery so **VBA macros, images, and data validation are preserved byte-for-byte**.

⚠️ **Never round-trip this workbook through openpyxl `load_workbook(...)` + `save()`** — openpyxl drops the wmf images and Data Validation extensions in this file. The provided scripts avoid this by editing `xl/sharedStrings.xml`, sheet XML, and `xl/styles.xml` directly inside the zip. openpyxl is used read-only (font inspection) only.

## Workbook structure (verified)

- Sheets: `MSN1003`, `MSN1004` (others like `Lista` are lookup/support sheets — never touch them).
- Column A of each MSN sheet holds the week number as an integer with number format `"CW" 00` (so cell value `27` displays as "CW 27").
- **Week N's block starts at row `5 + (N-1) × 10`** and spans 10 rows:
  | offset | content |
  |---|---|
  | +0 | CW number (col A) + weekday dates (cols B–H, `ddd - dd/mmm/yyyy`) |
  | +1 | `AM` label + activity type per day (Maintenance / Flight Test / Backup Day …) |
  | +2 | AM detail lines (multi-line, `\r\n`-separated within the cell) |
  | +3 | overflow / Non-Working Day markers |
  | +4 | `PM` label + activity types |
  | +5 | PM detail lines |
  | +6 | overflow / executed markers |
  | +7 | `RMK` + remarks |
  | +8 | overflow |
  | +9 | blank separator |
- Content columns scanned: **B through K**. Column A labels (AM/PM/RMK) and the date row are not translated.
- Cells are largely **rich text**: one cell holds many lines, each line with its own color (black `000000`, red `FF0000`, blue `0000FF`, some `indexed` colors). Line separator inside cells is `\r\n`. Fonts are Arial 8pt **bold** for English, regular (non-bold) for Chinese.

## Translation rules

1. **One English line → one Chinese line directly below it** in the same cell.
2. Chinese line format:
   - **Same colour** as the English line (inherited from the English run's `<rPr>` or from the cell's conditional-formatting font colour).
   - **Same font name** as the English line (typically Arial).
   - **Regular weight (not bold)** — the `<b/>` tag is stripped from the Chinese run.
   - **Font size capped at 9 pt** — if the English line uses a larger size, Chinese is set to 9 pt; if it uses a smaller size, Chinese keeps that smaller size. Chinese lines **never exceed 9 pt**.
3. **Skip (leave untranslated, insert nothing):**
   - lines containing `[` … `]` bracket tags (crew codes like `[MAS/KON]`, `LFTE` assignments)
   - lines containing `:` or `：`
   - empty lines
   - lines already containing Chinese (makes re-runs idempotent)
4. **FH inspection merging**: consecutive lines like `25FH Inspection (75%)` / `50FH Inspection (75%)` / `100FH Inspection (75%)` / `200FH Inspection (75%)` get **one** merged Chinese line inserted after the **last** of the group, in the form `25/50/100/200小时定检（75%）` (join the FH numbers with `/`, keep the common suffix). The extract script emits the merged key (e.g. `25/50/100/200FH Inspection (75%)`) automatically.
5. Only the requested week's cells change. Identical text in other weeks is untouched (the scripts append new shared strings and re-point only the target cells).
6. Both sheets: if the user names a week but not a sheet, do **both** `MSN1003` and `MSN1004` for that week (run the workflow once per sheet on the same file, chaining outputs). If they name a sheet, do just that one.

## Conditional formatting preservation

The workbook has extensive conditional formatting (CF) rules that colour cells based on their text content. For example:

| Activity | Background | Font |
|---|---|---|
| Flight Test | `FFFFC000` (yellow) | theme1, bold |
| Taxi Test | `FFFFC000` (yellow) | bold |
| Engine Run Test | `FFFFC000` (yellow) | theme1, bold |
| Maintenance | `FFD071FF` (purple) | theme1, bold |
| WOT Implementation | `FFD071FF` (purple) | theme1, bold |
| Troubleshooting | `FFD071FF` (purple) | theme1, bold |
| Backup Day | `FF00B0F0` (light blue) | theme1, bold |
| Cancelled / Select Activity | theme0 tint (gray) | `FFFF0000` (red), bold |
| AOG | `FFFF3300` (bright red) | theme0, bold |

These rules use `cellIs` with `operator="equal"` — the formula checks if the cell text **exactly equals** e.g. `"Flight Test"`. After translation the cell contains `"Flight Test\n飞行测试"` which no longer matches the exact-equality check, so the CF is effectively lost.

**The apply script fixes this automatically:**
1. It reads all CF `cellIs equal` rules from the sheet XML before modifying anything.
2. For each cell, it checks whether any CF rule matches the **original** (pre-translation) English text.
3. If a match is found, the script:
   - **Bakes the CF background fill** into the cell by creating a new `cellXf` style (in `styles.xml`) that combines the original cell format with the CF fill colour and points the cell's `s` attribute at the new style.
   - **Bakes the CF font colour** into the rich-text runs so the text colour matches what was originally displayed.
   - The Chinese run inherits this corrected colour, so both EN and CN lines display with the same visual colour.

4. **Font colour resolution** handles all three OOXML colour sources:
   - **Explicit RGB** (`<color rgb="FF0000"/>`) — used directly.
   - **Theme colours** (`<color theme="1"/>`) — resolved via `xl/theme/theme1.xml`'s `<a:clrScheme>` (dk1→1, lt1→0, etc.).
   - **Indexed colours** (`<color indexed="12"/>`) — resolved via `<indexedColors>` palette in `styles.xml`.
   - **CF DXF font colours** go through the same resolution: theme → indexed → none.
   - **Cell default font colours** (via openpyxl) also check indexed and theme in addition to RGB.

## Row hiding

After translation, **all rows except header rows 1–3 and the ENTIRE 10-row block of the target week are hidden** (set `hidden="1"` on `<row>` elements). The full week block includes: CW number row, AM label + detail + overflow, PM label + detail + overflow, RMK + overflow, and blank separator.

Row visibility for translated week:
| Offset | Content | Visible? |
|---|---|---|
| +0 | CW number + dates | **visible** |
| +1 | AM label + activity types | **visible** |
| +2 | AM detail lines | **visible** |
| +3 | AM overflow | **visible** |
| +4 | PM label + activity types | **visible** |
| +5 | PM detail lines | **visible** |
| +6 | PM overflow / executed | **visible** |
| +7 | RMK + remarks | **visible** |
| +8 | RMK overflow | **visible** |
| +9 | blank separator | **visible** |

When translating multiple weeks, each apply invocation only keeps its own week visible. After all translations are applied, a final post-processing pass must unhide the rows of ALL translated weeks (and hide everything else except rows 1–3).

## Workflow

Step 0 — read `/mnt/skills/public/xlsx/SKILL.md` first (general xlsx handling context), then:

1. **Extract** the week's translatable lines:
   ```bash
   python scripts/extract_week.py input.xlsm MSN1003 27 > w27.json
   ```
   Output includes:
   - `to_translate`: deduplicated dict of every line needing translation (skip rules + FH merge already applied)
   - `cells`: per-cell inventory with `si`, `lines`, `col`, `row`, `text` (full raw text for CF matching)
   - `cf_rules`: relevant CF rules found in the sheet
   - `week_row`: the row number of the CW cell (auto-detected; falls back to `5 + (N-1)×10` formula)

   The script auto-detects `week_row` by scanning column A for the matching integer. **Sanity-check** the `week_row` it reports: load the workbook read-only with openpyxl (`data_only=True`) and confirm `A{week_row}`'s value equals the requested week number.

2. **Translate**: fill every `""` value in `to_translate` with the Chinese translation. Use `references/glossary.md` for consistent program terminology (WOT items, inspection names, activity types). Keep acronyms/WOT numbers/percentages untranslated inside the Chinese line (e.g. `FDR加速度参数 - WOT 532（0%，缺少ADAU）`). Save the filled JSON as `mapping.json`.

3. **Apply**:
   ```bash
   python scripts/apply_translation.py input.xlsm MSN1003 27 mapping.json output.xlsm
   ```
   This single invocation handles ALL of:
   - Inserting Chinese translations into shared strings (with correct font properties)
   - Baking CF fills into cell styles
   - Baking CF font colours into text runs
   - Hiding non-AM rows of the translated week

4. **Verify** (all of these, every time):
   - **Font colours**: Spot-check 3+ cells: Chinese runs have `color` matching English runs (including resolution of indexed and theme colours), no bold on Chinese, font size ≤ 9 pt, font name matches (typically Arial).
   - **CF backgrounds**: Spot-check activity cells (e.g. "Flight Test" → yellow bg, "Maintenance" → purple bg). The bilingual cell should display the same background colour as the original conditionally-formatted cell. Verify that the cellXfs count increased (new styles created for baked fills).
   - **XML validity**: Parse each modified sheet XML with a real XML parser (e.g. `xml.etree.ElementTree.fromstring`) to catch any malformed tags or missing attribute spaces.
   - Confirm `xl/vbaProject.bin` in the output zip is byte-identical to the input's.
   - Confirm another (untranslated) week's cells are unchanged.

5. **Post-process for multi-week**: after chaining all weeks, fix row visibility so ALL translated weeks are shown (not just the last one). Also adjust row heights for bilingual content:
   ```python
   # visible_rows = {1,2,3} | all rows of every translated week block
   # For each visible row, calc height: 15 + (max_lines - 1) × 12 pt
   # Set ht="X" customHeight="1" on <row> elements
   # IMPORTANT: use safe attribute manipulation — never regex-replace
   # entire row tags; only target ht= and customHeight= attributes.
   ```
   This step is essential because after translation cells have roughly twice as many lines as before, so default row heights will clip content.

6. Name the output like the input (it replaces it functionally) and present it. Tell the user which sheet(s)/week were translated, how many cells changed, how many CF fills were baked, and how many rows were hidden.

### Multiple-week translation

When translating multiple weeks (e.g. CW27 + CW28), run the workflow in sequence per week per sheet, chaining outputs:

```bash
# Week 27, MSN1003
python scripts/extract_week.py input.xlsm MSN1003 27 > w27_m3.json
# ... fill translations ...
python scripts/apply_translation.py input.xlsm MSN1003 27 w27_m3_mapping.json step1.xlsm

# Week 27, MSN1004 (continues from step1)
python scripts/extract_week.py step1.xlsm MSN1004 27 > w27_m4.json
# ... fill translations ...
python scripts/apply_translation.py step1.xlsm MSN1004 27 w27_m4_mapping.json step2.xlsm

# Week 28, MSN1003
python scripts/extract_week.py step2.xlsm MSN1003 28 > w28_m3.json
# ... fill translations ...
python scripts/apply_translation.py step2.xlsm MSN1003 28 w28_m3_mapping.json step3.xlsm

# Week 28, MSN1004
python scripts/extract_week.py step3.xlsm MSN1004 28 > w28_m4.json
# ... fill translations ...
python scripts/apply_translation.py step3.xlsm MSN1004 28 w28_m4_mapping.json step4.xlsm

# POST-PROCESS (required for multi-week):
# 1. Fix row visibility — unhide rows of ALL translated weeks
# 2. Adjust row heights for bilingual content
python post_process.py step4.xlsm output.xlsm
```

Each apply step modifies only its target cells; previous steps' translations are unaffected. The post-processing step is REQUIRED for multi-week to ensure:
- All translated week rows are visible (not just the last week's)
- Row heights accommodate the doubled line count from bilingual content
- Row height adjustment uses safe attribute manipulation (only touches `ht=` and `customHeight=`, never mangles other attributes like `x14ac:dyDescent`)

## New glossary terms
When you translate a term not yet in `references/glossary.md`, append it there so future weeks stay consistent with past ones.

## Style reference

### Font properties (Chinese lines)
- **Font name**: match English run (usually Arial)
- **Font size**: min(English run size, 9 pt) — never exceed 9 pt
- **Font colour**: match English run; if English run has no explicit colour, use CF-mandated colour or cell default
- **Font weight**: regular (strip `<b/>`)
- **Other**: `family val="2"` (Swiss/modern sans-serif)

### Font properties (English lines — unchanged)
- Arial 8pt bold is the default for activity-type cells and detail lines
- Black (`000000`) is the default colour; red (`FF0000`) for warnings/executed markers; blue (`0000FF`) for some annotations
- Some colours come from conditional formatting DXF font overrides

### CF-to-direct-formatting mapping
The apply script resolves the DXF format to direct cell formatting using this hierarchy:
1. CF fill (background colour) → new cellXf style with matching fill, created by `_ensure_cellxf` which uses split-on-`<xf`-boundary parsing to handle both self-closing and child-bearing xf tags.
2. CF font colour → resolved through theme → indexed → RGB chain, then baked into `<color rgb="..."/>` on rich-text runs.
3. Cell default font colour → resolved via openpyxl Font (RGB → indexed palette → theme file), used when the English run has no explicit rPr.
4. CF bold → English runs already have `<b/>`; Chinese runs strip it.

### Common bugs fixed (v3)
- **`_ensure_cellxf` regex**: the old `<xf>(.*?)</xf>` pattern missed self-closing `<xf …/>` tags, causing xf-count undercount and silently skipping fill baking. Now uses `re.split(r"(?=<xf\b)", …)`.
- **Theme colour resolution**: added `_resolve_theme_colors()` reading `xl/theme/theme1.xml` to map theme indices (0=lt1, 1=dk1, etc.) to RGB.
- **Indexed colour resolution**: added `_resolve_indexed_colors()` reading `<indexedColors>` from `styles.xml` to map palette indices to RGB.
- **Row visibility scope**: changed from hiding-only-3-AM-rows to showing the entire 10-row week block (user feedback: AM/PM/RMK rows were incorrectly hidden).
- **Row height post-processing**: after translation, row heights must be adjusted for bilingual content (formula: `15 + (max_lines - 1) × 12` pt). Must use safe attribute manipulation — only touch `ht=` and `customHeight=`, never regex-replace entire row tags (which can mangle `x14ac:dyDescent` and other attributes).
