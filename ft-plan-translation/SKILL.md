---
name: ft-plan-translation
description: "Use this skill whenever the user asks to translate a specific calendar week (CW number) of the flight test planning Excel workbook (FT_Short_Term_Planning_*.xlsm style, Dornierseawings/Seastar program) into bilingual English+Chinese — e.g. '翻译27周', 'translate CW 27 in the flight test plan', '把SN1003的第28周翻译成中英文'. The workbook has MSN1003 and MSN1004 sheets, column A holds CW week numbers, and each week is a 10-row block. The ORIGINAL .xlsm is never modified — the bilingual result is written to a new copy named '<original>_CH_EN.xlsm' (macros preserved). Each English line in a cell gets a Chinese line below it with the SAME colour as the original English line (resolved and baked explicitly, not inherited), same size, regular (non-bold) font; lines with [brackets] or colons are skipped; consecutive 'N FH/Month(s) Inspection' lines of the SAME precision get one merged translation. Activity-cell background/font colours come from a stored static table (references/cf_colors.json), not re-derived per run. Also trigger when the user asks to fix or re-run a previous week translation of this workbook."
---

# Flight Test Plan Week Translation (xlsm, EN → EN+CN, saved to a `_CH_EN` copy)

Translates ONE calendar week of the flight-test planning workbook into bilingual form via direct XML surgery so **VBA macros, images, and data validation are preserved byte-for-byte**.

**Output goes to a NEW file — the original is never touched.** Name the result `<original stem>_CH_EN.xlsm` in the same folder (e.g. `FT Short Term Planning 2026.xlsm` → `FT Short Term Planning 2026_CH_EN.xlsm`). Read from the original every run; write only to the `_CH_EN` copy.

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
   - **Same colour as the ORIGINAL English line**, always baked explicitly. Priority: English run's explicit `<color>` (rgb/indexed/theme) → the resolved cell default font colour → black. The same colour element is also baked into the English runs that lacked one (see [Colour resolution](#colour-resolution--why-baking-is-required)). Chinese is **never** left to inherit a colour.
   - **Same font name** as the English line (typically Arial).
   - **Regular weight (not bold)** — the `<b/>` tag is stripped from the Chinese run.
   - **Font size capped at 9 pt** — if the English line uses a larger size, Chinese is set to 9 pt; if it uses a smaller size, Chinese keeps that smaller size. Chinese lines **never exceed 9 pt**.
3. **Skip (leave untranslated, insert nothing):**
   - lines containing `[` … `]` bracket tags (crew codes like `[MAS/KON]`, `LFTE` assignments)
   - lines containing `:` or `：`
   - empty lines
   - lines already containing Chinese (makes re-runs idempotent)
4. **Inspection merging**: consecutive lines like `25FH Inspection (75%)` / `50FH Inspection (75%)` / `100FH Inspection (75%)` / `200FH Inspection (75%)` get **one** merged Chinese line inserted after the **last** of the group, in the form `25/50/100/200小时定检（75%）` (join the numbers with `/`, keep the common suffix). This applies to both `N FH Inspection` and `N Month(s) Inspection` lines, and **only when the unit type AND the suffix (precision) are identical across the whole run** — e.g. `50FH Inspection` + `25FH Inspection` merge to `50/25FH Inspection`, but `25FH Inspection (75%)` + `50FH Inspection (100%, completed)` are kept separate because their precision differs. The extract script emits the merged key automatically; `apply_translation.py` uses the same rule so the keys match.
5. Only the requested week's cells change. Identical text in other weeks is untouched (the scripts append new shared strings and re-point only the target cells).
6. Both sheets: if the user names a week but not a sheet, do **both** `MSN1003` and `MSN1004` for that week (run the workflow once per sheet on the same file, chaining outputs). If they name a sheet, do just that one.

## Colour resolution — why baking is required

**Never leave a Chinese run to inherit its colour, and never default the English/no-colour case to black.** The rule is: the Chinese line's colour must equal the colour the English line displays **in the original workbook**, baked explicitly into both the English and Chinese runs.

### The rich-text conversion pitfall
Cells in the detail rows are often **plain-string cells** (`<si><t>…</t></si>`) whose text colour comes from the **cell's font** (e.g. `indexed="12"` → blue, `theme="0"` → white-on-green). Once translated, the cell becomes **rich text** (`<si><r>…</r>…</si>`), and the two colour paths diverge:
- English runs that carry **no `<color>`** render **black** in Excel,
- while a Chinese run whose rPr **omits `<color>`** inherits the **cell's font** colour (blue/red/white).

The result is a visible EN/CN mismatch, and a translation that no longer matches the original. Fix: **bake an explicit `<color>` into every run** — the English runs that lacked one (so the output English keeps its original colour) **and** the Chinese runs.

### Colour for a line (priority)
1. **English run's explicit `<color>`** (rgb / indexed / theme) — copy that exact element.
2. Else **the cell's true default font colour**, resolved from `styles.xml`: `cellXf` → `fontId` → `fonts[fontId]` → `<color>`, respecting `applyFont="0"`, resolving `indexed` via the workbook's `indexedColors` palette and `theme` via `theme1.xml`.
3. Else **black** (`FF000000`) — only when no colour is applied anywhere.

### Do NOT trust openpyxl for the cell font colour
`_cell_default_colour_argb()` resolves the cell colour directly from `styles.xml` (via `_resolve_indexed_colors` / `_resolve_theme_colors`), not from `ws[cell].font`. openpyxl's colour resolution has proven unreliable on this workbook (e.g. reporting a red rgb for a cell whose real font is a plain black indexed colour), so the styles.xml path is authoritative. Normalise palette values to 8-hex ARGB (some palette entries are already 8-hex like `00FFFFFF`; prepend `FF` only to 6-hex values, never to 8-hex).

## Activity colours (static table — `references/cf_colors.json`)

The workbook uses conditional formatting (CF) to colour an activity cell by its exact text (e.g. a cell equal to `"Flight Test"` gets a yellow background). After translation the cell reads `"Flight Test\n飞行测试"`, which no longer matches the CF exact-equality check, so the colour would be lost.

Rather than re-parse and re-evaluate the CF rules on every run, the colours are **stored once** in `references/cf_colors.json` and looked up directly by cell text. This is faster, fully deterministic (better prompt-cache hits), and keeps CF data out of the extract JSON (fewer tokens). The stored table for this program:

| Activity text | Background | Font |
|---|---|---|
| Flight Test / Taxi Test / Engine Run Test | `FFFFC000` (amber) | inherit |
| Maintenance / WOT Implementation / Troubleshooting | `FFD071FF` (purple) | inherit |
| Paperwork | `FF00B0F0` (light blue) | inherit |
| Select Activity | theme0 tint (gray) | `FFFF0000` (red) |

Each entry stores a ready-to-use `<patternFill>` (so theme+tint and indexed fills survive exactly) and an optional font ARGB. `apply_translation.py` reads this table and, for each translated cell whose full text matches a key:

- **Bakes the stored fill** into a new `cellXf` style (in `styles.xml`) that clones the cell's original format with the stored fill, and repoints the cell's `s` attribute at it.
- **Applies the stored font colour** so the Chinese run inherits the same colour as the (conditionally-formatted) English text.

**If the workbook's conditional formatting ever changes, regenerate the table** — do not hand-edit it:

```bash
python scripts/dump_cf_colors.py "<workbook>.xlsm" > references/cf_colors.json
```

`dump_cf_colors.py` reads the live CF rules, resolves theme/indexed/tint colours to concrete values, and picks the highest-precedence rule per activity text (lowest OOXML `priority` number wins). This is the ONLY place CF rules are parsed; translation runtime never does.

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

Step 0 — read `/mnt/skills/public/xlsx/SKILL.md` first (general xlsx handling context). Decide the output name up front: `<original stem>_CH_EN.xlsm`. Then:

1. **Extract** the week's translatable lines. The script writes a machine-only context sidecar and prints ONLY the small `to_translate` payload to stdout:
   ```bash
   python scripts/extract_week.py "<original>.xlsm" MSN1003 27 work/ctx_m3_27.json > work/todo_m3_27.json
   ```
   - `work/ctx_m3_27.json` — `{sheet, week, week_row, cells:{ref→si}}`. **Do not read this file** — it exists only for `apply_translation.py`. Keeping it out of context is the main token saving.
   - stdout (`todo_m3_27.json`) — `{"to_translate": {line: ""}}`, the only thing you read/fill.
   - stderr — a one-line summary (cell count, detected `week_row`, line count).

   `week_row` is auto-detected by scanning column A for the matching integer. **Sanity-check** it: load the workbook read-only with openpyxl (`data_only=True`) and confirm `A{week_row}` equals the requested week.

2. **Translate**: fill every `""` value in `to_translate` with the Chinese translation. Use `references/glossary.md` for consistent program terminology (WOT items, inspection names, activity types). Keep acronyms/WOT numbers/percentages untranslated inside the Chinese line (e.g. `FDR加速度参数 - WOT 532（0%，缺少ADAU）`). Save the filled JSON as the mapping file (keep the `{"to_translate": {...}}` shape — a bare `{english: chinese}` dict is also accepted).

3. **Apply** — reads the original, writes the `_CH_EN` copy (original untouched); colours come from `references/cf_colors.json`:
   ```bash
   python scripts/apply_translation.py "<original>.xlsm" work/ctx_m3_27.json work/todo_m3_27.json "<original>_CH_EN.xlsm"
   ```
   This single invocation handles ALL of:
   - Inserting Chinese translations into shared strings (correct font: same colour/name, ≤9 pt, non-bold)
   - Baking an explicit colour into the Chinese run **and** into English runs that lacked one (resolved from the original cell font, see [Colour resolution](#colour-resolution--why-baking-is-required))
   - Baking the stored activity fill into a new cell style
   - Applying the stored activity font colour to the Chinese run
   - Hiding all rows except headers 1–3 and the translated week's 10-row block

   For a second sheet/week, chain: pass the previous output as the input `<src>` (see multi-week below).

4. **Verify** (all of these, every time):
   - **Font colours**: For **every** EN→CN line pair in the translated weeks, compare the Chinese run's colour element with the English run's — they must be byte-identical (both baked). Resolve indexed/theme to concrete ARGB for the report. Confirm no bold on Chinese, size ≤ 9 pt, font name matches (typically Arial). Also cross-check the baked English colour against the ORIGINAL file's colour for that line (explicit run colour, else the original cell's default font colour).
   - **Activity backgrounds**: Spot-check activity cells (e.g. "Flight Test" → amber, "Maintenance"/"WOT Implementation" → purple, "Paperwork" → light blue). Confirm the baked cell style's `fillId` points at the expected colour and that `cellXfs count` increased.
   - **XML validity**: Parse each modified part (sheet XML, `sharedStrings.xml`, `styles.xml`) with `xml.etree.ElementTree.fromstring`.
   - Confirm `xl/vbaProject.bin` in the output is byte-identical to the original's.
   - Confirm the original file on disk is unchanged (size/mtime) and that an untranslated week's cells are untouched.

5. **Post-process for multi-week**: after chaining all weeks, unhide the rows of ALL translated weeks and fix bilingual row heights:
   ```python
   # visible_rows = {1,2,3} | all rows of every translated week block
   # For each visible row, calc height: 15 + (max_lines - 1) × 12 pt
   # Set ht="X" customHeight="1" on <row> elements
   # IMPORTANT: safe attribute manipulation — never regex-replace whole row
   # tags; only touch ht= and customHeight=.
   ```
   Essential because bilingual cells have ~twice the lines, so default heights clip.

6. Present the `_CH_EN` file. Tell the user it is a NEW copy (original preserved), which sheet(s)/week were translated, how many cells changed, how many activity fills were baked, and how many rows were hidden.

### Token / cache-hit notes
- Only `to_translate` (small) enters context; the per-cell inventory lives in the ctx sidecar that only `apply` reads.
- Colours are a stable stored table (`references/cf_colors.json`), not re-derived per run — deterministic inputs improve prompt-cache hits and remove CF data from the extract JSON.
- Keep a `work/` scratch dir for ctx/todo/mapping files; the only artefact that matters is the `_CH_EN.xlsm`.

### Multiple-week translation

Chain per week per sheet, passing each output as the next input. Keep the final name `<original>_CH_EN.xlsm`; use temp names for intermediate steps.

```bash
ORIG="FT Short Term Planning 2026.xlsm"
mkdir -p work

# Week 27, MSN1003  (read original -> step1)
python scripts/extract_week.py "$ORIG" MSN1003 27 work/ctx_m3_27.json > work/todo_m3_27.json
# ... fill work/todo_m3_27.json ...
python scripts/apply_translation.py "$ORIG" work/ctx_m3_27.json work/todo_m3_27.json work/step1.xlsm

# Week 27, MSN1004  (step1 -> step2)
python scripts/extract_week.py work/step1.xlsm MSN1004 27 work/ctx_m4_27.json > work/todo_m4_27.json
# ... fill ...
python scripts/apply_translation.py work/step1.xlsm work/ctx_m4_27.json work/todo_m4_27.json work/step2.xlsm

# Week 28, MSN1003  (step2 -> step3)
python scripts/extract_week.py work/step2.xlsm MSN1003 28 work/ctx_m3_28.json > work/todo_m3_28.json
# ... fill ...
python scripts/apply_translation.py work/step2.xlsm work/ctx_m3_28.json work/todo_m3_28.json work/step3.xlsm

# Week 28, MSN1004  (step3 -> step4)
python scripts/extract_week.py work/step3.xlsm MSN1004 28 work/ctx_m4_28.json > work/todo_m4_28.json
# ... fill ...
python scripts/apply_translation.py work/step3.xlsm work/ctx_m4_28.json work/todo_m4_28.json work/step4.xlsm

# POST-PROCESS (required for multi-week) -> final _CH_EN copy:
#   1. unhide rows of ALL translated weeks   2. adjust bilingual row heights
python scripts/post_process.py work/step4.xlsm "${ORIG%.xlsm}_CH_EN.xlsm"
```

The original `$ORIG` is read but never written. Each apply step modifies only its target cells. The post-processing step is REQUIRED for multi-week to ensure:
- All translated week rows are visible (not just the last week's) — `post_process.py`'s `VISIBLE_ROWS`/`row_height` are keyed to specific week blocks, so update them to match the weeks you translated.
- Row heights accommodate the doubled line count from bilingual content.
- Row height adjustment uses safe attribute manipulation (only touches `ht=`/`customHeight=`, never mangles other attributes like `x14ac:dyDescent`).

## New glossary terms
When you translate a term not yet in `references/glossary.md`, append it there so future weeks stay consistent with past ones.

## Style reference

### Font properties (Chinese lines)
- **Font name**: match English run (usually Arial)
- **Font size**: min(English run size, 9 pt) — never exceed 9 pt
- **Font colour**: **always baked explicitly** — English run's explicit `<color>` if present, else the resolved cell default font colour (styles.xml), else black. Never inherited, never defaulted to black when the original is coloured. See [Colour resolution](#colour-resolution--why-baking-is-required).
- **Font weight**: regular (strip `<b/>`)
- **Other**: `family val="2"` (Swiss/modern sans-serif)

### Font properties (English lines — unchanged)
- Arial 8pt bold is the default for activity-type cells and detail lines
- Black (`000000`) is the default colour; red (`FF0000`) for warnings/executed markers; blue (`0000FF`) for some annotations
- Some colours come from conditional formatting DXF font overrides

### Colour-to-direct-formatting mapping
For each translated cell whose full text matches a key in `references/cf_colors.json`:
1. Stored `fill` (a ready `<patternFill>`) → new cellXf style via `_ensure_fill` + `_ensure_cellxf` (split-on-`<xf`-boundary parsing handles both self-closing and child-bearing xf tags); the cell's `s` attribute is repointed.
2. Stored `font` colour (ARGB or null) → applied to the Chinese run and to English runs that lack an explicit colour (CF colour wins over the cell default).
3. Cell default font colour, resolved from `styles.xml` (via `_cell_default_colour_argb`), is used when the English run has no explicit rPr colour — **not** openpyxl's `cell.font`.
4. Bold → English runs keep `<b/>`; Chinese runs strip it.

The table itself is produced offline by `dump_cf_colors.py`, which does the heavy CF/theme/indexed/tint resolution ONCE. At translation time there is no CF-rule parsing — just a dict lookup.

### Version notes
- **v5 (current)**:
  - **Explicit colour baking for faithful EN→CN colour match.** Colours are no longer inherited: for every line the colour is resolved as English run explicit `<color>` → cell default font colour (resolved from `styles.xml` by `_cell_default_colour_argb`, handling rgb/indexed/theme and `applyFont`) → black, then baked into **both** the English runs that lacked one and the Chinese run. This fixes the plain-string→rich-text divergence (English no-colour runs render black while Chinese no-colour runs inherit the cell font) and stops the earlier "default to black" behaviour from erasing the original blue/red/white cell-font colours.
  - **Precision-aware inspection merging**: only consecutive `N FH/Month(s) Inspection` lines with the SAME unit type AND SAME suffix merge into one Chinese line (e.g. `50/25FH Inspection`), so `25FH Inspection (75%)` + `50FH Inspection (100%, completed)` stay separate.
- **v4**:
  - **Non-destructive output**: the original workbook is never modified; the bilingual result is a new `<original>_CH_EN.xlsm` copy.
  - **Static colour table**: activity backgrounds/fonts are stored in `references/cf_colors.json` (generated by `dump_cf_colors.py`) and looked up by text, instead of re-parsing CF rules each run. Fixes a latent precedence bug (Excel: lowest `priority` number wins — e.g. "Select Activity" is gray+red, not the lower-precedence red fill).
  - **Token/cache**: `extract_week.py` prints only `to_translate`; the per-cell inventory goes to a machine-only ctx sidecar. `apply_translation.py` signature is now `<in> <ctx> <mapping> <out>` (sheet/week come from the ctx).
- **v3**: `_ensure_cellxf` split-on-`<xf`-boundary parsing (self-closing tags); theme/indexed colour resolution (now used only by `dump_cf_colors.py`); row-visibility scope = full 10-row week block; bilingual row-height post-processing (`15 + (max_lines - 1) × 12` pt) using safe attribute manipulation (only `ht=`/`customHeight=`, never whole-tag regex, to avoid mangling `x14ac:dyDescent`).
