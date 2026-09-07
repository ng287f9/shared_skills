# Short Term Flight Test Plan — Native Bilingual Table Spec

Since CW36 the Short Term Flight Test Plan slides (the last 2 slides of the
source deck) are **rebuilt as native bilingual tables** (previously left
untouched). Screenshots are EMF pastes of the "FT Short Term Planning" Excel
week blocks; OCR them and rebuild per this spec. Working example:
`scripts/build_plan_tables.py` (CW36-based, adapt data per week).

## Extraction pipeline (Windows)

1. **Map pictures**: slide rels → `image*.emf`. Each plan slide has 2 pictures:
   a short wide **banner strip** (~12.5in x 0.98in, legend, mostly already
   bilingual) and the **table screenshot** (~12.5in x 3.5-4.0in).
2. **EMF → PNG at 4x** via GDI+ (`scripts/emf_to_png.ps1`). EMFs are vector;
   4x renders crisp 8pt text for OCR/visual reading.
3. **Grid + OCR**: detect row/col lines (dark-pixel profile; verticals may be
   light-grey — relax threshold). Per-cell tesseract `--psm 6` + dominant-color
   sampling for fill AND font color. **Italic red/blue lines OCR badly** —
   verify them with visual reading of 4x crops. Cross-check dates/arithmetic.
   (The pasted range may differ from the live Excel workbook — trust the IMAGE,
   not the workbook, unless the user says otherwise.)
4. Tables are **6 columns**: label col (CW/AM/PM/RMK) + Mon..Fri (weekends not
   pasted). Day headers use German abbreviations `Mo/Di/Mi/Do/Fr - DD.Mon.YYYY`.

## Table layout (user-specified, CW36)

- **Font: Arial, 8 pt** (sz=800) for EN and CN. latin=Arial; ea can stay
  SimSun-ExtB for CN glyph rendering.
- **EN bold, CN regular** (non-bold). Italic EN lines (e.g. battery maint,
  flap actuator) → CN regular upright.
- **CN line color = EN line color** (red FF0000 / blue 0000FF / black / white
  on status bars). Never inherit — bake explicit solidFill on every run.
- **Table width = banner image width** (12.50in slide SN1003 / 12.44in SN1004),
  same x=0.37in; table starts just below the banner (y≈1.90in).
- **First column narrow** (0.5-0.63in), **remaining 5 columns equal width**.
- Borders: outer 12700 / inner 6350 black; cell margins small
  (marL/R=27432, marT/B=9000); lnSpc 95%, spc 0; anchor ctr.
- Label cells (CW n / AM / PM / RMK): grey `C0C0C0`, black, bold, centered.

## Palette (from source screenshots)

| Element | Fill | Font |
|---|---|---|
| Date row (CW n + Mo-Fr dates) | `C0C0C0` | black bold |
| "Today" date cell (report date, e.g. Fr of CW36) | `000000` (black) | white bold |
| Maintenance / WOT Implementation activity row | `D071FF` purple | black bold |
| Flight Test / Engine Run Test activity row | `FFC000` orange | black bold |
| Select Activity | `BFBFBF` grey | `FF0000` red bold |
| Detail cells | `FFFFFF` | per-line red/blue/black |
| Status bar "…Executed" | `00B050` green | white bold |
| Status bar cancelled (empty red bar) | `FF0000` | — |
| RMK row | `FFFFFF` | red or black |

## Structure & row order (mirror the image's band order exactly)

SN1003-type week: date row / AM activity row / AM detail (3+ sub-lines) /
PM activity row / PM detail / RMK.
SN1004-type week: date row / AM activity row / AM detail / executed-bar row /
PM activity row / PM detail / executed-bar row / RMK.
Label col: `AM` on the activity row (merge label cells vertically across the
AM block if the image shows it centered), `PM` likewise, `RMK` single.
Executed/cancelled bars are **separate rows under the detail row** at the
day's column (green bar text: `Flight Test Executed 飞行测试已执行` /
`Engine Run Executed 发动机试车已执行`; cancelled = empty red bar + RMK note
`Flight is cancelled for EASA Activity 因EASA活动取消飞行`).

## Translation rules (from ft-plan-translation)

1. One EN line → one CN line directly below, same color, CN regular ≤ EN size.
2. **Skip (EN-only, no CN inserted):** lines with `[brackets]` (crew codes
   `[MAS/KON]`, `[09:00]` timestamps stay untranslated as EN), lines containing
   `:` or `：` (`LFTE:TRK`), empty lines.
3. **Inspection merging**: consecutive `25FH/100FH/200FH Inspection (75%)`
   lines (same unit + same precision suffix) → ONE merged CN line
   `25/100/200小时定检（75%）` after the last line of the group.
4. Keep acronyms/WOT numbers/percentages Latin inside CN lines
   (`FDR加速度参数 - WOT 532（80%，缺少ADAU）`). Use full-width （） in CN.
5. Terminology: use `references/glossary.md` § Short Term Flight Test Plan.

## Multi-week split rule

If one screenshot contains **two week blocks** (e.g. CW36+CW37 stacked in one
image), **split them onto two slides** — the week's slide keeps banner + its
table; the extra week becomes a NEW slide appended after it (clone the plan
slide XML + rels minus notesSlide, register slide in Content_Types /
presentation rels / sldIdLst — or reuse the clone function in
`scripts/build_plan_tables.py`), title suffix `(CWnn)`. Slide count becomes
original + extra weeks.

## Other conventions

- Banner strip images stay as-is (legend is already bilingual). Do NOT delete
  them: they are wide-but-short (~0.98in); the table image is the wide-and-TALL
  (≥2in) one — remove by HEIGHT, not width.
- Plan-slide callouts (e.g. `Rectangle 7`) get one appended CN paragraph,
  sz=1000, grey `5F6F82`, centered (same as slide-4 Rectangle 4).
- Row heights: set explicit minimums per estimated line count (≈190000 emu per
  2-line row, +110000 per extra line); PowerPoint grows rows that need more.
  Keep the whole table above the footer bar (compress lnSpc to 95% / margins if
  tight).
