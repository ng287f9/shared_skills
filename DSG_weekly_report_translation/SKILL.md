---
name: DSG_weekly_report_translation
description: "Use this skill whenever the user asks to translate, localize, or produce a bilingual (Chinese/English) version of a weekly status report, program status report, or flight test report PowerPoint — especially aerospace/aircraft program reports (Dornier Seawings Seastar, SN1002/1003/1004, CW-numbered weekly reports). Trigger on phrases like '周报', '双语', '中英文版', '中英对照', 'weekly report translation', 'bilingual pptx', or when the user uploads an English-only .pptx weekly report and asks for a Chinese+English version. Converts every English-only slide into bilingual EN+Chinese: cover (EN-only), TOC inline CN, Highlights/Lowlights & Status bullet CN below EN, Flight Test Program Progress screenshot tables rebuilt as native bilingual tables (approved formatting baked into references/template_slide5_spec.md — the template reference slide is NO LONGER inserted since CW36), and Short Term Flight Test Plan EMF screenshots rebuilt as native bilingual Arial-8pt tables with per-line color matching, translation skip/merge rules, and multi-week screenshots split onto one slide per week (references/short_term_plan_tables.md). Ships with helper scripts and a domain glossary (incl. ft-plan terminology) for consistent output."
---

# Skill: DSG_weekly_report_translation
# Weekly Report Translation (EN → EN/CN bilingual)

Converts English-only slides in a weekly status report `.pptx` into bilingual English+Chinese, following the exact formatting conventions used in this program's reports. Chinese is always added as a supplement to the English — never replaces it.

**Scope (since CW36):**
- **Handled:** Cover page (kept English-only), Table of Contents (inline CN), Weekly Highlights / Lowlights (bullet CN below), Status SN#### (bullet CN below), **Flight Test Program Progress** (screenshot tables → native bilingual tables; formatting from `references/template_slide5_spec.md`), **Short Term Flight Test Plan** (EMF screenshot tables → native bilingual tables; multi-week images split one slide per week; spec in `references/short_term_plan_tables.md`).
- **Discontinued:** inserting the template reference slide after the data slide (removed as of CW36). Its approved formatting (widths, fonts, fills, borders) is saved in `references/template_slide5_spec.md` instead.

**Read `scripts/USAGE.md` first** for the concrete per-phase workflow (Windows: PowerShell + PowerPoint COM), then the two spec files for table formatting.

## Reference files and scripts in this skill

- `references/template_slide5_spec.md` — **Approved Flight Test Program Progress table formatting** (4 tables, column widths, row heights, Arial 6pt, border weights, full palette). Source of truth for Phase 2; replaces copying template slide 5 into outputs.
- `references/short_term_plan_tables.md` — **Short Term Flight Test Plan table spec**: EMF→PNG→OCR extraction pipeline, Arial 8pt / EN-bold-CN-regular, per-line CN color matching, palette, row order, translation skip/merge rules, multi-week split rule.
- `references/glossary.md` — EN↔CN glossary, including § Short Term Flight Test Plan (merged from ft-plan-translation).
- `references/formatting_patterns.md` — exact XML patterns for the text-slide bilingual layouts.
- `references/template.pptx` — original approved bilingual template deck (backup reference only; no longer copied into outputs).
- `scripts/USAGE.md` — practical per-phase commands (setup, OCR, QA, Windows gotchas). READ FIRST.
- `scripts/translate_text_slides.py` — base pattern for Phase 1 (TOC / Highlights / Status). Each week gets a CW-specific copy with hardcoded translations.
- `scripts/rebuild_flight_test_progress.py` — base pattern for Phase 2 (Flight Test Program Progress → 4 native bilingual tables).
- `scripts/build_plan_tables.py` — Phase 3 (Short Term Plan tables; CW36-based working example including slide-cloning for the multi-week split).
- `scripts/emf_to_png.ps1` — EMF → 4x PNG (GDI+) for OCR of plan screenshots.
- `scripts/cell_ocr.py`, `scripts/grid_detect.py`, `scripts/ascii_view.py` — per-cell OCR / grid-line detection / ASCII sanity views.
- `scripts/render_slides.py` — LibreOffice renderer (for containers); on Windows use PowerPoint COM export (USAGE §5).
- `scripts/insert_template_slide.py` — **DEPRECATED** (former template-reference insertion). Do not run; kept only as a reference for slide-cloning XML surgery.

## Workflow (4 phases)

### Phase 1 — Text slides: translate and format

Create a CW-week-specific script (e.g. `translate_text_slides_cw36.py`) from the base pattern — read the source slide XMLs, identify paragraph indices, hardcode the week's translations.

1. Unpack source. **Slide 2 (TOC):** inline CN per line, sz=1400 (per-line override allowed — gotcha below).
2. **Slide 3 (Highlights/Lowlights):** headers inline CN (`Highlights亮点` / `Lowlights不足`, sz=1600, no color, inserted between the EN run and the trailing-tab run); each bullet gets a CN paragraph below (sz=1400, grey `808080`, `buNone`, marL = EN marL+indent); **remove trailing empty paragraphs** and **tighten aggressively** (lnSpc→100%, spcBef/spcAft→0) so content clears the footer bar.
3. **Slide 4 (Status):** bullets CN below (indices relative to the post-empty-removal paragraph list); callout Rectangle 4 gets one appended centered CN paragraph (sz=1000, `5F6F82`) and, when the added CN lines would push bullets underneath it, **move the box below the text block** (~y=4.85in).
4. Repack.

**Gotchas (hit & fixed during CW36):**
- ⚠️ **endParaRPr ordering** — never append a run after `<a:endParaRPr>`; PowerPoint silently drops it (symptom: inline CN missing in render). Insert with `endParaRPr.addprevious(run)`.
- ⚠️ **TOC CN size** — at sz=1600 the CN on the longest TOC line runs behind the team photo (left edge x=5.90in). Default sz=1400; drop a still-clipping line to sz=1200.
- Footer "Program risks are outlined…" stays English-only; slide-4 section labels (`SN1002:` etc.) stay English-only.

### Phase 2 — Flight Test Program Progress: native bilingual tables

CW-week rebuild script from the base pattern: OCR the source slide-5 screenshots (grid detect + per-cell OCR + fill-color sampling; validate arithmetic), then remove all `<p:pic>` from slide 5 and rebuild the 4 native tables (Weekly Overview, Flight Hours & Flights, Progress Summary, Weekly Block/Flight Time):

| Attribute | Source of truth |
|---|---|
| Table structure, column widths, merged cells | `references/template_slide5_spec.md` |
| Font typeface/size (Arial 6pt), border weights (12700 outer / 6350 inner) | `references/template_slide5_spec.md` |
| Cell content / data values, font color, background fill | Source English screenshot (OCR + color sampling) |
| Row height | Auto-fit to the row's actual bilingual content |

"Flight Test Status" banner gets inline CN ` 飞行测试状态` (sz=1400, `002060`, inserted before endParaRPr). Remove the orphan highlight "Oval 2" if present after replacing screenshots.

### Phase 3 — Short Term Flight Test Plan: native bilingual tables

Per `references/short_term_plan_tables.md`, adapting `scripts/build_plan_tables.py` each week:

1. Each plan slide has 2 pictures: the **banner strip** (wide, ~0.98in tall — KEEP as image) and the **table screenshot** (wide, ≥2in tall — REPLACE). Identify by HEIGHT, not width.
2. Convert EMF→4x PNG (`emf_to_png.ps1`), grid-detect + per-cell OCR + color sampling; **visually verify italic red/blue lines** (OCR misreads them); cross-check dates.
3. Rebuild each week block as a 6-column native table: narrow label column (0.5–0.63in) + 5 equal day columns; **table width = banner image width**, x=0.37in; **Arial 8pt** (sz=800); EN bold / CN regular; **CN line color = EN line color**; palette & row order per spec; `[bracket]` crew/timestamp lines and `LFTE:xxx` colon lines stay EN-only; consecutive same-precision inspection lines get ONE merged CN line (e.g. `25/100/200小时定检（75%）`); executed/cancelled status bars as separate green `00B050` / red `FF0000` rows with white EN+CN text.
4. **Multi-week split**: one screenshot holding two CW blocks (e.g. CW36+CW37) → two slides. The original slide keeps week 1 (title suffix `(CW36)`); a cloned slide appended at the end holds week 2 (title suffix `(CW37)`). Clone via XML surgery (slideN.xml + rels minus notesSlide, `[Content_Types].xml` Override `/ppt/slides/slideN.xml`, presentation.xml.rels new rId, sldIdLst append with fully-qualified `r:id` namespace) — the clone function in `build_plan_tables.py` is ready-made.
5. Plan-slide callouts (Rectangle 7 etc.) get one appended CN paragraph (sz=1000, `5F6F82`, centered).

### Phase 4 — QA

Structural checks (USAGE §5) + **visual render via PowerPoint COM** + judge pass on changed slides:
- Every EN run has its CN nearby; no orphans; **no run after endParaRPr** (nesting check).
- Slide 5: 4 tables, 0 pictures, cell-for-cell match vs source screenshots.
- Plan slides: banner images kept; tables match the EMF source cell-for-cell (data, fills, font colors); skip/merge rules respected; split slides carry week-suffixed titles.
- No overflow over footer bars (slides 3, 7, 8, …).
- Unedited slides byte-identical to the input.
- **Slide count = original + extra weeks from the plan split** (e.g. CW36: 7 → 8). No template reference slide is added anymore.

## Output naming convention

Final bilingual output: `<original_filename>_CH_EN.pptx` (e.g. `weekly report cw 36_CH_EN.pptx`). User-requested experimental variants: `<original>_test.pptx` / `<original>_CH_EN_test.pptx` — never overwrite the approved `_CH_EN` file with an experiment.

## Environment notes (Windows)

- Visual QA: PowerPoint COM `Slides.Export(..., 'PNG', 1600, 900)` (LibreOffice not installed). Judge the exported PNGs.
- **File lock**: if the deck is open in PowerPoint (`PackageNotFoundError` / Errno 13), attach to the running instance via COM and `SaveCopyAs` a work copy; edit the copy. Never force-close the user's PowerPoint.
- EMF screenshots need GDI+ 4x conversion before OCR; write tesseract temp crops to relative paths.

## Pre-translation notes (every new week)

**The user wants the working files kept for reuse.** Do NOT delete the previous week's `_cw*_work/` directory, its CW-week scripts, the OCR helpers, unpacked `source/`/`template/`, or phase intermediates — the next week's translation adapts them. (`build_plan_tables.py` lives in the skill's scripts/; copy it into the week's work dir when adapting.)

Only transient junk may be removed: empty QA dirs, temp OCR crops (`ocr_tmp/`, `*_view.jpg`, zoom crops), `.DS_Store`.

## Notes
- This is a document-translation-and-relayout task — preserve the original look exactly.
- Short Term Flight Test Plan slides ARE handled since CW36 (they were previously out of scope).
- Add new terms to `references/glossary.md` after each run for consistency.
