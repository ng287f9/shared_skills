---
name: DSG_weekly_report_translation
description: "Use this skill whenever the user asks to translate, localize, or produce a bilingual (Chinese/English) version of a weekly status report, program status report, or flight test report PowerPoint — especially aerospace/aircraft program reports (Dornierseawings / Seastar / CW-numbered weekly reports). Trigger on phrases like '周报', '双语', '中英文版', '中英对照', 'weekly report translation', 'bilingual pptx', or when the user uploads an English-only .pptx weekly report and asks for a Chinese+English version. Ships with the approved bilingual template deck (references/template.pptx), helper scripts, exact bilingual formatting conventions, and a domain glossary for consistent terminology. Handles: text slides (cover, TOC, highlights/lowlights, status), Flight Test Program Progress (screenshot→native bilingual table + template reference slide copy), and Short Term Flight Test Plan (left untouched)."
---

# Weekly Report Translation (EN → EN/CN bilingual)

Converts English-only slides in a weekly status report `.pptx` into bilingual English+Chinese, following the exact formatting conventions used in this program's reports. Chinese is always added as a supplement to the English — never replaces it.

**Scope:**
- **Handled:** Cover page (English-only), Table of Contents (inline CN), Weekly Highlights / Lowlights (bullet CN below), Status SN#### (bullet CN below), **Flight Test Program Progress** (screenshot tables → native bilingual tables + template reference slide copied as new slide).
- **NOT handled:** Short Term Flight Test Plan slides — left completely untouched.

**Always read `/mnt/skills/public/pptx/SKILL.md` and `editing.md` first.** This skill only adds the bilingual-specific conventions and glossary on top of the general pptx editing workflow — it does not replace it (unpack → edit XML → clean → pack, thumbnail-based QA, etc. all still apply).

## Reference files and scripts in this skill

- `references/template.pptx` — **The approved bilingual template deck (9 slides).** Slide 5 is the reference for Flight Test Program Progress (its table structure is copied into the output as a new slide after the rebuilt data slide). Unpack it and read its XML whenever a rule is ambiguous.
- `references/glossary.md` — English↔Chinese term glossary for this program.
- `references/formatting_patterns.md` — Exact XML patterns for each bilingual layout.
- `scripts/translate_text_slides.py` — Translates cover, TOC, highlights/lowlights, and status slides (Bucket B/C formatting). Operates on unpacked slide XML.
- `scripts/rebuild_flight_test_progress.py` — Rebuilds Flight Test Program Progress tables as native bilingual tables (Bucket D). Reads source screenshots via OCR, uses template styling.
- `scripts/insert_template_slide.py` — Copies template slide 5 (reference Flight Test Program Progress) as a new slide after the data slide in the output. Handles slide shifting, relationship updates, and content type registration.
- `scripts/render_slides.py` — Render any pptx to per-slide JPGs for visual QA.
- `scripts/extract_table_style.py` — Dump a slide's native-table structural style.
- **`scripts/USAGE.md` — READ THIS FIRST for the concrete, copy-pasteable workflow** (setup, per-phase commands, OCR/color-extraction, QA checks, and convention gotchas like the English-only footer and the Rectangle-4 separate-CN-paragraph). It is the practical companion to the phases below.

## Workflow (3 phases)

### Phase 1 — Text slides: translate and format

**Create a CW-week-specific translation script** (e.g. `translate_text_slides_cw30.py`) by adapting the general `scripts/translate_text_slides.py` pattern — read the source slide XMLs to identify paragraph indices and content, then hardcode the week's translations. Each week's content differs, so the script must be rebuilt per-week based on actual slide content.

Run it which:

1. Unpacks the source `.pptx` and the template `.pptx`.
2. **Slide 2 (TOC):** Appends Chinese inline on each TOC line (Bucket B: same line, smaller sz, muted slate-grey `5F6F82`).
3. **Slide 3 (Highlights/Lowlights):** Section headers get inline CN (Bucket B); each bullet gets a Chinese paragraph below it (Bucket C: new `<a:p>`, `buNone`, grey `808080`, indented to text start). Footer text gets inline CN.
4. **Slide 4 (Status):** Same Bucket C approach. Removes empty spacer paragraphs. Callout boxes (Rectangle 7, Rectangle 4) get inline CN (Bucket B).
5. Tightens paragraph spacing and line spacing to fit content.
6. Repacks the output `.pptx`.

**Slides left untouched:** 1 (cover — English-only), 5-7 (table/screenshot slides passed through as-is at this stage).

### Phase 2 — Flight Test Program Progress: native bilingual tables

**Create a CW-week-specific rebuild script** (e.g. `rebuild_flight_test_progress_cw30.py`) by adapting the general `scripts/rebuild_flight_test_progress.py` pattern — OCR the source slide 5 screenshots to extract the week's data, then hardcode the table cell values.

Run it which:

1. Reads the Phase 1 output `.pptx`.
2. On **slide 5 (Flight Test Program Progress):**
   - Removes all screenshot images (`<p:pic>` elements).
   - Rebuilds 4 native bilingual tables using template slide 5 formatting (font typeface/size, border weight) populated with CW-week data from the source screenshots.
   - Tables: Weekly Overview (8×3), Flight Hours & Flights (5×9), Progress Summary (5×4), Weekly Block/Flight Time (3×3).
   - Each cell: English paragraph + Chinese paragraph below it, same font color (Bucket D).
3. Saves the updated `.pptx`.

**Where table attributes come from:**

| Attribute | Source of truth |
|---|---|
| Table structure, column widths, merged cells | Template slide 5 |
| Font typeface, font size, bold/italic | Template slide 5 |
| Border line width/weight/style (w=12700 outer, 6350 inner) | Template slide 5 |
| Cell content / data values | Source English screenshot (OCR) |
| Font color (EN + CN identical in each cell) | Source English screenshot |
| Cell/table background fill color | Source English screenshot |
| Row height | Auto-fit to that row's actual bilingual content |

Common fill colors: `DAF2D0` (green header), `DAE9F8` (blue header), `FFFF00` (yellow data), `F2F2F2` (grey totals).

### Phase 3 — Insert template reference slide

Run `scripts/insert_template_slide.py` which:

1. Unpacks the Phase 2 output and the template.
2. **Shifts slides:** moves existing slide 6 → 7, slide 7 → 8.
3. **Copies template slide 5** (the bilingual Flight Test Program Progress reference with its table structure and example data) → new **slide 6**.
4. The output now has **8 slides**:
   - 1: Cover (English-only)
   - 2: TOC (bilingual)
   - 3: Highlights/Lowlights (bilingual)
   - 4: Status (bilingual)
   - **5: Flight Test Program Progress** (CW-week data, native bilingual tables)
   - **6: Flight Test Program Progress** (template reference copy with example data)
   - 7: Short Term Flight Test Plan SN1003 (untouched)
   - 8: Short Term Flight Test Plan SN1004 (untouched)
5. Updates `[Content_Types].xml`, `presentation.xml` (sldIdLst), and `presentation.xml.rels` to register the new slide correctly.
6. Verifies the output has 8 valid slides.

**⚠️ Content_Types paths:** When adding overrides for new slides, use the correct path format: `/ppt/slides/slideN.xml` (not `/ppt/slideN.xml`). Same for notes: `/ppt/notesSlides/notesSlideN.xml`.

**⚠️ sldIdLst r:id:** When creating `<p:sldId>` elements, the `r:id` attribute must use the `{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id` fully-qualified namespace URI. Register the `r` and `p` namespaces before parsing the XML to ensure proper serialization.

### Phase 4 — QA

- Every English text run has a corresponding Chinese translation nearby.
- No Chinese line is orphaned without its English counterpart.
- No text overflow on any slide (verify with visual render).
- Flight Test Program Progress tables (slide 5) match source data cell-for-cell.
- Template reference slide (slide 6) shows the template's bilingual table layout correctly.
- Short Term Flight Test Plan slides (7-8) are untouched — same images, same text.
- Total slide count = original + 1 (added template reference).

## Output naming convention

The final bilingual output **must** be named `<original_filename>_CH_EN.pptx`. For example, `weekly report cw 30.pptx` → `weekly report cw 30_CH_EN.pptx`.

## Pre-translation notes (every new week)

**The user wants the working files kept for reuse.** Do NOT delete the previous week's `_cw*_work/` directory, its CW-week scripts (`translate_text_slides_cw*.py`, `rebuild_flight_test_progress_cw*.py`), the OCR helper scripts (`cell_ocr.py`, `grid_detect.py`, `ascii_view.py`), the unpacked `source/`/`template/`, or the `phase1/2/3.pptx` intermediates — the next week's translation adapts them.

Only transient junk may be removed if present: empty QA dirs, failed-export AppleScripts, temp OCR crops (`ocr_tmp/`, `*_view.jpg`, `*_hi.png`), `.DS_Store`.

## Notes
- This is a document-translation-and-relayout task, not a from-scratch design — preserve the original look exactly.
- The template deck (`references/template.pptx`) ships with the skill — do not ask the user for one unless they volunteer a newer version.
- **Short Term Flight Test Plan slides are out of scope** — leave them completely untouched at every phase.
- Add new terms to `references/glossary.md` after each run for consistency.
