---
name: dsw_weekly_report_translation
description: "Use this skill whenever the user asks to translate a Dornier Seawings (DSW) weekly status report PowerPoint from Chinese into an English-only version — e.g. '将30-31周周报翻译成英文版', '翻译成英文版', 'translate the weekly report to English', '英文版周报', '生成_en.pptx'. Trigger on Chinese-source weekly reports (DSW研发部工作汇报_CW* / DSW研发部周报_CW* / DSW研发部工作报告_CW*). Produces a format-identical English deck named '<源文件名>_en.pptx'. Key features: aerospace/aircraft design domain glossary (Seastar program, EASA certification, flight test, compliance documents, drawings/MON, structures/stress/systems/avionics), exact preservation of the original slide layout and the embedded data screenshots, and automatic English-font handling (Arial) with text-fit checking so nothing overflows the original shapes. Does NOT depend on the DSG_weekly_report_translation skill (which is the EN→EN/CN bilingual direction)."
---

# DSW Weekly Report Translation (CN → EN, English-only)

Converts a **Chinese-only** DSW weekly report `.pptx` into an **English-only** version. The original layout, images, screenshots, and paragraph structure are preserved exactly; every translatable text paragraph is replaced with its English translation and re-set to a professional English font.

This skill is the CN→EN direction. It is independent of the `DSG_weekly_report_translation` skill (EN→EN/CN bilingual). Do not mix the two.

## Scope

**Handled (translated in place):**
- Cover / title slides
- Table of contents
- Progress & issues bullets (weekly work progress, problems / non-conformances)
- Section titles (e.g. 海星–图纸 / 符合性文件 / 试验, department tags 【研发】【强度】【结构】【系统】【航电】【总体】)
- Status summary lines (drawing counts, compliance-document counts)
- Test-flow diagram labels (LTP/SADD preparation, test article production, material, equipment calibration, COC, lab contract, test execution, test report, legend, test names)
- Upcoming work objectives

**NOT handled / preserved as-is:**
- **Embedded screenshots / chart images** (e.g. drawing-status bar charts, compliance-closure charts) — these are data images and are kept unchanged. Only the surrounding editable text is translated.
- Shapes with only numbers/IDs (e.g. `COC`, `GTR-2FT-330.002`, `MON-200-001.609`) — left untouched.
- Master / layout structure, colors, borders, positions.

## Reference files in this skill

- `references/glossary.md` — CN↔EN aerospace/aircraft-design glossary used for these reports (Seastar program terms, EASA/certification terms, department & specialty names, test/engineering terms).
- `references/conventions.md` — formatting & naming conventions (English font, fit rules, screenshot policy, output naming).
- `scripts/translate_slides.py` — the translation engine (CN paragraph → EN, font re-set, fit pass).
- `scripts/render_qa.py` — render any `.pptx` to per-slide JPGs for visual QA (LibreOffice + pdftoppm).

## Workflow

### Phase 1 — Inventory the source deck

Unpack the source deck and dump every paragraph's exact text (this is the set of strings you must translate):

```bash
mkdir -p "<week>_work/source"
unzip -o -q "<source>.pptx" -d "<week>_work/source"
```

Then list paragraphs with shape names (see the inventory snippet at the end of this file, or run a small python-pptx dump). Build the translation list from that output — every paragraph of prose needs an English translation.

### Phase 2 — Build the mapping JSON

Create `<week>_work/<week>_mapping.json`, a flat dict:
`{"<exact original paragraph text>": "<English translation>", ...}`

- The **key must exactly equal the paragraph's full concatenated run-text** (matching is exact; unmatched paragraphs are left untouched).
- Apply the `references/glossary.md` terminology. Aircraft-design industry terms matter (see glossary).
- Keep already-English IDs/document numbers (`GTR-2FT-330.002`, `MON-200-001.609`, `Teamcenter`, `Compliance tool`, `COC`) inside the sentence as-is.
- Preserve leading/trailing spaces if the original had them (e.g. `"已完成 "` → `"Completed "`).
- **Optional per-paragraph overrides** (only when a paragraph would not fit): the value may be an object instead of a string — `{"en": "...", "sz": 13.5}` forces the run font size in points; `{"en": "...", "spcBef": 4}` sets the paragraph space-before in points. Prefer keeping the original font/spacing; reach for these only to avoid overflow on dense bullet slides.

### Phase 3 — Translate

```bash
python scripts/translate_slides.py --input "<source>.pptx" \
    --mapping "<week>_work/<week>_mapping.json" --out "<source>_en.pptx"
```

The engine:
1. Walks every shape (incl. grouped shapes and table cells) and matches each paragraph against the mapping.
2. Replaces matched paragraphs with a single English run, copying the original run properties (size, bold, color) and setting the Latin and EA typefaces to the English font (default **Arial**, matching the deck's theme and the program's professional style).
3. Runs a **fit pass** on small fixed-width label shapes (< 3.5 in wide): if the English text is wider than the shape, the run font size is reduced (min 6.5 pt) so no label overflows — layout stays identical.
4. Saves `<source>_en.pptx`.

### Phase 4 — QA render

```bash
python scripts/render_qa.py "<source>_en.pptx" "<week>_work/qa_en"
python scripts/render_qa.py "<source>.pptx" "<week>_work/qa_orig"
```

Visually compare slide-by-slide (same slide numbers). Verify:
- Layout/positions/images are identical to the source.
- No text overflow, especially on the test-flow diagram slide (small labels) and the long progress/issues bullets.
- English terminology follows the glossary.
- The 4 embedded chart screenshots are unchanged.

### Phase 5 — Keep the work files

Keep `<week>_work/` (source dump, mapping JSON, QA renders) for the next week — the mapping is the reusable glossary-informed artifact.

## Output naming

Always `<original_filename>_en.pptx` (same folder as the source). Example: `DSW研发部工作汇报_CW30-31.pptx` → `DSW研发部工作汇报_CW30-31_en.pptx`.

## Notes / conventions

- This is translation-and-format-preservation, **not redesign** — keep the original look exactly.
- **English font:** Arial (the deck's theme Latin font; professional, universally available). Do not use CJK fonts (黑体/SimHei) for English text — set both latin and ea typefaces to Arial.
- **Screenshots:** never edit or replace the embedded data screenshots unless the user explicitly asks (offer a chart-rebuild as a separate step, since the OCR→rebuild path can never match the original pixel-for-pixel).
- Add any new term to `references/glossary.md` after each week's run.

## Paragraph inventory snippet (per-week discovery)

```python
from pptx import Presentation
prs = Presentation("<source>.pptx")
for si, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            txt = "".join(r.text for r in para.runs)
            if txt.strip():
                print(si, shape.name, repr(txt))
```
