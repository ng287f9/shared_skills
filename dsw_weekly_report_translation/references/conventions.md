# Formatting & naming conventions (DSW CN→EN translation)

## English font
- **Use Arial** for all translated runs — it is the deck's theme Latin font and the program's professional standard.
- Set **both** `<a:latin typeface="Arial"/>` and `<a:ea typeface="Arial"/>` on every translated run so no CJK font (黑体 / SimHei / +mn-ea) is left in the English deck.
- Copy the original run's scalar properties (size `sz`, `b`/`i`, color `solidFill`, baseline `baseline`) onto the new run — the English text must keep the same size/bold/color as the Chinese it replaces.

## Text-fit rule (small fixed labels)
- Small fixed-width shapes (< 3.5 in wide) are treated as single-line labels. If the English text is wider than the shape's content area, **reduce the run font size** until it fits (min 6.5 pt, step 0.5 pt). Keep the term full and professional — prefer a slight size reduction over abbreviations.
- Body text placeholders (wide, wrapping) are not fit-reduced: English is typically ~30–40 % shorter than the Chinese, so it fits.

## What must stay untouched
- Embedded screenshots / chart images (data screenshots) — never edit or replace unless the user explicitly asks.
- Already-English IDs and acronyms: `COC`, `MON-…`, `GTR-…`, `ANL-…`, `WOT-…`, `SPO`, `LTP`, `SADD`, `Teamcenter`, `Compliance tool`, `EASA`, `TC`, `1004`/`MSN1004`, `CW##`.
- Paragraph properties (`a:pPr` — bullets, indents, spacing), body properties, positions, sizes, colors, master/layout structure.

## Punctuation & spacing
- Preserve a leading/trailing space when the original had one (e.g. `"已完成 "` → `"Completed "`); set `xml:space="preserve"` on `<a:t>` so the space survives.
- Use the en-dash for ranges (CW30–31). Keep the original `–`/`~` style consistent with the source where it is a title range.

## Output naming
- `<source_filename>_en.pptx`, same folder as the source (e.g. `DSW研发部工作汇报_CW30-31_en.pptx`).
- Never overwrite the source.

## QA checklist (after every run)
1. Slide count and layout identical to source.
2. No Chinese remains in editable text (screenshots excepted).
3. No overflow: flow-diagram labels, long progress/issues bullets.
4. Terminology follows `references/glossary.md`.
5. 4 embedded chart screenshots unchanged.
