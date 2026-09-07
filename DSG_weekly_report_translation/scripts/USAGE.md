# Script Usage Guide — Weekly Report Translation

Concrete workflow for producing a CW-week bilingual deck. Practical companion to
`../SKILL.md` — read that first. Verified on CW28–CW31 (macOS/container) and
CW36 (Windows). **Windows is now the primary environment** (PowerShell,
PowerPoint COM, tesseract at `C:\Program Files\Tesseract-OCR`).

The general scripts in this folder are templates. Each week you make CW-week
copies (`translate_text_slides_cw<NN>.py`, `rebuild_flight_test_progress_cw<NN>.py`,
`build_plan_tables_cw<NN>.py` — copy `build_plan_tables.py` into the work dir),
edit in the week's content, and run the 4-phase pipeline. Keep the CW-week
scripts and the work dir for next week.

---

## 0. Prerequisites & environment (Windows)

- Python 3.9+ with `lxml`, `python-pptx`, `Pillow`; `tesseract` on PATH.
- PowerPoint installed → visual QA via COM export (no LibreOffice).
- **File lock**: if the deck is open in PowerPoint, `python-pptx` fails with
  `PackageNotFoundError` / Errno 13. Do NOT kill PowerPoint. Attach and save a copy:

  ```powershell
  $pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
  foreach ($p in $pp.Presentations) { if ($p.Name -eq 'weekly report cw 36_CH_EN.pptx') {
      $p.SaveCopyAs((Join-Path (Resolve-Path '.').Path 'input_current.pptx')) } }
  ```

  This also captures any unsaved manual edits in the open window.

## 1. Setup — unpack source

```bash
cd <working-dir>                       # e.g. .../ppt-translation
mkdir -p _cw<NN>_work/source
cp "weekly report cw <NN>.pptx" _cw<NN>_work/source/src.pptx
cd _cw<NN>_work/source && unzip -o -q src.pptx
cp "C:/Users/glenn/.agents/skills/DSG_weekly_report_translation/scripts/"*.py ../   # OCR helpers
```

Inspect structure (shapes/paragraphs/geometry) with the python-pptx / lxml
snippets below before writing the week's scripts.

```bash
# text dump of every slide
python - <<'EOF'
from pptx import Presentation
prs = Presentation("source/src.pptx")
for i, s in enumerate(prs.slides, 1):
    print(f"--- SLIDE {i} ---")
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            print(f"  [{sh.name}] {sh.text_frame.text[:120]!r}")
        elif sh.shape_type == 13:
            from pptx.util import Emu
            print(f"  [PIC {sh.name}] x={Emu(sh.left).inches:.2f} y={Emu(sh.top).inches:.2f} w={Emu(sh.width).inches:.2f} h={Emu(sh.height).inches:.2f}")
EOF
```

## 2. Phase 1 — text slides (TOC, Highlights/Lowlights, Status)

Copy last week's CW script (or `translate_text_slides.py`), edit the week's data:

- **Slide 2 TOC** `toc_cn`: now `(cn_text, sz)` tuples — default sz `1400`
  (1600 clips behind the team photo on the longest line; `1200` fallback).
- **Slide 3**: `bullets` = `(paragraph_index, marL, cn)`; remove trailing empty
  paragraphs; `tighten(body3, aggressive=True)` (lnSpc→100%, spc→0).
- **Slide 4**: `bullets4` indices relative to post-empty-removal list;
  Rectangle 4 → appended CN paragraph sz=1000 `5F6F82` centered **and** y moved
  to ~4.85in if bullets would run under it.
- ⚠️ **All inline CN runs must be inserted BEFORE `<a:endParaRPr>`**
  (`append_run()` helper). Runs after endParaRPr are silently dropped by
  PowerPoint — this bit CW36 ("Lowlights不足"/banner CN vanished).
- Footer & section labels stay English-only.

```bash
python translate_text_slides_cw<NN>.py source/src.pptx phase1.pptx
```

## 3. Phase 2 — Flight Test Program Progress (slide 5)

Formatting source of truth: `../references/template_slide5_spec.md` (the
template reference slide is no longer inserted into outputs — CW36 decision).

```bash
python grid_detect.py source/ppt/media/<img>.png     # grid lines
python cell_ocr.py <img>.png "[rows]" "[cols]"       # per-cell OCR
# color sampling: PIL Counter on cell crops (see spec file palette)
```

Adapt `rebuild_flight_test_progress.py` → `rebuild_flight_test_progress_cw<NN>.py`;
validate OCR arithmetic (hours/percent/counts). Run:

```bash
python rebuild_flight_test_progress_cw<NN>.py phase1.pptx phase2.pptx
```

After rebuild: slide 5 must have 4 tables, 0 pictures; remove orphan "Oval 2"
highlight `<p:sp>` if present.

## 4. Phase 3 — Short Term Flight Test Plan (slides 6-7 of source)

Full spec: `../references/short_term_plan_tables.md`. Working example:
`build_plan_tables.py` (CW36). Steps:

1. Map pictures via `ppt/slides/_rels/slideN.xml.rels` → `image*.emf`.
   Two per slide: **banner strip** (~0.98in tall, KEEP) and **table** (≥2in
   tall, REPLACE). Distinguish by height.
2. EMF → 4x PNG:

   ```powershell
   powershell -File "C:/Users/glenn/.agents/skills/DSG_weekly_report_translation/scripts/emf_to_png.ps1" source/ppt/media/image10.emf emf
   ```

3. Grid-detect (relax threshold for light-grey verticals), per-cell OCR +
   fill/font-color sampling. **Italic red/blue lines misOCR** — verify by
   viewing 4x crops (slice into ~2500px chunks for reading). Trust the IMAGE
   over the live Excel workbook (paste may predate edits).
4. Adapt `build_plan_tables.py` data blocks (rows/cols/colors per week) and run:
   `python build_plan_tables_cw<NN>.py <in.pptx> <out.pptx>` — it does pass A
   (tables on the plan slides), slide-clone (split), pass B (week-2 table).
5. Multi-week image ⇒ one slide per week; clone rels must DROP the
   notesSlide relationship; register slide in `[Content_Types].xml` (Override,
   `/ppt/slides/slideN.xml`), `presentation.xml.rels` (new rId), `sldIdLst`
   (fully-qualified `r:id` namespace).

## 5. Phase 4 — QA

Structural:

```bash
python - <<'EOF'
import zipfile
from lxml import etree
from pptx import Presentation
out = "weekly report cw <NN>_CH_EN.pptx"
prs = Presentation(out)
print("slides:", len(prs.slides))            # = original + split weeks (CW36: 8)
zf = zipfile.ZipFile(out); print("zip ok:", zf.testzip() is None)
for n in zf.namelist():
    if n.endswith('.xml'): etree.fromstring(zf.read(n))
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
bad = sum(1 for s in prs.slides for p in s._element.iter(A+'p')
          for i,c in enumerate(p) if c.tag==A+'r' and any(x.tag==A+'endParaRPr' for x in p[i+1:]))
print("runs after endParaRPr:", bad)          # must be 0
# unedited slides byte-identical to input:
def body(f,i): return etree.tostring(etree.fromstring(zipfile.ZipFile(f).read(f'ppt/slides/slide{i}.xml')))
# ... compare appropriate indices ...
EOF
```

Visual render (Windows PowerPoint COM):

```powershell
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open((Resolve-Path 'out.pptx').Path, $true, $false, $false)
$out = (Resolve-Path 'qa').Path
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
  $pres.Slides.Item($i).Export((Join-Path $out ('slide-' + $i + '.png')), 'PNG', 1600, 900) }
$pres.Close(); $pp.Quit()
```

Then run a judge pass on the rendered PNGs of every changed slide (bilingual
pairing, colors vs source, no overflow over footer bars, banner kept on plan
slides, table data cell-for-cell vs the EMF source).

Row-height sanity: if a plan table collides with the footer bar, compress
lnSpc to 95% and cell marT/B to 9000 before shrinking anything else.

## 6. Output naming & wrap-up

```bash
cp phase3.pptx "weekly report cw <NN>_CH_EN.pptx"
```

- Approved output: `<original>_CH_EN.pptx`. Experiments: `<original>_test.pptx`
  — keep both, never overwrite the approved file with an experiment.
- Keep `_cw<NN>_work/` (CW scripts, OCR helpers, unpacked source, phase
  intermediates, `input_current.pptx` lock-copies) for next week.
- Append new terms to `../references/glossary.md`; update the spec files if a
  convention changes.
