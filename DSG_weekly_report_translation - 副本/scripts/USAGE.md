# Script Usage Guide — Weekly Report Translation

Concrete, copy-pasteable workflow for producing a CW-week bilingual deck. This is the *practical* counterpart to `../SKILL.md` — read that first for the conventions, use this for the exact commands. Verified on CW28–CW31 runs.

**The general scripts in this folder are templates.** Each week you make two CW-week copies (`translate_text_slides_cw<NN>.py`, `rebuild_flight_test_progress_cw<NN>.py`), edit in the week's content, and run the 3-phase pipeline. Keep the CW-week scripts and the work dir for next week — the user wants them preserved.

---

## 0. Prerequisites & environment

- Python 3.9+, with `lxml`, `python-pptx`, `Pillow`.
- `tesseract` for screenshot OCR (grid detection + per-cell OCR). *macOS note:* tesseract fails on absolute `/tmp/...` paths with `Image file ... cannot be read!` — always write temp crops to a **local relative dir** (e.g. `ocr_tmp/`) and pass relative paths.
- LibreOffice for rendering is **not present** on this machine; visual QA via rendering is unavailable. Fall back to the structural/analytical QA in Phase 4.
- PowerPoint exists but its AppleScript PDF/PNG export blocks on a Save dialog — don't rely on it.

## 1. Setup — unpack source & template

```bash
cd <working-dir>                       # e.g. .../ppt-translation
mkdir -p _cw<NN>_work/source _cw<NN>_work/template
cp "weekly report cw <NN>.pptx" _cw<NN>_work/source/src.pptx
cd _cw<NN>_work/source && unzip -o -q src.pptx
cp ~/.claude/skills/DSG_weekly_report_translation/references/template.pptx ../template/template.pptx
cd ../template && unzip -o -q template.pptx
```

Inspect slide structure so you know what you're translating:

```bash
# text dump of every slide (shapes + paragraphs)
python3 - <<'EOF'
from pptx import Presentation
prs = Presentation("src.pptx")
for i, s in enumerate(prs.slides, 1):
    print(f"--- SLIDE {i} ---")
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            print(f"  [{sh.name}] {sh.text_frame.text[:120]}")
        elif sh.shape_type == 13:
            print(f"  [PIC] {sh.image.filename}")
        elif sh.has_table:
            print(f"  [TABLE]")
EOF
```

Useful introspection snippets (paragraph indices, `marL`/`indent`, run sizes, colors):

```bash
python3 - <<'EOF'
from lxml import etree
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
tree=etree.parse("ppt/slides/slide3.xml")
for sp in tree.iter(P+'sp'):
    tb=sp.find(P+'txBody')
    if tb is None: continue
    for i,p in enumerate(tb.findall(A+'p')):
        pPr=p.find(A+'pPr')
        attrs={}
        if pPr is not None:
            attrs['marL']=pPr.get('marL'); attrs['indent']=pPr.get('indent')
        txt=''.join((t.text or '') for t in p.iter(A+'t'))
        if txt.strip(): print(i, attrs, txt[:60])
EOF
```

## 2. Phase 1 — text slides (TOC, Highlights/Lowlights, Status)

1. Copy the template: `cp scripts/translate_text_slides.py scripts/translate_text_slides_cw<NN>.py` (or reuse the previous week's CW script).
2. Edit the per-week data blocks in the script:
   - **Slide 2 TOC** `toc_cn` dict — one entry per TOC line.
   - **Slide 3**: section-header CN (`亮点` / `不足`), and the `bullets` list `(paragraph_index, marL, cn_text)`. `marL` = EN paragraph's `marL + indent` (text start). Indices are **relative to the original paragraph list**; the script inserts in reverse order so earlier indices stay valid.
   - **Slide 4**: `bullets4` list. **Indices are relative to the post-empty-removal list** (empties removed first). Compute from the slide dump.
   - Callout boxes (Rectangle 7 / Rectangle 4) get inline CN runs (sz=1000).
3. Run:
   ```bash
   python3 translate_text_slides_cw<NN>.py source/src.pptx phase1.pptx
   ```
4. Verify: dump slide 2/3/4 and confirm every EN line has its CN below/beside it.

**⚠️ Convention gotchas (match the approved output, not the template docs alone):**
- **Footer** ("Program risks are outlined…") stays **English-only**. The bundled `translate_text_slides.py` appends footer CN *inside* the EN `<a:r>` (malformed nested run) — do NOT use that block. Delete it in your CW script.
- Slide-3 section headers **do** get inline CN (`Highlights 亮点` / `Lowlights 不足`), inserted *between* the "Highlights" run and the trailing-tabs run.
- Slide-4 section labels (`SN1002:`, `SN1003:`, `SN1004:`, `Status:`) stay English-only.
- Rectangle 7 → inline CN per line; Rectangle 4 → CN as a **separate centered paragraph** (sz=1000, grey `5F6F82`) under the EN lines.

## 3. Phase 2 — Flight Test Program Progress (screenshot → native tables)

The source slide 5 is screenshots. Extract data + colors, then rebuild 4 native bilingual tables.

### 3a. Locate the screenshots

```bash
python3 - <<'EOF'
from pptx import Presentation
from pptx.util import Emu
prs = Presentation("src.pptx"); s5 = prs.slides[4]
for sh in s5.shapes:
    print(sh.shape_type, sh.name, f"x={Emu(sh.left).inches:.2f} y={Emu(sh.top).inches:.2f} w={Emu(sh.width).inches:.2f} h={Emu(sh.height).inches:.2f}")
EOF
```
Typical CW layout: left = Weekly Overview + Block/Flight (one tall screenshot), middle = Flight Hours & Flights, right = Progress Summary. Check `ppt/slides/_rels/slide5.xml.rels` to map pictures→media files.

### 3b. OCR the table data

Use the helper scripts (`cell_ocr.py`, `grid_detect.py`, `ascii_view.py` — kept in the previous week's `_cw<NN>_work/`):

```bash
# 1. detect grid lines (returns row/col midpoints per image)
python3 grid_detect.py ppt/media/<image>.png

# 2. ASCII "view" of each screenshot to sanity-check structure
python3 ascii_view.py ppt/media/<image>.png 110 40

# 3. per-cell OCR: pass detected h/v lines as eval'd lists
python3 cell_ocr.py ppt/media/<image>.png "[13,78,163,229,294,360]" "[6,254,465,677,904]"
```
- Upscale + binarize happens inside `cell_ocr.py`. If a cell OCRs empty, re-run it with a **wider box** or a different threshold — column separators are often 1–2px and land slightly off the detected line.
- **Validate arithmetic**: flight hours/flights columns must be internally consistent (e.g. `500:00 − 326:31 = 173:29`, `326:31/500:00 ≈ 65%`, per-SN flight counts sum). If a number breaks the pattern, re-OCR it — it's usually a misread (e.g. `99` vs `55`).

### 3c. Sample the fill colors

```bash
python3 - <<'EOF'
from PIL import Image
from collections import Counter
def sample(img, box, label):
    im = Image.open(img).convert('RGB'); c = Counter(im.crop(box).getdata()).most_common(1)[0]
    print(f"{label}: #{c[0][0]:02X}{c[0][1]:02X}{c[0][2]:02X}")
# sample an empty corner of each cell/row you care about
sample('ppt/media/image10.png',(40,20,120,60),"overview hdr")
EOF
```
Record each region's color (label column vs SN1003 column vs SN1004 column vs header rows) — the source wins over the template.

### 3d. Build & run the rebuild script

Copy the previous week's `rebuild_flight_test_progress_cw<NN>.py` (or `scripts/rebuild_flight_test_progress.py`), then edit the four data arrays + color constants to match **this week's source** (structure, merged cells, fills all from the source; font typeface/size and border weights from the template). Run:

```bash
python3 rebuild_flight_test_progress_cw<NN>.py phase1.pptx phase2.pptx
```
Verify: slide 5 has 4 tables, 0 pictures, and every cell matches the OCR'd data.

**⚠️ Leftover decoration:** the source slide 5 often has a transparent "Oval 2" (a highlight circle over the screenshot). After the screenshot is replaced it's an orphan over the new table — remove that `<p:sp>` from `ppt/slides/slide5.xml` (edit the unpacked XML and repack, or patch the final zip).

## 4. Phase 3 — insert template reference slide

```bash
python3 scripts/insert_template_slide.py phase2.pptx phase3.pptx
# (uses the skill's references/template.pptx by default; pass a 3rd arg to override)
```
Result: 8 slides — 1 Cover, 2 TOC, 3 Highlights/Lowlights, 4 Status, 5 Flight Progress (data), 6 Flight Progress (template reference), 7–8 Short Term Plan (untouched).

**⚠️ Check before running:** the script assumes slide rIds `slide1-5=rId5-9, slide6=rId10, slide7=rId11` and shifts 6→7, 7→8. Verify against the source:

```bash
grep -o 'Id="rId[0-9]*"[^>]*slide[0-9]*\.xml' ppt/_rels/presentation.xml.rels
```
If the mapping differs, adjust the hardcoded `rids` list / rel-renaming in the script.

## 5. Phase 4 — QA (no renderer available)

Run these checks; all passed for CW31:

```bash
python3 - <<'EOF'
import zipfile
from lxml import etree
from pptx import Presentation
out = "weekly report cw <NN>_CH_EN.pptx"
prs = Presentation(out)
print("slides:", len(prs.slides))                       # = original + 1
zf = zipfile.ZipFile(out)
print("zip valid:", zf.testzip() is None)
for n in zf.namelist():
    if n.endswith('.xml'): etree.fromstring(zf.read(n)) # well-formed
print("slide5 tables:", sum(1 for s in prs.slides[4].shapes if s.has_table))  # 4
print("slide5 pics:",   sum(1 for s in prs.slides[4].shapes if s.shape_type==13)) # 0
print("slide7/8 pics:", sum(1 for s in prs.slides[6].shapes if s.shape_type==13),
      sum(1 for s in prs.slides[7].shapes if s.shape_type==13)) # 2,2 (untouched)
EOF
```

Additional structural checks:
- **Untouched Short-Term Plan:** compare `out slide7/8` bodies to `src slide6/7` after stripping the `<?xml …?>` declaration — only the declaration quote style and trailing newline may differ (python-pptx re-serialization). Embeds and `../media/*.emf` targets must match.
- **Template reference slide:** `out slide6` must equal `template.pptx slide5` (md5 the unpacked XML).
- **Bilingual pairing:** every English bullet has a CN line; every CN has its EN (dump slide XML and check adjacency).
- **No malformed nesting:** assert no `<a:r>` has a parent that is also `<a:r>` (the footer bug from Phase 1 would show up here).

Overflow: without a renderer, estimate per-shape (font size × text length vs box width/height) only as a sanity check; recommend a quick manual look in PowerPoint at slide 4 callout boxes and slide 5 tables.

## 6. Output naming & wrap-up

```bash
cp phase3.pptx "weekly report cw <NN>_CH_EN.pptx"
```
- Final output must be `<original_filename>_CH_EN.pptx` in the working directory.
- **Keep** `_cw<NN>_work/` (CW scripts, OCR helpers `cell_ocr.py`/`grid_detect.py`/`ascii_view.py`, unpacked source/template, phase1–3 outputs) — the user wants these for next week.
- Add any new terms to `../references/glossary.md` and update `../SKILL.md` if a convention changes.
