# Bilingual Formatting Patterns

These patterns are taken directly from the approved bilingual template shipped with this skill at **`references/template.pptx`** (9 slides: cover, TOC, highlights/lowlights, status, flight-test progress, and 4 split short-term-plan slides). Match them exactly — don't invent a new bilingual style. When anything below is ambiguous, unpack `references/template.pptx` and read the equivalent slide's XML directly; run `scripts/extract_table_style.py references/template.pptx <slide#>` to dump any table's structural style. All XML below is drawingml/pptx (`a:` namespace) inside an unpacked slide's `slideN.xml`.

Template quick facts (verified from the file):
- Body bullets (slide 3): EN sz=1400, CN translation line sz=1400 grey `808080` — Bucket C below.
- Flight Test Program Progress (slide 5): 4 native tables, font Arial (latin) + SimSun-ExtB (ea), mostly sz=600, some sz=700; structural fills like `DAF2D0`/`FFFF00`/`DAE9F8`/`F2F2F2` per the source data's own coloring.
- Short Term Flight Test Plan (slides 6–9): one table per slide, 6 columns (first ~1700870 EMU, five weekday columns ~1753044 EMU each), borders `w="6350"` solid black on all four sides of every cell, font Arial + SimSun-ExtB sz=700 bold for status entries; label-row fill purple `D071FF`, header/non-working grey `C0C0C0`; run colors are data (black `000000`, red `FF0000`, blue `0000FF`).

## Bucket A — Title placeholders (stay English-only)

The main slide title placeholder (`<p:ph type="title"/>`) and the cover slide's title/date box are **not** translated. Leave them exactly as in the English source. This includes the cover slide's "Weekly Status Report / CW23 / June 6th, 2026" and each content slide's title placeholder (e.g. "Flight Test Program Progress").

Exception: if the user explicitly asks for the title to be bilingual too, apply the Bucket B pattern to it instead.

## Bucket B — Section header / TOC lines (inline, smaller, muted color)

Used for short banner-style single lines: TOC/agenda entries, section headers, callout box headers. The Chinese is appended **on the same line**, not a new paragraph, usually after a trailing space. Chinese runs are smaller than the English (roughly 80% of the English run's size) and use a muted slate-grey (`5F6F82`), while the English stays whatever color it already was (often navy `002060` for TOC/section titles).

```xml
<a:p>
  <a:r>
    <a:rPr lang="en-GB" sz="2000">
      <a:solidFill><a:srgbClr val="002060"/></a:solidFill>
    </a:rPr>
    <a:t>Weekly Highlights / Lowlights </a:t>
  </a:r>
  <a:r>
    <a:rPr lang="zh-CN" altLang="en-US" sz="1600">
      <a:solidFill><a:srgbClr val="5F6F82"/></a:solidFill>
      <a:latin typeface="+mj-ea"/>
      <a:ea typeface="+mj-ea"/>
    </a:rPr>
    <a:t>本周亮点 </a:t>
  </a:r>
  <a:r>
    <a:rPr lang="en-US" sz="1600">
      <a:solidFill><a:srgbClr val="5F6F82"/></a:solidFill>
      <a:latin typeface="+mj-ea"/>
      <a:ea typeface="+mj-ea"/>
    </a:rPr>
    <a:t>/ </a:t>
  </a:r>
  <a:r>
    <a:rPr lang="zh-CN" altLang="en-US" sz="1600">
      <a:solidFill><a:srgbClr val="5F6F82"/></a:solidFill>
      <a:latin typeface="+mj-ea"/>
      <a:ea typeface="+mj-ea"/>
    </a:rPr>
    <a:t>风险及问题</a:t>
  </a:r>
</a:p>
```

Rule of thumb: English run size `sz`, Chinese run size ≈ `sz * 0.8` (round to nearest 100, i.e. nearest 1pt in the `sz` hundredths-of-a-point unit — e.g. 2000→1600, 1600→1200), same pattern of alternating EN/CN runs if the sentence mixes languages mid-line (e.g. "SN1002 / SN1003 / SN1004" stays as Latin runs, only the trailing label word gets a CN run appended).

## Bucket C — Body bullet / paragraph text (new line below, no bullet, grey)

Used inside content placeholders for Highlights/Lowlights and similar narrative bullets. Structure:
- The English bullet paragraph keeps its original bullet (`buChar`/`buFont`) and `marL`/`indent`.
- Immediately after it, insert a **new sibling `<a:p>`** containing only the Chinese translation:
  - `buNone` (no bullet glyph)
  - `marL` set to the English bullet's *text start position* (i.e. `marL + indent` of the English paragraph), `indent="0"` — this makes the Chinese line start flush under the English text, not under the bullet glyph.
  - Same font size (`sz`) as the English bullet text.
  - Text color grey `808080` (via `<a:solidFill><a:srgbClr val="808080"/></a:solidFill>` on each run).
  - `spcBef`/`spcAft` set to `0` so the translation hugs its English line tightly (visually reads as one bullet, two lines) rather than looking like a separate bullet point.

```xml
<!-- English bullet -->
<a:p>
  <a:pPr marL="714375" indent="-355600">
    <a:spcBef><a:spcPts val="0"/></a:spcBef>
    <a:spcAft><a:spcPts val="0"/></a:spcAft>
    <a:buFont typeface="Wingdings" pitchFamily="2" charset="2"/>
    <a:buChar char="Ø"/>
  </a:pPr>
  <a:r>
    <a:rPr lang="en-US" sz="1400" dirty="0"/>
    <a:t>Management Meeting held on Thursday with MoM available was Honeywell RL5.2.1 software download approved</a:t>
  </a:r>
</a:p>

<!-- Chinese translation, new paragraph, no bullet, aligned under text not bullet -->
<a:p>
  <a:pPr marL="358775" indent="0">
    <a:spcBef><a:spcPts val="0"/></a:spcBef>
    <a:spcAft><a:spcPts val="0"/></a:spcAft>
    <a:buNone/>
  </a:pPr>
  <a:r>
    <a:rPr lang="zh-CN" altLang="en-US" sz="1400" dirty="0">
      <a:solidFill><a:srgbClr val="808080"/></a:solidFill>
      <a:ea typeface="Arial" pitchFamily="34" charset="0"/>
    </a:rPr>
    <a:t>       周四召开的管理层会议已生成会议纪要，会上批准了霍尼韦尔</a:t>
  </a:r>
  <a:r>
    <a:rPr lang="en-US" altLang="zh-CN" sz="1400" dirty="0">
      <a:solidFill><a:srgbClr val="808080"/></a:solidFill>
      <a:ea typeface="Arial" pitchFamily="34" charset="0"/>
    </a:rPr>
    <a:t>RL5.2.1</a:t>
  </a:r>
  <a:r>
    <a:rPr lang="zh-CN" altLang="en-US" sz="1400" dirty="0">
      <a:solidFill><a:srgbClr val="808080"/></a:solidFill>
      <a:ea typeface="Arial" pitchFamily="34" charset="0"/>
    </a:rPr>
    <a:t>软件的下载</a:t>
  </a:r>
</a:p>
```

Notes:
- Note the leading spaces in the Chinese `<a:t>` (`"       周四..."`) — the reference deck pads with spaces to visually indent since `marL` alone sometimes isn't enough given the font; match this padding style if the rendered result doesn't align well with pure `marL`.
- When the English sentence contains an embedded Latin acronym/version number (e.g. "RL5.2.1"), split the Chinese line into multiple runs so the acronym stays in a Latin run (`lang="en-US"`) sandwiched between Chinese runs (`lang="zh-CN"`), exactly as shown above.
- If a bullet header line mixes bold+underline formatting with an inline translation (e.g. "Highlights亮点"), that's Bucket B style instead (inline, same paragraph) — use judgement: **section-label headers within a body placeholder use inline Bucket B style; the actual content bullets below them use Bucket C new-line style.**

## Bucket D — Table cells (new paragraph inside cell, same size/color as English)

Inside a `<a:tc>` cell's `<a:txBody>`, add the Chinese as a second `<a:p>` directly after the English paragraph, same `sz`, **same color as English** (typically black `000000`, not greyed out — tables use full color for both languages since space/contrast is already tight), `algn` matching the cell's existing alignment (e.g. `ctr` for centered schedule cells), `buNone`.

```xml
<a:tc>
  <a:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr" fontAlgn="ctr"><a:buNone/></a:pPr>
      <a:r>
        <a:rPr lang="en-GB" sz="700" b="1" i="0" u="none" strike="noStrike">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:effectLst/>
          <a:latin typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
          <a:ea typeface="SimSun-ExtB" panose="02010609060101010101" pitchFamily="49" charset="-122"/>
        </a:rPr>
        <a:t>Maintenance</a:t>
      </a:r>
    </a:p>
    <a:p>
      <a:pPr algn="ctr" fontAlgn="ctr"><a:buNone/></a:pPr>
      <a:r>
        <a:rPr lang="zh-CN" altLang="en-US" sz="700" b="1" i="0" u="none" strike="noStrike">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:effectLst/>
          <a:latin typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>
          <a:ea typeface="SimSun-ExtB" panose="02010609060101010101" pitchFamily="49" charset="-122"/>
        </a:rPr>
        <a:t>维护</a:t>
      </a:r>
    </a:p>
  </a:txBody>
  <a:tcPr marL="0" marR="0" marT="0" marB="0" anchor="ctr">
    <!-- all four borders: w="6350" solid; fill from source screenshot, e.g. purple label row: -->
    <a:lnL w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnL>
    <a:lnR w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnR>
    <a:lnT w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnT>
    <a:lnB w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnB>
    <a:solidFill><a:srgbClr val="D071FF"/></a:solidFill>
  </a:tcPr>
</a:tc>
```

This is a verbatim cell from `references/template.pptx` slide 6 (the purple "Maintenance 维护" label row). Note: EN and CN are separate `<a:p>` paragraphs in the same cell, identical `sz`/`b`/color, both carrying `<a:latin>Arial</a:latin>` + `<a:ea>SimSun-ExtB</a:ea>`; mixed-language lines inside a cell (e.g. "25/50/100/200小时定检（75%）") split into alternating Latin/Chinese runs the same way as Bucket C.

If adding the second paragraph makes the cell content taller than the row, always **grow the row height to fit** the combined English+Chinese content rather than shrinking the font — row heights in a rebuilt table should be auto-fit per row's actual content, never copied as a fixed value from a template. Only shrink table font as a last resort if a row genuinely cannot grow (e.g. it would push the table off the slide), and never below ~6pt (`sz="600"`).

**⚠️ Common mistake to avoid**: do not carry over the Bucket C convention (grey `808080` Chinese text) into tables. Inside table cells, the English run's `<a:solidFill>` color and the Chinese run's `<a:solidFill>` color must be **identical to each other** in every case.

**⚠️ General rule for rebuilt/screenshot tables — where each attribute comes from:**
- **Content, font color, cell/table background color, and table style/structure** (row & column arrangement, merged cells, header/label rows, block layout, relative column proportions) → always from **the source (to-be-translated) English screenshot**, read cell by cell. This applies to every rebuilt table, not just specific named ones — data, colors, and structure are never invented or borrowed from the template. If the source's table structure differs from the template's example, the source wins.
- **Font typeface and font size** → from **the template deck's** equivalent table.
- **Table border line width/weight and line type** → from **the template deck's** equivalent table.
- **Row height** → auto-fit to that row's actual bilingual content, not copied fixed from either source.

Whatever color the English run in a given cell uses (read from the source screenshot), copy that exact same color to the Chinese run directly below it in the same cell.

## Choosing between Bucket B and Bucket C for a given text box
Ask: is this a short single-line label/header (→ B, inline) or a wrapping multi-line narrative bullet (→ C, new line below)? When in doubt, check the reference deck's equivalent slide/element directly rather than guessing — open `unpacked/ppt/slides/slideN.xml` from the reference deck and look at the analogous element.
