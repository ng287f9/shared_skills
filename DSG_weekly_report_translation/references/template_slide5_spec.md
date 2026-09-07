# Template Slide 5 Spec — Flight Test Program Progress (bilingual tables)

Format reference extracted from the approved bilingual template deck
(`references/template.pptx`, slide 5). **As of CW36 the template reference slide
is no longer copied into the output deck** — rebuild the data slide from this
spec instead. All values below are the approved formatting.

## Global table conventions (all 4 tables)

| Attribute | Value |
|---|---|
| Font (latin) | Arial, sz=600 (6 pt) |
| Font (East Asian, CN runs) | SimSun-ExtB, pitchFamily=49, charset=-122 |
| EN paragraph + CN paragraph below it in the same cell (Bucket D) | same font color in both |
| Border weight | outer 12700, inner 6350, black, solid, round caps |
| Cell margins | marL/marR/marT/marB = 0, anchor=ctr |
| Row height | auto-fit to the row's actual bilingual content |
| Paragraph props | algn=ctr, lnSpc 100%, spcBef/spcAft 0, buNone |

## Palette

| Color | Hex | Used for |
|---|---|---|
| Yellow | `FFFF00` | Overview day labels / WEEKLY / BLOCK TIME / FLIGHT TIME rows |
| Green | `DAF2D0` | SN1003 column + 2026 highlight cells |
| Green header | `B5E6A2` | "Flights" header band |
| Blue header | `DAE9F8` | "Flight hours" + Progress Summary headers |
| Light blue | `CAEDFB` | SN1004 column header |
| Data blue | `C0E6F5` | SN1004 data cells |
| Grey | `F2F2F2` | totals / SN / year label cells |
| White | `FFFFFF` | data cells (no explicit fill where template omitted it) |

## Table 1 — Weekly Overview (7–8 rows x 3 cols), left column of slide

- Position: x=0.01in, y=1.78in, w=4.07in (older weeks 0.26/1.96/3.72 — take from source screenshot layout)
- Col widths: [0.966, 1.453, 1.648] in (scale proportionally to the week's screenshot)
- Structure: r0 merged date header ("Overview DD/MM/YY 概览 日期"), r1 SN1003/SN1004 headers, r2+ MONDAY..FRIDAY rows
- Fills: col0 yellow, col1 DAF2D0, col2 C0E6F5 (header row CAEDFB)

## Table 2 — Flight Hours & Flights (5 x 9), middle of slide

- Position: x=4.23in, y=1.78in, w=4.07in
- Col widths: [0.324, 0.468 x8] in (SN col narrow, 8 equal data cols; last col absorbs remainder)
- Row heights: header ~422136 emu, data rows ~324923 emu
- r0: "Flight hours 飞行小时" (DAE9F8) merged c0-c4 + "Flights 次数" (B5E6A2) merged c5-c8
- r1 labels: SN/2024/2025/2026/Total 累计 (F2F2F2; 2026 cols DAF2D0)
- Data rows 1002/1003/1004: SN cell F2F2F2 bold, values white bg (2026 value cells DAF2D0)

## Table 3 — Progress Summary (5 x 4 or 5 x 15 merged), right of slide

- Position: x=8.36in, y=1.78in, w=4.29in
- 4 logical columns: Total scheduled flight hours 计划飞行总小时数 | Hours flown so far 已完成飞行小时数 | Percentage completed 完成百分比 | Hours to go (some need to be repeated) 剩余小时数（部分需重复）
- Headers DAE9F8; r2 values (500:00 / flown / % / open) white bold; r3 second header band (FRQ block) DAE9F8; r4 values white bold
- Header cells are two-line (EN line + CN line), values EN-only

## Table 4 — Weekly Block/Flight Time (3 x 3), bottom-left under Overview

- Position: x=0.02in, y=4.36in, w=4.06in
- Col widths: [1.667, 1.219, 1.171] in
- r0 WEEKLY 周度 (yellow) | SN1003 (DAF2D0) | SN1004 (CAEDFB)
- r1 BLOCK TIME Block时间（滑出->滑入） (yellow) + values (DAF2D0 / C0E6F5)
- r2 FLIGHT TIME 飞行时间（起飞->着陆） (yellow) + values (DAF2D0 / C0E6F5)

## Cell-content style

- Label/header cells: bold, centered
- Day-activity cells: EN line + CN line below, centered, regular weight
- Callout text boxes on the slide (Rectangle 6/3 style): EN paragraphs stay, append one CN paragraph sz=1050 centered
- "Flight Test Status" banner: inline CN run " 飞行测试状态" sz=1400 color 002060 (append BEFORE `endParaRPr`)
