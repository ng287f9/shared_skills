---
name: FT_Plan_Update
description: "Use this skill whenever the user asks to update or regenerate the flight test plan from the MPP Gantt chart and bilingual task Excel workbooks — e.g. '更新飞行测试计划', '从mpp更新计划', '更新甘特图到Excel', 'update FT plan from MPP'. The skill: (1) builds an EN→CN terminology glossary, (2) reads MPP Gantt data and builds a formatted Excel matching MPP appearance, and (3) translates English task names to Chinese using the glossary (with built-in aviation dictionary fallback)."
---

# FT Plan Update — MPP → Excel (MPP formatting, Chinese task names)

Reads MS Project `.mpp` Gantt chart data via MPXJ/JPype, builds a 17-column Excel matching MPP table formatting, and translates English task names to Chinese using the terminology glossary.

## Workflow overview

The skill runs three phases in sequence:

1. **Build glossary** (`scripts/build_glossary.py`) — reads the two terminology Excel files (术语对照_中文.xlsx and 术语对照_英文.xlsx), pairs Column C (Task Name) from each row-by-row, and writes a unified EN→CN glossary to `references/术语对照.md`.

2. **MPP → Excel** (`scripts/read_mpp_to_excel.py`) — reads the `.mpp` file via `mpxj` (Java 17+ required), extracts ALL 17 columns from the MPP Entry table (ID, % Complete, Task Name, Text1-Number2 custom fields, Duration, Start, Finish, Predecessors, etc.), extracts per-cell formatting from `GanttChartView.ColumnFontStyle` (background colors, font colors, bold), builds a new `.xlsx` with matching MPP appearance. Column headers are read from `术语对照_中文.xlsx` (row 1 only). **All data comes from MPP — the Chinese Excel is used only for column header names.**

3. **Translate** (`scripts/translate_excel.py`) — reads the Phase 2 output, looks up each English task name in the glossary (exact → fuzzy → FRQ-stripped → keyword match → built-in aviation dictionary fallback), and **replaces** English names with Chinese equivalents. Names without a translation are kept in English. Output is **Chinese-only** (not bilingual).

### Output files

| File | Description |
|---|---|
| `FT_Plan_Update_<mpp_stem>_<date>.xlsx` | Final Excel (17 columns, MPP formatting, Chinese task names) |
| `references/术语对照.md` | Updated EN→CN glossary |

## Input files (expected in the working directory)

```
飞行测试计划/
├── 术语对照_中文.xlsx              ← Column headers reference (row 1)
├── 术语对照_英文.xlsx              ← English task names for glossary
├── FlightTestSchedule_*.mpp        ← MS Project Gantt chart
```

## Phase 1 — Build glossary

**Script:** `scripts/build_glossary.py`

Reads both Excel files, pairs them row-by-row (2..177), extracts column C values, and writes a markdown glossary. Terms are deduplicated.

## Phase 2 — MPP → Excel with MPP formatting

**Script:** `scripts/read_mpp_to_excel.py`

### Prerequisites
- **Java 17+** (for mpxj/JPype bridge)
- `mpxj` Python package
- `openpyxl` Python package

### Data extraction

Reads ALL 17 columns from the MPP Entry table:

| Excel Col | Column Header | MPP Source |
|---|---|---|
| 1 | 任务号 | ID |
| 2 | 完成百分比 | % Complete (0-1 decimal) |
| 3 | 任务名称 | Task Name (indented by outline level) |
| 4 | 计划飞机 | Text1 (Planned A/C) |
| 5 | 需要的架次 | Number1 (Required Sortie#) |
| 6 | 需要的飞行时间 | Number2 (Required FH) |
| 7 | 前提条件 | Text16 (Prerequisite) |
| 8 | 备注 | Text12 (RMKS) |
| 9 | 文件状态 | Text26 (Paperwork status) |
| 10 | 工期 | Duration (formatted as "X 天/周/小时") |
| 11 | 开始时间 | Start ("YYYY年M月D日") |
| 12 | 完成时间 | Finish |
| 13 | 前置任务 | Predecessors (comma-separated UIDs) |
| 14 | 使用特定飞机原因 | Text10 |
| 15 | 额外需求 | Text18 |
| 16 | 实际开始时间 | Actual Start |
| 17 | 实际完成时间 | Actual Finish |

### Formatting (from MPP GanttChartView)

- **Default font**: Calibri 11pt, regular weight
- **Summary tasks** (outline level 1): bold
- **% Complete column**: Calibri 10pt
- **Headers**: Calibri 10pt, grey background (#DADCDD)
- **Per-cell background/font colors**: extracted from `ColumnFontStyle` entries in the GanttChartView
- **Grid lines**: thin, colour #DADCDD (matching MPP SheetRowsGridLines)
- **Task indentation**: 2 spaces per outline level below level 1

## Phase 3 — Translate to Chinese

**Script:** `scripts/translate_excel.py`

Replaces English task names in Column 3 with Chinese translations.

### Translation order

1. **Exact match** — look up in glossary
2. **Case-insensitive exact**
3. **FRQ-stripped match** (strip FRQ-XXX-XXX.XXX-XXX prefix)
4. **Fuzzy match** (difflib, cutoff 0.85)
5. **Keyword/Jaccard match** (significant word overlap, cutoff 0.4)
6. **Built-in aviation dictionary** (200+ terms covering avionics, structures, flight controls, systems)
7. **Keep English** — if no match found

### Output format

- **Cell shows Chinese text only** (not bilingual)
- Cell formatting (font, fill, alignment) is preserved from Phase 2
- Rows already in Chinese are skipped

## Running the full workflow

```bash
cd "D:/.../飞行测试计划"
python scripts/run_all.py
```

Or run phases individually:

```bash
# Phase 1
python scripts/build_glossary.py \
  --cn 术语对照_中文.xlsx \
  --en 术语对照_英文.xlsx \
  --output references/术语对照.md

# Phase 2
python scripts/read_mpp_to_excel.py \
  --mpp FlightTestSchedule_*.mpp \
  --template 术语对照_中文.xlsx \
  --output "FT_Plan_Update_<date>.xlsx"

# Phase 3
python scripts/translate_excel.py \
  --input "FT_Plan_Update_<date>.xlsx" \
  --glossary references/术语对照.md \
  --output "FT_Plan_Update_<date>.xlsx"
```

## Notes

- **The original MPP and Excel files are never modified.**
- Java 17+ is required. Install from https://adoptium.net if needed.
- MS Project does not need to be installed — `mpxj` handles the binary format.
- The MPP file must have 176 data rows (excluding the root project task).
- All 17 columns of data come from the MPP Entry table. `术语对照_中文.xlsx` is used ONLY for its column header names (row 1).
