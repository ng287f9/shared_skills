# Output Excel Structure — FT Plan Update

## Column layout (final bilingual output)

| Col | Header (EN) | Header (CN) | Content | Width |
|---|---|---|---|---|
| A | Task Name | 任务名称 | Bilingual EN+CN (newline separated), indented by outline level | 50 |
| B | Duration | 工期 | Task duration as text (e.g. "10 days", "2 wks") | 12 |
| C | Start | 开始时间 | Start date (YYYY-MM-DD) | 16 |
| D | Finish | 完成时间 | Finish date (YYYY-MM-DD) | 16 |
| E | % Complete | 完成百分比 | Percentage complete (e.g. 75%) | 10 |
| F | Predecessors | 前置任务 | Comma-separated predecessor IDs | 12 |
| G | Notes | 备注 | Free-text notes, wrapped | 30 |

## Source terminology Excel structure (术语对照_中文.xlsx / 术语对照_英文.xlsx)

Both files have identical column structure (17 columns, 177 rows). Column C is the key pairing column:

| Col | Chinese header | English header |
|---|---|---|
| A | 计划模式 | 计划模式 |
| B | 完成百分比 | 完成百分比 |
| C | **任务名称** | **Task Name** |
| D | 计划飞机 | Planned A/C |
| E | 需要的架次 | Required Sortie# |
| F | 需要的飞行时间 | Required FH |
| G | 前置条件 | Prerequisite |
| H | 备注 | RMKS |
| I | 文件状态 | Paperwork status |
| J | 工期 | 工期 |
| K | 开始时间 | 开始时间 |
| L | 完成时间 | 完成时间 |
| M | 前置任务 | 前置任务 |
| N | 使用特定飞机原因 | Reason for specific A/C |
| O | 附加要求 | Additional Requirements |
| P | 实际开始时间 | 实际开始时间 |
| Q | 实际完成时间 | 实际完成时间 |

## Row mapping

Rows 2-177 in both Excel files are 1:1 aligned by task ID (column A values match). Never re-sort rows independently in either file.

## MPP hierarchy levels

The MPP outline level maps to Excel indentation:
- Level 1 → bold, no indent (summary / major phase)
- Level 2 → regular, 2-char indent (sub-phase)
- Level 3 → regular, 4-char indent (work package)
- Level 4+ → regular, 6-char indent (individual task)
