---
name: telegram_notifier
description: Use whenever the user asks to send notifications or progress reports via Telegram (发Telegram通知、Telegram进度、telegram notify、send telegram message、telegram progress report). Sends messages from a fixed bot to Chat ID 5886956106. Supports one-off sends, silent sends, and a watch mode that monitors a progress log file and sends a summary every N minutes.
---

# Telegram Notifier Skill

Send Telegram notifications via a pre-configured bot.

## Fixed configuration

- **Bot Token**: `8636941711:AAHOTmSyY-Qa34-y5yS3pXwpkHL9RSjtt6g`
- **Chat ID**: `5886956106`
- **Script**: `scripts/send_telegram.py` (relative to this SKILL.md's directory)

## Usage

### One-off send

```powershell
python "<skill_dir>/scripts/send_telegram.py" -m "任务完成：批处理已结束，共处理 128 个文件"
```

Options:

| Flag | Description |
| ---- | ----------- |
| `-m, --message` | Message text (required unless `--watch`) |
| `--chat-id` | Override chat ID (default 5886956106) |
| `--token` | Override bot token |
| `-s, --silent` | Send silently (no notification sound) |

### Watch mode (progress log monitoring)

Monitor a progress log file and send a summary every N minutes (default 30):

```powershell
python "<skill_dir>/scripts/send_telegram.py" --watch "F:\logs\job.log" --interval 30
```

- Each cycle it reads the WHOLE log and extracts progress stats (percentages,
  counts like `12/100`, `done`, `failed`, `error` lines), so multiple workers
  writing the same log get merged into one cumulative summary.
- Stop with Ctrl+C (a final summary is sent on exit).

### Getting the skill directory

Resolve at runtime with:

```powershell
Split-Path (Get-ChildItem "$env:USERPROFILE\.config\opencode\skills\telegram_notifier\SKILL.md").FullName
```

Or simply use the absolute path:
`C:\Users\glenn\.config\opencode\skills\telegram_notifier\scripts\send_telegram.py`

## Notes

- Script uses only the Python standard library (`urllib.request`) — no
  third-party dependencies required.
- 脚本自动使用系统代理（读取 `HTTPS_PROXY`/`HTTP_PROXY` 环境变量，本机为
  `http://127.0.0.1:10809`）。api.telegram.org 在直连下会被墙，必须确保代理
  环境变量已设置且代理软件正在运行。
- 发送失败时检查：代理是否开启、token 是否有效（401）、用户是否拉黑了 bot（403）。
- Messages support Telegram MarkdownV2 when passed with `--parse-mode MarkdownV2`,
  but plain text is the safe default (no escaping needed).

## Workflow

1. Confirm what message/content to send (or which log file to watch).
2. Run the script as shown above.
3. Check exit output for `Message sent` confirmation; report failures verbatim
   (e.g. 401 = bad token, 403 = bot blocked by user).
