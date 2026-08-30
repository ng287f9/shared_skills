---
name: email_notifier
description: Use whenever the user asks to send email notifications or progress reports by email (发邮件、进度邮件、邮件通知、定时汇报、email notification、email progress、send progress email). Sends email from glennchou2025@gmail.com (Gmail SMTP + 应用程序专用密码) to glenn.chou@outlook.com. Supports one-off sends and a watch mode that monitors a progress log file and emails a summary every N minutes (default 30). Useful for reporting long-running batch job progress. The watch mode recomputes combined cumulative stats from the WHOLE log, so multiple workers writing the same log get merged into one summary.
---

# email_notifier — 邮件进度通知 (Gmail SMTP)

## 账号与配置

| 项 | 值 |
|----|----|
| 发件人 | `glennchou2025@gmail.com` |
| 收件人 | `glenn.chou@outlook.com` |
| SMTP | `smtp.gmail.com:587` (STARTTLS) |
| 授权方式 | Gmail 应用程序专用密码 `qmkglyteeoniheso` (输入时空格需移除) |
| 依赖 | Python 3 标准库 `smtplib` / `email` (无需额外安装) |

## 环境陷阱 (必须遵守)

1. **专用密码**: 应用中输入的密码形如 `qmkg lyte eoni heso`, 使用前必须 `replace(" ", "")` 去掉空格 (`qmkglyteeoniheso`)。
2. **端口**: 用 587 + STARTTLS (`smtplib.SMTP().starttls()`), 不要用 465。
3. **代理**: smtplib 不走系统 HTTP 代理, 无需处理。
4. **错误码**: 若 Gmail 返回 `534`/`535`, 多为密码错误或未启用两因素认证的专用密码; 若 `550`, 检查收件人地址。
5. **凭据安全**: 本 skill 内含明文应用专用密码, 不要提交到公开仓库/日志。

## 用法

```powershell
# 1. 单次发送
D:\pdf-summary-ai\.venv\Scripts\python.exe <skill>\scripts\email_notifier.py --subject "标题" --body "正文"

# 2. 监控进度日志, 每 30 分钟发一次进度邮件 (后台运行)
D:\pdf-summary-ai\.venv\Scripts\python.exe <skill>\scripts\email_notifier.py --watch <进度日志路径> [--interval 30] [--subject "前缀"]
```

## 脚本

`scripts/email_notifier.py` — 发送与 watch 模式 (读进度日志 → 汇总 → 每 N 分钟发邮件)。

- `--subject` 邮件标题; `--body` 正文 (单次发送用)
- `--watch <log>` 持续监控模式: 解析日志中的 JSON 行, **从全日志重算合并累计统计** (多工人写同一日志自动汇总), 每 `--interval` 分钟发送一次进度邮件
- 日志尾部推送: 记录上次已读行号, 每次只汇总增量与累计
- 无内容时发送"暂无新进度"提示