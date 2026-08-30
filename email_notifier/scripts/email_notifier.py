"""email_notifier: Gmail SMTP 发送邮件 / 监控进度日志定时汇报

用法:
    python email_notifier.py --subject "标题" --body "正文"           # 单次发送
    python email_notifier.py --watch <进度日志> [--interval 30]       # 每30分钟发进度邮件

进度日志格式: 每行一个 JSON, 支持 {"cum": {...}} 累计统计记录;
              watch 模式记录已读行号, 只发增量 + 累计摘要。
"""
import argparse
import json
import smtplib
import time
from datetime import datetime
from email.mime.text import MIMEText
from email.header import Header
from pathlib import Path

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
FROM_ADDR = "glennchou2025@gmail.com"
TO_ADDR = "glenn.chou@outlook.com"
APP_PASS = "qmkg lyte eoni heso".replace(" ", "")


def send_email(subject: str, body: str) -> None:
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = FROM_ADDR
    msg["To"] = TO_ADDR
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(FROM_ADDR, APP_PASS)
        server.sendmail(FROM_ADDR, [TO_ADDR], msg.as_string())
    print(f"[OK] {datetime.now().strftime('%H:%M:%S')} 邮件已发送: {subject}")


def parse_log(log_path: Path, start_line: int) -> tuple[int, list, dict, int]:
    """返回 (新行数, 新增记录, 全日志合并累计统计, 总行数)

    累计统计直接从全日志所有文件记录重算, 支持多工人写同一日志的合并汇总。
    """
    lines = log_path.read_text(encoding="utf-8", errors="ignore").splitlines()
    total = len(lines)
    new = lines[start_line:]
    recs = []
    stats = {"done": 0, "wm_removed": 0, "wm_none": 0, "ocr": 0, "skip_ocr": 0, "fail": 0}
    for ln in lines:
        try:
            d = json.loads(ln)
        except json.JSONDecodeError:
            continue
        if "cum" in d or "file" not in d:
            continue
        stats["done"] += 1
        if d.get("err"):
            stats["fail"] += 1
        elif d.get("wm") == "none":
            stats["wm_none"] += 1
        elif d.get("wm"):
            stats["wm_removed"] += 1
        if d.get("ocr") == "done":
            stats["ocr"] += 1
        elif d.get("ocr") == "skip(text-layer)":
            stats["skip_ocr"] += 1
    # total: 取所有记录里最大 total 字段; 兼容 worker 记录 (有 total)
    totals = [d["total"] for ln in lines for d in [safe_load(ln)] if d and "file" in d and "total" in d]
    stats["total"] = max(totals) if totals else stats["done"]
    new_recs = []
    for ln in new:
        try:
            d = json.loads(ln)
        except json.JSONDecodeError:
            continue
        if "file" in d:
            new_recs.append(d)
    return len(new_recs), new_recs, stats, total


def safe_load(ln):
    try:
        return json.loads(ln)
    except json.JSONDecodeError:
        return None


def build_body(new_recs: list, stats: dict, since: int) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [f"批次处理进度汇报(双工人)", f"时间: {now}", ""]
    c = stats
    if c.get("total"):
        lines += [
            f"累计进度: {c.get('done', 0)} / {c.get('total', '?')}",
            f"去水印: {c.get('wm_removed', 0)} | 无水印: {c.get('wm_none', 0)}",
            f"OCR(扫描版): {c.get('ocr', 0)} | 跳过OCR(有文字层): {c.get('skip_ocr', 0)}",
            f"失败: {c.get('fail', 0)} | 剩余: {c.get('total', 0) - c.get('done', 0)}",
        ]
    else:
        lines.append("累计统计暂不可用。")
    if new_recs:
        errs = [r for r in new_recs if r.get("err")]
        lines.append(f"\n本周期({since}个新文件): 失败 {len(errs)} 个")
        for r in errs[:5]:
            lines.append(f"  - {Path(r['file']).name}: {r['err']}")
    else:
        lines.append("\n本周期暂无新进度。")
    return "\n".join(lines)


def watch(log_path: Path, interval: int, subject_prefix: str) -> None:
    start_line = 0
    print(f"[WATCH] 监控 {log_path} 每 {interval} 分钟发送进度邮件")
    while True:
        try:
            n, recs, cum, total = parse_log(log_path, start_line)
            body = build_body(recs, cum, n)
            send_email(f"{subject_prefix}进度 [{datetime.now().strftime('%m-%d %H:%M')}]", body)
            start_line = total
        except Exception as e:
            print(f"[ERROR] {e}")
        time.sleep(interval * 60)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--subject", default="批处理进度")
    ap.add_argument("--body", default="")
    ap.add_argument("--watch", default=None, help="监控的进度日志路径")
    ap.add_argument("--interval", type=int, default=30, help="发送间隔(分钟)")
    args = ap.parse_args()

    if args.watch:
        watch(Path(args.watch), args.interval, args.subject)
    else:
        send_email(args.subject, args.body or "无正文。")


if __name__ == "__main__":
    main()