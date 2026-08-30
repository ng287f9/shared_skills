#!/usr/bin/env python3
"""Send Telegram notifications via a fixed bot.

One-off send:
    python send_telegram.py -m "hello"

Watch mode (monitor a progress log, send summary every N minutes):
    python send_telegram.py --watch path/to/log.log --interval 30
"""

import argparse
import os
import re
import sys
import time
from datetime import datetime
from urllib import request as urlrequest
from urllib.parse import urlencode

DEFAULT_TOKEN = "8636941711:AAHOTmSyY-Qa34-y5yS3pXwpkHL9RSjtt6g"
DEFAULT_CHAT_ID = "5886956106"
API_BASE = "https://api.telegram.org"


def send_message(text: str, chat_id: str = DEFAULT_CHAT_ID,
                 token: str = DEFAULT_TOKEN, silent: bool = False,
                 parse_mode: str | None = None) -> bool:
    payload = {
        "chat_id": chat_id,
        "text": text,
        "disable_notification": "true" if silent else "false",
    }
    if parse_mode:
        payload["parse_mode"] = parse_mode

    # Use system proxy if set (api.telegram.org is blocked in some regions)
    opener = urlrequest.build_opener(urlrequest.ProxyHandler())

    data = urlencode(payload).encode("utf-8")
    req = urlrequest.Request(
        f"{API_BASE}/bot{token}/sendMessage",
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with opener.open(req, timeout=30) as resp:
            body = resp.read().decode("utf-8", errors="replace")
            if '"ok":true' in body:
                print("Message sent")
                return True
            print(f"Telegram API error: {body}", file=sys.stderr)
            return False
    except Exception as exc:  # noqa: BLE001
        print(f"Failed to send message: {exc}", file=sys.stderr)
        return False


STAT_PATTERNS = [
    re.compile(r"(?P<num>\d+)\s*/\s*(?P<den>\d+)\b"),
    re.compile(r"(?P<pct>\d+(?:\.\d+)?)\s*%"),
]


def summarize_log(log_path: str) -> str:
    """Extract cumulative progress stats from the WHOLE log file."""
    done = failed = total_pairs = 0
    last_pct = 0.0
    errors: list[str] = []
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as fh:
            for line in fh:
                line = line.strip()
                low = line.lower()
                for m in STAT_PATTERNS[0].finditer(line):
                    den = int(m.group("den"))
                    if den > 0:
                        num = int(m.group("num"))
                        total_pairs = max(total_pairs, den)
                for m in STAT_PATTERNS[1].finditer(line):
                    last_pct = max(last_pct, float(m.group("pct")))
                if re.search(r"\b(done|completed|finished|成功|完成)\b", low):
                    done += 1
                if re.search(r"\b(fail(?:ed)?|error|exception|失败|错误)\b", low):
                    failed += 1
                    if len(errors) < 5:
                        errors.append(line[:120])
    except FileNotFoundError:
        return f"日志文件不存在: {log_path}"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [f"[进度汇报] {now}", f"日志: {os.path.basename(log_path)}"]
    if total_pairs:
        lines.append(f"进度: {total_pairs} 项中已完成计数值见最新记录 ({last_pct:.0f}%)" if last_pct else f"计数: 最大编号到 {total_pairs}")
    elif last_pct:
        lines.append(f"进度: {last_pct:.0f}%")
    if done or failed:
        lines.append(f"完成行: {done} / 失败行: {failed}")
    if errors:
        lines.append("最近失败样例:")
        lines.extend(f"  {e}" for e in errors[-3:])
    if len(lines) == 3 and not (total_pairs or last_pct or done or failed):
        lines.append("(尚未解析到进度信息)")
    return "\n".join(lines)


def tail_log(log_path: str, n: int = 5) -> str:
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as fh:
            return "".join(fh.readlines()[-n:])[-500:]
    except OSError as exc:
        return str(exc)


def watch_mode(log_path: str, interval_min: float, chat_id: str,
               token: str, silent: bool) -> None:
    interval_sec = interval_min * 60
    print(f"Watching {log_path} every {interval_min} min. Ctrl+C to stop.")
    while True:
        summary = summarize_log(log_path)
        tail = tail_log(log_path)
        text = f"{summary}\n\n--- 日志末尾 ---\n{tail}"
        send_message(text, chat_id, token, silent)
        try:
            time.sleep(interval_sec)
        except KeyboardInterrupt:
            send_message("[进度汇报] 监控已停止。\n\n" + summarize_log(log_path),
                         chat_id, token, silent=True)
            break


def main() -> int:
    parser = argparse.ArgumentParser(description="Send Telegram notifications")
    parser.add_argument("-m", "--message", help="Message text to send")
    parser.add_argument("--watch", metavar="LOGFILE",
                        help="Watch a progress log file and report periodically")
    parser.add_argument("--interval", type=float, default=30,
                        help="Watch interval in minutes (default 30)")
    parser.add_argument("--tail", type=int, default=5,
                        help="Lines of log tail to include in watch reports")
    parser.add_argument("--chat-id", default=DEFAULT_CHAT_ID)
    parser.add_argument("--token", default=DEFAULT_TOKEN)
    parser.add_argument("-s", "--silent", action="store_true",
                        help="Send without notification sound")
    parser.add_argument("--parse-mode", default=None,
                        help="Telegram parse_mode (MarkdownV2/HTML); default plain text")
    args = parser.parse_args()

    if args.watch:
        watch_mode(args.watch, args.interval, args.chat_id, args.token, args.silent)
        return 0
    if not args.message:
        parser.error("either -m/--message or --watch is required")
    return 0 if send_message(args.message, args.chat_id, args.token,
                             args.silent, args.parse_mode) else 1


if __name__ == "__main__":
    sys.exit(main())
