#!/usr/bin/env python3
r"""OCR 风格 Telegram 进度报告（可复用模板）。

发送形如以下的 MarkdownV2 进度卡片：
  📊 *<title>*（<worker_desc>）
  ━━━━━━━━━━━━━━━━━━━
  📁 待处理总数：*N* 个
  ✅ 已完成：*d* 个 (p%)
  ⏳ 剩余：*r* 个

  *👷 工人状态*
   ⚙️ <worker>: d/N 运行中
   ├ 🚀 <done_label>：d 个
   ├ ⏭️ 跳过：s 个
   └ ❌ 出错：e 个

  ⚡️ 处理速率：*rate* 个/小时
  🕐 预计剩余：*eta*
  🏁 预计完成：*eta_time*

  [████────] p%

  🕐 汇报时间：*now*

用法：
  python progress_report.py --total 16614 --done 561 \
      --title "摘要生成进度报告" --worker build_summary --secs-per-doc 1.9
  # 或直接从 JSONL 统计已完成数：
  python progress_report.py --total 16614 --log F:\abbotspace\summary_new.jsonl \
      --title "摘要生成进度报告" --worker build_summary --rate 1894
"""
import argparse
import os
import subprocess
import sys
import time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
SEND = os.path.join(HERE, "send_telegram.py")
DEFAULT_CHAT = "5886956106"


def mdv2_escape(s: str) -> str:
    """MarkdownV2 转义：保留 * 作为粗体定界符。"""
    out = []
    for ch in str(s):
        if ch in r'_*[]()~`>#+-=|{}.!' and ch != '*':
            out.append('\\' + ch)
        else:
            out.append(ch)
    return ''.join(out)


def bar(pct: float, width: int = 20) -> str:
    f = int(round(pct / 100 * width))
    return '█' * f + '─' * (width - f)


def build_report(total, done, title, worker, worker_desc, rate, skip, error,
                 done_label):
    now = datetime.now()
    now_s = now.strftime('%Y-%m-%d %H:%M:%S')
    done = min(done, total)
    pct = 100.0 * done / total if total else 0.0
    remain = total - done

    if done >= total:
        head = '🎉 *全部文件已完成！*'
    else:
        head = '📊 *%s*（%s）' % (title, worker_desc)

    if rate and rate > 0:
        eta_min = remain / rate * 60.0
        eta_t = now.timestamp() + eta_min * 60
        eta_s = datetime.fromtimestamp(eta_t).strftime('%m-%d %H:%M')
        eta_m = ('%d 分钟' % int(eta_min)) if eta_min < 600 else \
                ('%.1f 小时' % (eta_min / 60))
    else:
        eta_s, eta_m = '—', '—'

    lines = [
        head,
        '━━━━━━━━━━━━━━━━━━',
        '📁 待处理总数：*%d* 个' % total,
        '✅ 已完成：*%d* 个 (%.1f%%)' % (done, pct),
        '⏳ 剩余：*%d* 个' % remain,
        '',
        '*👷 工人状态*',
        ' ⚙️ %s: %d/%d %s' % (worker, done, total,
                               '✅已完成' if done >= total else '运行中'),
        ' ├ 🚀 %s：%d 个' % (done_label, done),
        ' ├ ⏭️ 跳过：%d 个' % skip,
        ' └ ❌ 出错：%d 个' % error,
        '',
        '⚡️ 处理速率：*%s* 个/小时' % ('%d' % int(rate) if rate else '—'),
        '🕐 预计剩余：*%s*' % eta_m,
        '🏁 预计完成：*%s*' % eta_s,
        '',
        '[%s] %.1f%%' % (bar(pct), pct),
        '',
        '🕐 汇报时间：*%s*' % now_s,
    ]
    return '\n'.join(lines)


def send(msg: str, chat_id: str, token: str, silent: bool):
    env = os.environ.copy()
    env['PYTHONUTF8'] = '1'
    env['HTTP_PROXY'] = env.get('HTTP_PROXY', 'http://127.0.0.1:10809')
    env['HTTPS_PROXY'] = env.get('HTTPS_PROXY', 'http://127.0.0.1:10809')
    cmd = [sys.executable, SEND, '-m', msg, '--chat-id', chat_id]
    if token:
        cmd += ['--token', token]
    if silent:
        cmd += ['-s']
    subprocess.run(cmd, env=env, check=False,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def main():
    ap = argparse.ArgumentParser(description="OCR 风格 Telegram 进度报告")
    ap.add_argument('--total', type=int, required=True)
    ap.add_argument('--done', type=int, default=None, help="已完成数（与 --log 二选一）")
    ap.add_argument('--log', help="按行统计已完成数的文件（每行一条 JSON）")
    ap.add_argument('--title', default="批处理进度报告")
    ap.add_argument('--worker', default="worker")
    ap.add_argument('--worker-desc', default="单工人")
    ap.add_argument('--rate', type=float, default=0, help="速率（个/小时）")
    ap.add_argument('--secs-per-doc', type=float, default=0,
                   help="单篇耗时（秒），用于推算速率")
    ap.add_argument('--skip', type=int, default=0)
    ap.add_argument('--error', type=int, default=0)
    ap.add_argument('--done-label', default="已生成摘要")
    ap.add_argument('--chat-id', default=DEFAULT_CHAT)
    ap.add_argument('--token', default=None)
    ap.add_argument('-s', '--silent', action='store_true')
    args = ap.parse_args()

    if args.done is None:
        if args.log and os.path.exists(args.log):
            with open(args.log, encoding='utf-8') as f:
                args.done = sum(1 for line in f if line.strip())
        else:
            args.done = 0

    rate = args.rate
    if not rate and args.secs_per_doc:
        rate = 3600.0 / args.secs_per_doc

    msg = build_report(args.total, args.done, args.title, args.worker,
                       args.worker_desc, rate, args.skip, args.error,
                       args.done_label)
    send(msg, args.chat_id, args.token, args.silent)
    print(msg)


if __name__ == '__main__':
    main()
