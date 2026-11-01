"""用 sellout_daily.csv 產生週報（Markdown），給 weekly-report skill 使用

輸入：data/processed/sellout_daily.csv
輸出：output/週報_<本週最後一天>.md（含銷售資料，不放進 Git）

規則：
- 一週＝週一到週日。各通路分開取「最近一個完整週」當本週，前一週當上週
  （兩通路資料期間沒有重疊，所以不能用同一週比較）。
- 康是美最後一天（報表資料最新日期，資料不完整）排除，和 Dashboard 一致。
- 品項：有條碼用條碼，沒有用通路品號；依「金額增減」排前 5 名成長、前 5 名衰退。
- 成長率＝（本週－上週）÷ 上週；上週為 0 或負數時不計算。
- 「重點結論」留一行標記 <!-- 重點結論 -->，由 skill 填入 3 句結論。

執行方式：
  python scripts\\weekly_report.py
"""

import argparse
import csv
import sys
from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CHANNELS = ['屈臣氏', '康是美']
PLACEHOLDER = '<!-- 重點結論 -->'
TOP_N = 5
NO_NAME = '（原始檔無品名）'


def load(csv_path):
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        rows = list(csv.DictReader(f))
    last_cosmed = max((r['日期'] for r in rows if r['通路'] == '康是美'), default=None)
    rows = [r for r in rows if not (r['通路'] == '康是美' and r['日期'] == last_cosmed)]
    return rows, last_cosmed


def latest_full_week(days):
    """資料範圍內最後一個完整的週一～週日；不足兩週時回傳 None。"""
    first, last = min(days), max(days)
    end = last - timedelta(days=(last.weekday() + 1) % 7)   # 最近的週日（含當天）
    start = end - timedelta(days=6)
    if start - timedelta(days=7) < first:
        return None
    return start, end


def growth(cur, prev):
    if prev <= 0:
        return None
    return (cur - prev) / prev


def pct(x):
    return '—' if x is None else f'{x:+.1%}'


def money(x):
    return f'{x:,.0f}'


def signed(x):
    return f'{x:+,.0f}'


def summarize(rows):
    by_ch = defaultdict(list)
    for r in rows:
        by_ch[r['通路']].append(r)

    result = []
    for ch in CHANNELS:
        ch_rows = by_ch.get(ch, [])
        if not ch_rows:
            continue
        days = {date.fromisoformat(r['日期']) for r in ch_rows}
        week = latest_full_week(days)
        if not week:
            result.append({'channel': ch, 'week': None})
            continue
        start, end = week
        prev_start, prev_end = start - timedelta(days=7), end - timedelta(days=7)

        tot = {'cur': [Decimal(0), Decimal(0)], 'prev': [Decimal(0), Decimal(0)]}
        items = defaultdict(lambda: {'name': '', 'cur': Decimal(0), 'prev': Decimal(0),
                                     'cur_qty': Decimal(0), 'prev_qty': Decimal(0)})
        for r in ch_rows:
            d = date.fromisoformat(r['日期'])
            if start <= d <= end:
                k = 'cur'
            elif prev_start <= d <= prev_end:
                k = 'prev'
            else:
                continue
            qty, amt = Decimal(r['銷售數量'] or 0), Decimal(r['銷售金額'] or 0)
            tot[k][0] += qty
            tot[k][1] += amt
            it = items[r['條碼'] or r['通路品號']]
            if not it['name']:
                it['name'] = r['品名'] if r['品名'] != NO_NAME else f"{NO_NAME[:-1]}，品號 {r['通路品號']}）"
            it[k] += amt
            it[k + '_qty'] += qty

        ranked = sorted(items.values(), key=lambda i: i['cur'] - i['prev'])
        result.append({
            'channel': ch, 'week': (start, end), 'prev_week': (prev_start, prev_end),
            'qty': (tot['cur'][0], tot['prev'][0]), 'amt': (tot['cur'][1], tot['prev'][1]),
            'up': [i for i in reversed(ranked) if i['cur'] > i['prev']][:TOP_N],
            'down': [i for i in ranked if i['cur'] < i['prev']][:TOP_N],
            'items': len(items),
        })
    return result


def fmt_range(w):
    return f'{w[0]:%m/%d}～{w[1]:%m/%d}'


def render(result, last_cosmed):
    ok = [r for r in result if r['week']]
    latest_end = max(r['week'][1] for r in ok)
    lines = [f'# Sell-out 週報（{latest_end:%Y-%m-%d}）', '', '## 重點結論', '', PLACEHOLDER, '',
             '## 資料期間', '',
             '| 通路 | 本週 | 上週 |', '|---|---|---|']
    for r in ok:
        lines.append(f"| {r['channel']} | {r['week'][0]:%Y-%m-%d}～{r['week'][1]:%Y-%m-%d} "
                     f"| {r['prev_week'][0]:%Y-%m-%d}～{r['prev_week'][1]:%Y-%m-%d} |")
    lines.append('')
    notes = ['一週為週一到週日，各通路取資料中最近一個完整週。']
    if len({r['week'] for r in ok}) > 1:
        notes.append('各通路的本週不是同一週（資料期間沒有重疊），通路之間只能比較各自的成長率，不加總。')
    if last_cosmed:
        notes.append(f'康是美 {last_cosmed} 是報表資料最新日期，當天資料不完整，已排除。')
    for r in result:
        if not r['week']:
            notes.append(f"{r['channel']} 資料不足兩個完整週，未列入。")
    lines += [f'- {n}' for n in notes] + ['']

    lines += ['## 各通路本週 vs 上週', '',
              '| 通路 | 本週數量 | 上週數量 | 數量成長率 | 本週金額 | 上週金額 | 金額成長率 |',
              '|---|---:|---:|---:|---:|---:|---:|']
    for r in ok:
        (cq, pq), (ca, pa) = r['qty'], r['amt']
        lines.append(f"| {r['channel']} | {money(cq)} | {money(pq)} | {pct(growth(cq, pq))} "
                     f"| {money(ca)} | {money(pa)} | {pct(growth(ca, pa))} |")
    if len(ok) > 1 and len({r['week'] for r in ok}) == 1:
        cq, pq = sum(r['qty'][0] for r in ok), sum(r['qty'][1] for r in ok)
        ca, pa = sum(r['amt'][0] for r in ok), sum(r['amt'][1] for r in ok)
        lines.append(f"| **合計** | {money(cq)} | {money(pq)} | {pct(growth(cq, pq))} "
                     f"| {money(ca)} | {money(pa)} | {pct(growth(ca, pa))} |")
    lines.append('')

    head = ['| # | 品名 | 上週金額 | 本週金額 | 金額增減 | 成長率 | 本週數量 |',
            '|---:|---|---:|---:|---:|---:|---:|']
    for r in ok:
        lines += [f"## {r['channel']}：品項變化（{fmt_range(r['week'])} vs {fmt_range(r['prev_week'])}）", '']
        for title, key in (('成長最多（依金額增減）', 'up'), ('衰退最多（依金額增減）', 'down')):
            lines += [f'### {title}', '']
            if not r[key]:
                lines += ['（無）', '']
                continue
            lines += head
            for n, i in enumerate(r[key], 1):
                lines.append(f"| {n} | {i['name']} | {money(i['prev'])} | {money(i['cur'])} "
                             f"| {signed(i['cur'] - i['prev'])} | {pct(growth(i['cur'], i['prev']))} "
                             f"| {money(i['cur_qty'])} |")
            lines.append('')
    lines += ['---', '', '成長率「—」表示上週金額為 0 或負數（新上架或退貨），無法計算。', '']
    return '\n'.join(lines), latest_end


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description='產生 sell-out 週報')
    ap.add_argument('--csv', default=ROOT / 'data' / 'processed' / 'sellout_daily.csv', type=Path)
    ap.add_argument('--out-dir', default=ROOT / 'output', type=Path)
    args = ap.parse_args()

    rows, last_cosmed = load(args.csv)
    result = summarize(rows)
    if not any(r['week'] for r in result):
        sys.exit('沒有任何通路有兩個完整週的資料，無法產生週報。')
    text, latest_end = render(result, last_cosmed)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    out = args.out_dir / f'週報_{latest_end:%Y-%m-%d}.md'
    out.write_text(text, encoding='utf-8')
    print(f'週報：{out}')


if __name__ == '__main__':
    main()
