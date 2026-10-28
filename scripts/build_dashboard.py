"""把 sellout_daily.csv 做成 Dashboard（單一 HTML 檔，資料包在檔案裡，不需要網路）

輸入：data/processed/sellout_daily.csv
樣板：dashboard/template.html
輸出：dashboard/sellout_dashboard.html（含銷售資料，不放進 Git）

執行方式：
  python scripts\\build_dashboard.py
（clean_sellout.py 跑完也會自動執行）
"""

import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CHANNELS = ['屈臣氏', '康是美']   # 顏色依這個順序固定：屈臣氏藍、康是美橘


def is_trial(name):
    """試用品、體驗包（品名以「88-33」或「試)」開頭）。"""
    return name.startswith('88-33') or name.startswith('試)') or name.startswith('試）')


def build(csv_path, template_path, out_path):
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        rows = list(csv.DictReader(f))

    # 康是美報表的「資料最新日期」當天只有部分資料，排除該通路的最後一天
    excluded = []
    last_cosmed = max((r['日期'] for r in rows if r['通路'] == '康是美'), default=None)
    if last_cosmed:
        excluded.append({'channel': '康是美', 'date': last_cosmed, 'reason': '報表資料最新日期，當天資料不完整'})
        rows = [r for r in rows if not (r['通路'] == '康是美' and r['日期'] == last_cosmed)]

    # 商品：有條碼的用條碼合併跨通路的同一商品；沒有條碼的（康是美待補）用通路品號
    products, prod_index = [], {}
    dates = sorted({r['日期'] for r in rows})
    date_index = {d: i for i, d in enumerate(dates)}
    stores = sorted({r['門市'] for r in rows if r['門市']})
    store_index = {s: i for i, s in enumerate(stores)}

    # 品名優先用屈臣氏的（有條碼、寫法較統一）
    ordered = sorted(rows, key=lambda r: CHANNELS.index(r['通路']))
    facts = []
    for r in ordered:
        ch = CHANNELS.index(r['通路'])
        key = r['條碼'] or f"{r['通路']}:{r['通路品號']}"
        if key not in prod_index:
            prod_index[key] = len(products)
            products.append({'name': r['品名'], 'barcode': r['條碼'], 'trial': is_trial(r['品名']),
                             'codes': {}})
        p = products[prod_index[key]]
        p['codes'].setdefault(r['通路'], r['通路品號'])
        if is_trial(r['品名']):
            p['trial'] = True
        facts.append([ch, store_index.get(r['門市'], -1), date_index[r['日期']], prod_index[key],
                      float(r['銷售數量']), float(r['銷售金額'])])
    facts.sort(key=lambda x: (x[2], x[0]))
    facts = [[a, b, c, d, int(q) if q == int(q) else q, int(m) if m == int(m) else m]
             for a, b, c, d, q, m in facts]

    data = {
        'generated': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'channels': CHANNELS,
        'excluded': excluded,
        'dates': dates,
        'stores': stores,
        'products': products,
        # 每筆：[通路, 門市(-1 = 無), 日期, 商品, 數量, 金額]
        'facts': facts,
    }
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')

    html = template_path.read_text(encoding='utf-8')
    marker = '/*__DATA__*/null'
    if marker not in html:
        raise SystemExit(f'錯誤：{template_path} 找不到資料位置標記 {marker}')
    out_path.write_text(html.replace(marker, payload), encoding='utf-8')
    print(f'Dashboard：{out_path}（{len(facts)} 筆、{len(products)} 個商品；'
          f"已排除 {', '.join(e['channel'] + ' ' + e['date'] for e in excluded) or '無'}）")


def main():
    parser = argparse.ArgumentParser(description='產生 sellout Dashboard')
    parser.add_argument('--csv', type=Path, default=ROOT / 'data' / 'processed' / 'sellout_daily.csv')
    parser.add_argument('--out', type=Path, default=ROOT / 'dashboard' / 'sellout_dashboard.html')
    args = parser.parse_args()
    build(args.csv, ROOT / 'dashboard' / 'template.html', args.out)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
