"""整理屈臣氏、康是美的 sellout Excel，輸出統一格式的每日銷售 CSV

輸出：
  data/processed/sellout_daily.csv     每日銷售明細（兩個通路合併）
  data/processed/康是美品號對照.csv      康是美品號 → 國際條碼對照表

康是美報表沒有條碼，程式會用品名自動配對屈臣氏的條碼。
配對不到或配錯的，直接在對照表的「條碼」欄填入正確條碼後重新執行即可，手動填寫的內容不會被覆蓋。
原始檔缺品名的品號，也可以在對照表的「品名」欄補上。

執行方式（不需要 Excel，需要 xlrd、openpyxl 套件）：
  python scripts\\clean_sellout.py
"""

import argparse
import csv
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl
import xlrd

ROOT = Path(__file__).resolve().parent.parent
COLUMNS = ['通路', '門市', '日期', '條碼', '通路品號', '品名', '通路分類', '銷售數量', '銷售金額']
MAP_COLUMNS = ['品號', '品名', '條碼', '配對方式']


class DataError(Exception):
    """原始檔內容和預期不同，需要使用者處理。"""


def find_raw_file(raw_dir, pattern):
    files = [f for f in raw_dir.iterdir() if f.is_file() and f.match(pattern)]
    if len(files) != 1:
        names = ', '.join(f.name for f in files)
        raise DataError(f"data/raw 裡應該剛好有一個符合 '{pattern}' 的檔案，目前找到 {len(files)} 個：{names}")
    return files[0]


def to_code(x):
    """條碼、品號等代碼轉成字串；Excel 存成數字的，轉回完整整數字串。"""
    if x is None:
        return ''
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x).strip()


def to_number(x):
    if x is None or x == '':
        return 0.0
    return float(x)


def fmt_number(x):
    """整數不帶小數點（449 而不是 449.0），其他四捨五入到小數 2 位。"""
    x = round2(x)
    if x == int(x):
        return str(int(x))
    return f'{x:.2f}'.rstrip('0')


def normalize_name(s):
    """品名正規化，用來跨通路配對同一商品。

    「88-33」「試)」等前綴代表體驗包、試用品，刻意保留，避免和正式商品配到同一個條碼。
    """
    s = s.upper()
    s = re.sub(r'[（(]', '(', s)
    s = re.sub(r'[）)]', ')', s)
    s = re.sub(r'DR\.?\s*WU', '', s)
    return re.sub(r'\s+', '', s)


def round2(x):
    """四捨五入到小數 2 位。內建 round() 會把 1227.135 捨成 1227.13（浮點誤差），所以改用 Decimal。"""
    return float(Decimal(repr(x)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


def make_row(channel, store, day, barcode, item, name, category, qty, amount):
    return dict(zip(COLUMNS, [channel, store, day, barcode, item, name, category, qty, round2(amount)]))


def read_watsons(raw_dir):
    """屈臣氏：每列一筆 門市 × 商品 × 日期。回傳 (rows, 正規化品名 → 條碼)。"""
    path = find_raw_file(raw_dir, '*明細*.xls*')
    print(f'讀取屈臣氏：{path.name}')
    wb = xlrd.open_workbook(str(path), on_demand=True)
    try:
        # 「DailySaleExcel (2)」包含 5～6 月完整資料，其他工作表是重複或樞紐分析
        ws = wb.sheet_by_name('DailySaleExcel (2)')
        expected = ['貨號', '門市', '條碼', '品名', '銷售數量', '銷售金額', '日期', '部門分類']
        header = [str(ws.cell_value(2, c)).strip() for c in range(8)]
        if header != expected:
            raise DataError(f"屈臣氏欄位和預期不同：{', '.join(header)}")

        rows, name_to_barcode = [], {}
        for r in range(3, ws.nrows):
            v = ws.row_values(r, 0, 8)
            item = to_code(v[0])
            if not item.isdigit():   # 跳過空白列與合計列
                continue
            store = to_code(v[1])
            if store.isdigit():
                store = store.zfill(4)
            barcode = to_code(v[2])
            name = str(v[3]).strip()
            day = datetime.strptime(to_code(v[6]), '%Y%m%d').strftime('%Y-%m-%d')
            rows.append(make_row('屈臣氏', store, day, barcode, item, name, to_code(v[7]),
                                 to_number(v[4]), to_number(v[5])))
            name_to_barcode[normalize_name(name)] = barcode
        return rows, name_to_barcode
    finally:
        wb.release_resources()


def parse_cosmed_date(x):
    if isinstance(x, (datetime, date)):
        return x.strftime('%Y-%m-%d')
    m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', str(x or ''))
    return date(int(m[1]), int(m[2]), int(m[3])).strftime('%Y-%m-%d') if m else None


def cosmed_code(cell):
    """品號保留開頭的 0：通常是文字；若存成數字，依儲存格格式（如 000000）補 0。"""
    code = to_code(cell.value)
    fmt = cell.number_format or ''
    if isinstance(cell.value, (int, float)) and re.fullmatch(r'0+', fmt):
        code = code.zfill(len(fmt))
    return code


def read_cosmed(raw_dir):
    """康是美：商品 × 日期 寬表格，上半部金額、下半部數量。回傳 品號 → {name, category, days}。"""
    path = find_raw_file(raw_dir, '康是美*.xls*')
    print(f'讀取康是美：{path.name}')
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb.worksheets[0]
        # 第 3 列 H 欄起是日期，例如「2026年7月1日」
        dates = {}
        for c in range(8, ws.max_column + 1):
            d = parse_cosmed_date(ws.cell(3, c).value)
            if d:
                dates[c] = d
        if not dates:
            raise DataError('康是美第 3 列找不到日期欄')

        metric, category = None, ''
        products = {}   # 品號 → {name, category, days: {日期: {amount, qty}}}
        for row in ws.iter_rows(min_row=4):
            label = str(row[0].value or '').strip()
            if label:
                if label.startswith('銷售淨額'):
                    metric = 'amount'
                elif label.startswith('銷售數量'):
                    metric = 'qty'
                else:
                    raise DataError(f'康是美 A 欄出現未知的指標：{label}')
            if not metric:
                continue
            if row[3].value not in (None, ''):   # 小分類只寫在每組第一列
                category = str(row[3].value).strip()

            code = cosmed_code(row[5])
            if not code:
                continue
            if code not in products:
                name = str(row[4].value or '').strip() or '（原始檔無品名）'
                products[code] = {'name': name, 'category': category,
                                  'days': defaultdict(lambda: {'amount': 0.0, 'qty': 0.0})}
            p = products[code]
            for c, d in dates.items():
                value = row[c - 1].value
                if value is None or value == '':
                    continue
                p['days'][d][metric] = float(value)
        return products
    finally:
        wb.close()


def update_cosmed_map(map_path, products, name_to_barcode):
    """保留對照表裡已有的品號（含手動填寫），新品號才自動配對。"""
    mapping = {}
    if map_path.exists():
        with open(map_path, encoding='utf-8-sig', newline='') as f:
            for m in csv.DictReader(f):
                mapping[m['品號']] = m
    for code, p in products.items():
        if code in mapping:
            continue
        barcode = name_to_barcode.get(normalize_name(p['name']), '')
        mapping[code] = {'品號': code, '品名': p['name'], '條碼': barcode,
                         '配對方式': '自動（品名相同）' if barcode else '待補'}
    write_csv(map_path, MAP_COLUMNS, mapping.values())
    return mapping


def write_csv(path, columns, rows):
    # UTF-8 BOM 讓 Excel 正確顯示中文；全部欄位加引號
    def cell(v):
        return fmt_number(v) if isinstance(v, float) else v
    try:
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f, quoting=csv.QUOTE_ALL)
            w.writerow(columns)
            for r in rows:
                w.writerow([cell(r.get(c, '')) for c in columns])
    except PermissionError:
        raise DataError(f'無法寫入 {path}，請先關掉開著這個檔案的程式（例如 Excel）')


def print_summary(rows, out_path, mapping, map_path):
    print()
    print(f'輸出：{out_path}（{len(rows)} 筆）')
    groups = defaultdict(lambda: [0, 0.0, 0.0])
    for r in rows:
        g = groups[f"{r['通路']} {r['日期'][:7]}"]
        g[0] += 1
        g[1] += r['銷售數量']
        g[2] += r['銷售金額']
    for key in sorted(groups):
        n, qty, amt = groups[key]
        print(f'  {key}  筆數 {n:6}  數量 {qty:8,.0f}  金額 {amt:14,.2f}')
    unmatched = sum(1 for m in mapping.values() if not m['條碼'])
    print(f'康是美品號對照：共 {len(mapping)} 個品號，{len(mapping) - unmatched} 個有條碼，'
          f'{unmatched} 個待補 → {map_path}')


def main():
    parser = argparse.ArgumentParser(description='整理各通路 sellout Excel')
    parser.add_argument('--raw-dir', type=Path, default=ROOT / 'data' / 'raw')
    parser.add_argument('--out-dir', type=Path, default=ROOT / 'data' / 'processed')
    args = parser.parse_args()

    rows, name_to_barcode = read_watsons(args.raw_dir)
    products = read_cosmed(args.raw_dir)

    args.out_dir.mkdir(parents=True, exist_ok=True)
    map_path = args.out_dir / '康是美品號對照.csv'
    mapping = update_cosmed_map(map_path, products, name_to_barcode)

    for code, p in products.items():
        for d in sorted(p['days']):
            x = p['days'][d]
            if x['amount'] == 0 and x['qty'] == 0:
                continue
            # 品名、條碼以對照表為準，方便手動補上原始檔缺的資料
            m = mapping[code]
            rows.append(make_row('康是美', '', d, m['條碼'], code, m['品名'], p['category'],
                                 x['qty'], x['amount']))

    out_path = args.out_dir / 'sellout_daily.csv'
    write_csv(out_path, COLUMNS, rows)
    print_summary(rows, out_path, mapping, map_path)

    # 順便更新 Dashboard
    import build_dashboard
    build_dashboard.build(out_path, ROOT / 'dashboard' / 'template.html', ROOT / 'dashboard' / 'sellout_dashboard.html')


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    try:
        main()
    except DataError as e:
        print(f'錯誤：{e}', file=sys.stderr)
        sys.exit(1)
