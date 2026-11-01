"""整理屈臣氏、康是美的 sellout Excel，輸出統一格式的每日銷售資料

兩種用法：
1. 命令列：讀 data/raw，輸出
     data/processed/sellout_daily.csv     每日銷售明細（各通路合併）
     data/processed/康是美品號對照.csv      康是美品號 → 國際條碼對照表
   並更新 Dashboard。
     python scripts\\clean_sellout.py
2. 當成模組：app.py 等程式呼叫 clean_files()，傳入 [(檔名, 檔案內容 bytes)]，
   拿回整理結果和問題列，不會寫任何檔案。

康是美報表沒有條碼，程式會用品名自動配對屈臣氏的條碼。
配對不到或配錯的，直接在對照表的「條碼」欄填入正確條碼後重新執行即可，手動填寫的內容不會被覆蓋。
原始檔缺品名的品號，也可以在對照表的「品名」欄補上。
規則細節見 notes/format.md。
"""

import argparse
import csv
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from io import BytesIO
from pathlib import Path

import openpyxl
import xlrd

ROOT = Path(__file__).resolve().parent.parent
COLUMNS = ['通路', '門市', '日期', '條碼', '通路品號', '品名', '通路分類', '銷售數量', '銷售金額']
MAP_COLUMNS = ['品號', '品名', '條碼', '配對方式']
ISSUE_COLUMNS = ['檔案', '通路', '列號', '類型', '問題', '內容']
EXCLUDED, WARNING = '已排除', '需留意'      # 問題列的兩種類型
NO_NAME = '（原始檔無品名）'

WATSONS_SHEET = 'DailySaleExcel (2)'   # 包含 5～6 月完整資料，其他工作表是重複或樞紐分析
WATSONS_HEADER = ['貨號', '門市', '條碼', '品名', '銷售數量', '銷售金額', '日期', '部門分類']


class DataError(Exception):
    """原始檔內容和預期不同，需要使用者處理。"""


# ---------------------------------------------------------------- 基本轉換

def text(x):
    return '' if x is None else str(x).strip()


def to_code(x):
    """條碼、品號等代碼轉成字串；Excel 存成數字的，轉回完整整數字串。"""
    if x is None:
        return ''
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x).strip()


def to_number(x):
    """空白算 0；不是數字時丟出 ValueError。"""
    if x is None or x == '':
        return 0.0
    return float(x)


def round2(x):
    """四捨五入到小數 2 位。內建 round() 會把 1227.135 捨成 1227.13（浮點誤差），所以改用 Decimal。"""
    return float(Decimal(repr(x)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


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


def parse_watsons_date(x):
    if isinstance(x, (datetime, date)):
        return x.strftime('%Y-%m-%d')
    return datetime.strptime(to_code(x), '%Y%m%d').strftime('%Y-%m-%d')


def parse_cosmed_date(x):
    if isinstance(x, (datetime, date)):
        return x.strftime('%Y-%m-%d')
    m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', text(x))
    return date(int(m[1]), int(m[2]), int(m[3])).strftime('%Y-%m-%d') if m else None


def make_row(channel, store, day, barcode, item, name, category, qty, amount):
    return dict(zip(COLUMNS, [channel, store, day, barcode, item, name, category, qty, round2(amount)]))


def make_issue(filename, channel, excel_row, kind, problem, content=''):
    return dict(zip(ISSUE_COLUMNS, [filename, channel, excel_row, kind, problem, content]))


def get(row, i):
    return row[i] if i < len(row) else None


# ---------------------------------------------------------------- 讀 Excel

class Workbook:
    """用檔案內容開啟 Excel（.xls 用 xlrd，.xlsx 用 openpyxl），統一回傳每列的值。"""

    def __init__(self, data, filename):
        self.filename = filename
        self.is_xls = Path(filename).suffix.lower() == '.xls'
        try:
            if self.is_xls:
                self._wb = xlrd.open_workbook(file_contents=data, on_demand=True)
                self.sheet_names = self._wb.sheet_names()
            else:
                self._wb = openpyxl.load_workbook(BytesIO(data), read_only=True, data_only=True)
                self.sheet_names = self._wb.sheetnames
        except Exception as e:
            raise DataError(f'無法用 Excel 格式開啟（{e}）')

    def rows(self, sheet=0):
        """回傳工作表每一列的值（list of list）。空白儲存格為 None，Excel 日期轉成 datetime。

        .xlsx 中存成數字、但格式設為補 0（如 000000）的儲存格，會轉成補 0 後的文字，保留開頭的 0。
        """
        if isinstance(sheet, int):
            sheet = self.sheet_names[sheet]
        if sheet not in self.sheet_names:
            raise DataError(f'找不到工作表「{sheet}」（目前有：{"、".join(self.sheet_names)}）')
        if self.is_xls:
            ws = self._wb.sheet_by_name(sheet)
            out = []
            for r in range(ws.nrows):
                row = []
                for c in ws.row(r):
                    if c.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        row.append(None)
                    elif c.ctype == xlrd.XL_CELL_DATE:
                        row.append(xlrd.xldate_as_datetime(c.value, self._wb.datemode))
                    else:
                        row.append(c.value)
                out.append(row)
            return out
        out = []
        for cells in self._wb[sheet].iter_rows():
            row = []
            for c in cells:
                v = getattr(c, 'value', None)
                fmt = getattr(c, 'number_format', None) or ''
                if isinstance(v, (int, float)) and not isinstance(v, bool) and re.fullmatch(r'0+', fmt):
                    v = to_code(float(v)).zfill(len(fmt))
                row.append(None if v == '' else v)
            out.append(row)
        return out

    def close(self):
        if self.is_xls:
            self._wb.release_resources()
        else:
            self._wb.close()


def detect_channel(wb):
    """依檔名、再依內容判斷通路。認不出來回傳 None。"""
    name = wb.filename
    if '明細' in name or WATSONS_SHEET in wb.sheet_names:
        return '屈臣氏'
    if name.startswith('康是美'):
        return '康是美'
    rows = wb.rows(0)[:4]
    if len(rows) >= 4 and text(get(rows[2], 5)) == '品號' and text(get(rows[3], 0)).startswith('銷售'):
        return '康是美'
    return None


# ---------------------------------------------------------------- 屈臣氏

def parse_watsons(wb):
    """屈臣氏：每列一筆 門市 × 商品 × 日期。回傳 (rows, 正規化品名 → 條碼, issues)。"""
    src = wb.rows(WATSONS_SHEET)
    header = [text(get(src[2] if len(src) > 2 else [], c)) for c in range(8)]
    if header != WATSONS_HEADER:
        raise DataError(f"屈臣氏第 3 列欄位和預期不同：{', '.join(header)}"
                        f"（應為：{', '.join(WATSONS_HEADER)}）")

    rows, name_to_barcode, issues = [], {}, []
    issue = lambda r, kind, problem, content: issues.append(
        make_issue(wb.filename, '屈臣氏', r, kind, problem, content))
    for i in range(3, len(src)):
        v = [get(src[i], c) for c in range(8)]
        r = i + 1   # Excel 列號
        if all(text(x) == '' for x in v):
            continue   # 空白列
        item = to_code(v[0])
        if not item.isdigit():
            issue(r, EXCLUDED, '不是商品資料（例如合計列）', '｜'.join(text(x) for x in v if text(x)))
            continue
        store = to_code(v[1])
        if store.isdigit():
            store = store.zfill(4)
        barcode = to_code(v[2])
        name = text(v[3])
        try:
            day = parse_watsons_date(v[6])
        except (ValueError, TypeError):
            issue(r, EXCLUDED, '日期格式錯誤', f'貨號 {item}｜日期 {text(v[6])}')
            continue
        try:
            qty, amount = to_number(v[4]), to_number(v[5])
        except (ValueError, TypeError):
            issue(r, EXCLUDED, '數量或金額不是數字', f'貨號 {item}｜數量 {text(v[4])}｜金額 {text(v[5])}')
            continue
        row = make_row('屈臣氏', store, day, barcode, item, name, to_code(v[7]), qty, amount)
        row['_檔案'], row['_列號'] = wb.filename, r
        rows.append(row)
        if not barcode:
            issue(r, WARNING, '缺條碼', f'貨號 {item}｜{name}')
        if qty < 0 or amount < 0:
            issue(r, WARNING, '數量或金額為負數（可能是退貨）', f'貨號 {item}｜數量 {fmt_number(qty)}｜金額 {fmt_number(amount)}')
        name_to_barcode[normalize_name(name)] = barcode
    return rows, name_to_barcode, issues


# ---------------------------------------------------------------- 康是美

def parse_cosmed(wb):
    """康是美：商品 × 日期 寬表格，上半部金額、下半部數量。

    回傳 (products, issues, 報表資料最新日期)。products：品號 → {name, category, row, days: {日期: {amount, qty}}}
    """
    src = wb.rows(0)
    header = src[2] if len(src) > 2 else []
    # 第 3 列 H 欄起是日期，例如「2026年7月1日」
    dates = {c: d for c in range(7, len(header)) if (d := parse_cosmed_date(header[c]))}
    if not dates:
        raise DataError('康是美第 3 列找不到日期欄（應從 H 欄起，例如「2026年7月1日」）')
    latest = parse_cosmed_date(get(src[3], 6)) if len(src) > 3 else None   # G 欄「資料最新日期」

    metric, category = None, ''
    products, issues = {}, []
    seen = {'amount': set(), 'qty': set()}
    issue = lambda r, kind, problem, content: issues.append(
        make_issue(wb.filename, '康是美', r, kind, problem, content))
    for i in range(3, len(src)):
        row, r = src[i], i + 1
        label = text(get(row, 0))
        if label:
            if label.startswith('銷售淨額'):
                metric = 'amount'
            elif label.startswith('銷售數量'):
                metric = 'qty'
            else:
                raise DataError(f'康是美第 {r} 列 A 欄出現未知的指標：{label}')
        if not metric:
            continue
        if text(get(row, 3)):   # 小分類只寫在每組第一列
            category = text(get(row, 3))

        code = to_code(get(row, 5))
        if not code:
            continue
        if code in seen[metric]:
            issue(r, WARNING, '品號在同一段重複出現，後面的值覆蓋前面', f'品號 {code}')
        seen[metric].add(code)
        if code not in products:
            raw_name = text(get(row, 4))
            products[code] = {'name': raw_name or NO_NAME, 'category': category, 'row': r,
                              'days': defaultdict(lambda: {'amount': 0.0, 'qty': 0.0})}
            if not raw_name:
                issue(r, WARNING, '原始檔無品名', f'品號 {code}')
        p, negative = products[code], False
        for c, d in dates.items():
            value = get(row, c)
            if value is None or value == '':
                continue
            try:
                num = float(value)
            except (ValueError, TypeError):
                issue(r, EXCLUDED, '數值不是數字', f'品號 {code}｜{d}｜{text(value)}')
                continue
            p['days'][d][metric] = num
            negative = negative or num < 0
        if negative:
            issue(r, WARNING, '數量或金額為負數（可能是退貨）', f'品號 {code}')
    return products, issues, latest


# ---------------------------------------------------------------- 康是美品號對照表

def load_mapping(path):
    mapping = {}
    if Path(path).exists():
        with open(path, encoding='utf-8-sig', newline='') as f:
            for m in csv.DictReader(f):
                mapping[m['品號']] = m
    return mapping


def extend_mapping(mapping, products, name_to_barcode):
    """保留對照表裡已有的品號（含手動填寫），新品號才用品名自動配對。回傳 (新對照表, 新增的品號)。"""
    mapping, added = dict(mapping), []
    for code, p in products.items():
        if code in mapping:
            continue
        barcode = name_to_barcode.get(normalize_name(p['name']), '')
        mapping[code] = {'品號': code, '品名': p['name'], '條碼': barcode,
                         '配對方式': '自動（品名相同）' if barcode else '待補'}
        added.append(code)
    return mapping, added


def cosmed_rows(products, mapping, filename):
    """康是美寬表格轉成每列一筆；品名、條碼以對照表為準。回傳 (rows, issues)。"""
    rows, issues = [], []
    for code, p in products.items():
        m = mapping[code]
        if not m['條碼']:
            issues.append(make_issue(filename, '康是美', p['row'], WARNING,
                                     '沒有條碼（對照表待補），無法和屈臣氏合併', f"品號 {code}｜{m['品名']}"))
        for d in sorted(p['days']):
            x = p['days'][d]
            if x['amount'] == 0 and x['qty'] == 0:
                continue
            row = make_row('康是美', '', d, m['條碼'], code, m['品名'], p['category'], x['qty'], x['amount'])
            row['_檔案'], row['_列號'] = filename, p['row']
            rows.append(row)
    return rows, issues


# ---------------------------------------------------------------- 整合

@dataclass
class CleanResult:
    rows: list = field(default_factory=list)          # 統一格式（另含 _檔案、_列號）
    issues: list = field(default_factory=list)        # 問題列，欄位見 ISSUE_COLUMNS
    files: list = field(default_factory=list)         # 每個檔案的辨識結果
    errors: list = field(default_factory=list)        # 整個檔案無法處理的錯誤訊息
    mapping: dict = field(default_factory=dict)       # 康是美品號對照（含本次新增）
    new_codes: list = field(default_factory=list)     # 本次新增到對照表的品號
    partial_days: list = field(default_factory=list)  # [(通路, 日期)] 報表資料最新日期，當天資料不完整

    def problem_rows(self, kind):
        """有某類問題的原始列數（同一列有多個問題只算一次）。"""
        return len({(i['檔案'], i['列號']) for i in self.issues if i['類型'] == kind})


def clean_files(files, mapping=None):
    """整理多個 Excel。files：[(檔名, 檔案內容 bytes)]；mapping：康是美品號對照（load_mapping 的結果）。

    不寫任何檔案。單一檔案出錯時記在 errors，其他檔案照常整理。
    同一筆（通路＋門市＋日期＋通路品號）出現在多個檔案時，只保留第一筆。
    """
    res = CleanResult(mapping=dict(mapping or {}))
    parsed = {'屈臣氏': [], '康是美': []}
    for filename, data in files:
        channel, status = None, ''
        try:
            wb = Workbook(data, filename)
            try:
                channel = detect_channel(wb)
                if channel == '屈臣氏':
                    parsed[channel].append((filename, parse_watsons(wb)))
                elif channel == '康是美':
                    parsed[channel].append((filename, parse_cosmed(wb)))
                else:
                    status = '無法辨識通路（目前支援屈臣氏、康是美），已略過'
            finally:
                wb.close()
        except DataError as e:
            status = f'錯誤：{e}'
            res.errors.append(f'{filename}：{e}')
        res.files.append({'檔案': filename, '通路': channel or '—', '狀態': status})

    # 先屈臣氏（康是美自動配對條碼要用屈臣氏的品名）
    rows, name_to_barcode = [], {}
    for filename, (w_rows, n2b, issues) in parsed['屈臣氏']:
        rows += w_rows
        name_to_barcode.update(n2b)
        res.issues += issues
    for filename, (products, issues, latest) in parsed['康是美']:
        res.mapping, added = extend_mapping(res.mapping, products, name_to_barcode)
        res.new_codes += added
        c_rows, c_issues = cosmed_rows(products, res.mapping, filename)
        rows += c_rows
        res.issues += issues + c_issues
        if latest and ('康是美', latest) not in res.partial_days:
            res.partial_days.append(('康是美', latest))

    seen = set()
    for row in rows:
        key = (row['通路'], row['門市'], row['日期'], row['通路品號'])
        if key in seen:
            res.issues.append(make_issue(row['_檔案'], row['通路'], row['_列號'], EXCLUDED,
                                         '和其他檔案重複', f"{row['日期']}｜{row['通路品號']}｜{row['品名']}"))
            continue
        seen.add(key)
        res.rows.append(row)

    counts = defaultdict(int)
    for row in res.rows:
        counts[row['_檔案']] += 1
    for f in res.files:
        f['筆數'] = counts.get(f['檔案'], 0)
        if not f['狀態']:
            f['狀態'] = '已整理'
    return res


# ---------------------------------------------------------------- 輸出

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


def print_summary(res, out_path, map_path):
    print()
    print(f'輸出：{out_path}（{len(res.rows)} 筆）')
    groups = defaultdict(lambda: [0, 0.0, 0.0])
    for r in res.rows:
        g = groups[f"{r['通路']} {r['日期'][:7]}"]
        g[0] += 1
        g[1] += r['銷售數量']
        g[2] += r['銷售金額']
    for key in sorted(groups):
        n, qty, amt = groups[key]
        print(f'  {key}  筆數 {n:6}  數量 {qty:8,.0f}  金額 {amt:14,.2f}')
    unmatched = sum(1 for m in res.mapping.values() if not m['條碼'])
    print(f'康是美品號對照：共 {len(res.mapping)} 個品號，{len(res.mapping) - unmatched} 個有條碼，'
          f'{unmatched} 個待補 → {map_path}')
    by_problem = defaultdict(int)
    for i in res.issues:
        by_problem[(i['類型'], i['問題'])] += 1
    print(f'問題列：{EXCLUDED} {res.problem_rows(EXCLUDED)} 列、{WARNING} {res.problem_rows(WARNING)} 列')
    for (kind, problem), n in sorted(by_problem.items()):
        print(f'  [{kind}] {problem}：{n}')


def is_temp_file(name):
    """Excel 開檔時產生的暫存檔（~$ 開頭），不是資料。"""
    return name.startswith('~$')


def find_raw_file(raw_dir, pattern):
    files = [f for f in raw_dir.iterdir() if f.is_file() and f.match(pattern) and not is_temp_file(f.name)]
    if len(files) != 1:
        names = ', '.join(f.name for f in files)
        raise DataError(f"data/raw 裡應該剛好有一個符合 '{pattern}' 的檔案，目前找到 {len(files)} 個：{names}")
    return files[0]


def main():
    parser = argparse.ArgumentParser(description='整理各通路 sellout Excel')
    parser.add_argument('--raw-dir', type=Path, default=ROOT / 'data' / 'raw')
    parser.add_argument('--out-dir', type=Path, default=ROOT / 'data' / 'processed')
    args = parser.parse_args()

    paths = [find_raw_file(args.raw_dir, '*明細*.xls*'), find_raw_file(args.raw_dir, '康是美*.xls*')]
    for p in paths:
        print(f'讀取：{p.name}')
    map_path = args.out_dir / '康是美品號對照.csv'
    res = clean_files([(p.name, p.read_bytes()) for p in paths], load_mapping(map_path))
    if res.errors:
        raise DataError('\n'.join(res.errors))

    args.out_dir.mkdir(parents=True, exist_ok=True)
    write_csv(map_path, MAP_COLUMNS, res.mapping.values())
    out_path = args.out_dir / 'sellout_daily.csv'
    write_csv(out_path, COLUMNS, res.rows)
    print_summary(res, out_path, map_path)

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
