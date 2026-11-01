"""Sellout 銷售分析（Streamlit）

啟動：
  python -m streamlit run app.py

上傳各通路的 sellout Excel，用 scripts/clean_sellout.py 的整理邏輯轉成統一格式，
再用篩選器、指標卡、圖表檢視，並可下載整理後的 Excel。
本程式不會修改 data 資料夾裡的任何檔案。
"""

import calendar
import sys
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent / 'scripts'))
import clean_sellout as cs  # noqa: E402

RAW_DIR = cs.ROOT / 'data' / 'raw'
MAP_PATH = cs.ROOT / 'data' / 'processed' / '康是美品號對照.csv'
CHANNEL_COLORS = {'屈臣氏': '#2a78d6', '康是美': '#eb6834'}   # 已通過色盲檢查；新增通路依序加色

st.set_page_config(page_title='Sellout 銷售分析', page_icon='📊', layout='wide')


# ---------------------------------------------------------------- 整理（快取）

@st.cache_data(show_spinner='整理中…')
def run_clean(files, map_mtime):
    """files：((檔名, bytes), ...)。map_mtime 讓對照表被修改後重新整理。"""
    res = cs.clean_files(list(files), cs.load_mapping(MAP_PATH))
    return res


def to_frame(res):
    df = pd.DataFrame(res.rows, columns=cs.COLUMNS)
    if df.empty:
        return df
    df['日期'] = pd.to_datetime(df['日期'])
    df['月份'] = df['日期'].dt.strftime('%Y-%m')
    # 品項：有條碼的跨通路視為同一商品；沒有條碼的用「通路＋品號」
    key = df['條碼'].where(df['條碼'] != '', df['通路'] + ':' + df['通路品號'])
    first_name = df.groupby(key, sort=False)['品名'].first()
    df['品項'] = [f'{first_name[k]}（{b}）' if b else f'{first_name[k]}（{c} 品號 {i}，無條碼）'
                for k, b, c, i in zip(key, df['條碼'], df['通路'], df['通路品號'])]
    return df


def build_excel(detail, issues, mapping):
    """整理後的 Excel：銷售明細、問題列、康是美品號對照。5 萬筆約需 8 秒，所以只在按下載時才產生。"""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as xw:
        detail.to_excel(xw, sheet_name='銷售明細', index=False)
        issues.to_excel(xw, sheet_name='問題列', index=False)
        mapping.to_excel(xw, sheet_name='康是美品號對照', index=False)
        ws = xw.sheets['銷售明細']
        for cell in ws['C'][1:]:
            cell.number_format = 'yyyy-mm-dd'
        for name, widths in {'銷售明細': [8, 8, 12, 16, 10, 44, 18, 10, 14],
                             '問題列': [40, 8, 8, 8, 40, 50],
                             '康是美品號對照': [10, 44, 16, 18]}.items():
            for col, w in zip('ABCDEFGHI', widths):
                xw.sheets[name].column_dimensions[col].width = w
    return buf.getvalue()


def money(v):
    return f'NT${v:,.0f}'


def pct_change(cur, prev):
    return None if not prev else f'{(cur / prev - 1) * 100:+.1f}%'


# ---------------------------------------------------------------- 1. 上傳

st.title('Sellout 銷售分析')
uploads = st.file_uploader('上傳各通路的 sellout Excel（可一次選多個檔案）', type=['xls', 'xlsx'],
                           accept_multiple_files=True,
                           help='目前支援屈臣氏（明細報表）和康是美（商品銷售表現列表）。同一筆資料出現在多個檔案時只算一次。')

if uploads:
    files = tuple((u.name, u.getvalue()) for u in uploads)
elif st.query_params.get('source') == 'raw':
    files = tuple((p.name, p.read_bytes()) for p in sorted(RAW_DIR.glob('*.xls*')) if not cs.is_temp_file(p.name))
    st.caption(f'目前使用 `data\\raw` 裡的 {len(files)} 個檔案。上傳檔案後會改用上傳的檔案。')
else:
    st.info('請上傳 Excel 檔案開始分析。')
    if st.button('或直接使用 data\\raw 裡的檔案'):
        st.query_params['source'] = 'raw'
        st.rerun()
    st.stop()

map_mtime = MAP_PATH.stat().st_mtime if MAP_PATH.exists() else 0
res = run_clean(files, map_mtime)
df_all = to_frame(res)

# ---------------------------------------------------------------- 2. 整理結果與問題列

for err in res.errors:
    st.error(err)

c1, c2, c3, c4 = st.columns(4)
c1.metric('整理後筆數', f'{len(res.rows):,}')
c2.metric('辨識成功的檔案', f"{sum(f['狀態'] == '已整理' for f in res.files)} / {len(res.files)}")
c3.metric('問題列：已排除', f'{res.problem_rows(cs.EXCLUDED):,} 列',
          help='無法使用而沒有放進整理結果的原始列，例如合計列、日期格式錯誤、和其他檔案重複。')
c4.metric('問題列：需留意', f'{res.problem_rows(cs.WARNING):,} 列',
          help='有放進整理結果，但需要確認的列，例如負數（退貨）、缺條碼、原始檔無品名。')

issues_df = pd.DataFrame(res.issues, columns=cs.ISSUE_COLUMNS)
with st.expander('檔案辨識結果與問題列明細'):
    st.dataframe(pd.DataFrame(res.files, columns=['檔案', '通路', '狀態', '筆數']), hide_index=True,
                 width='stretch')
    if issues_df.empty:
        st.success('沒有問題列。')
    else:
        summary = (issues_df.groupby(['類型', '問題']).size().rename('筆數').reset_index()
                   .sort_values(['類型', '筆數'], ascending=[True, False]))
        st.dataframe(summary, hide_index=True, width='stretch')
        kinds = st.multiselect('顯示類型', [cs.EXCLUDED, cs.WARNING], default=[cs.EXCLUDED, cs.WARNING])
        st.dataframe(issues_df[issues_df['類型'].isin(kinds)], hide_index=True, width='stretch', height=320)
if res.new_codes:
    st.warning(f'有 {len(res.new_codes)} 個康是美品號不在對照表裡，已暫時用品名自動配對條碼（不會寫入對照表）。'
               '下載的 Excel 裡有完整的對照表可以參考。')

if df_all.empty:
    st.warning('沒有可以分析的資料。')
    st.stop()

# ---------------------------------------------------------------- 3. 篩選器

st.divider()
f1, f2, f3 = st.columns([2, 2, 4])
dmin, dmax = df_all['日期'].min().date(), df_all['日期'].max().date()
picked = f1.date_input('日期區間', (dmin, dmax), min_value=dmin, max_value=dmax)
start, end = (picked if isinstance(picked, (tuple, list)) and len(picked) == 2 else (dmin, dmax))
channels_all = [c for c in CHANNEL_COLORS if c in set(df_all['通路'])] + \
               sorted(set(df_all['通路']) - set(CHANNEL_COLORS))
channels = f2.multiselect('通路', channels_all, default=channels_all, placeholder='選擇通路')
products = f3.multiselect('品項', sorted(df_all['品項'].unique()), placeholder='全部品項（可輸入關鍵字搜尋）')

partial = [(c, pd.Timestamp(d)) for c, d in res.partial_days]
drop_partial = False
if partial:
    drop_partial = st.checkbox(
        '排除報表資料最新日期（當天資料不完整）：' + '、'.join(f'{c} {d:%Y-%m-%d}' for c, d in partial), value=True)

base = df_all[df_all['通路'].isin(channels)]
if products:
    base = base[base['品項'].isin(products)]
if drop_partial:
    for c, d in partial:
        base = base[~((base['通路'] == c) & (base['日期'] == d))]
df = base[(base['日期'] >= pd.Timestamp(start)) & (base['日期'] <= pd.Timestamp(end))]

if df.empty:
    st.warning('這個篩選條件沒有資料。')
    st.stop()

# ---------------------------------------------------------------- 4. 指標卡

def month_stats(frame, month):
    m = frame[frame['月份'] == month]
    days = m['日期'].nunique()
    return {'qty': m['銷售數量'].sum(), 'amt': m['銷售金額'].sum(), 'days': days,
            'avg': m['銷售金額'].sum() / days if days else 0, 'channels': set(m['通路'])}

latest = df['月份'].max()
prev = (pd.Period(latest, 'M') - 1).strftime('%Y-%m')
cur_s = month_stats(df, latest)
prev_s = month_stats(base, prev)   # 上月不受日期區間限制，取完整上月
label_cur, label_prev = f'{int(latest[5:])} 月', f'{int(prev[5:])} 月'

k1, k2, k3, k4 = st.columns(4)
k1.metric('總銷量', f"{df['銷售數量'].sum():,.0f} 件", help='篩選期間合計')
k2.metric('總金額', money(df['銷售金額'].sum()), help='篩選期間合計')
k3.metric(f'{label_cur}銷售金額（與上月比較）', money(cur_s['amt']),
          pct_change(cur_s['amt'], prev_s['amt']),
          help=f'{label_cur} {cur_s["days"]} 天 vs {label_prev} {prev_s["days"]} 天的合計金額')
k4.metric(f'{label_cur}日均金額（與上月比較）', money(cur_s['avg']),
          pct_change(cur_s['avg'], prev_s['avg']),
          help='每天平均銷售金額，不受天數多寡影響，比較較公平')

notes = []
full_days = calendar.monthrange(int(latest[:4]), int(latest[5:]))[1]
if cur_s['days'] < full_days:
    notes.append(f'{label_cur}只有 {cur_s["days"]} 天資料，合計金額和上月比會偏低，請參考日均金額。')
if not prev_s['days']:
    notes.append(f'{label_prev}沒有資料，無法比較。')
elif cur_s['channels'] != prev_s['channels']:
    notes.append(f'{label_cur}（{"、".join(sorted(cur_s["channels"]))}）和 {label_prev}'
                 f'（{"、".join(sorted(prev_s["channels"]))}）的資料來自不同通路，比較僅供參考。')
if notes:
    st.caption('　'.join(notes))

# ---------------------------------------------------------------- 5. 圖表

st.divider()
metric = st.radio('圖表指標', ['銷售金額', '銷售數量'], horizontal=True)
fmt = ',.0f'
# 軸標籤用「萬」：1,500,000 → 150萬
WAN = alt.Axis(labelExpr="abs(datum.value) >= 10000 ? format(datum.value / 10000, ',.0f') + '萬' : format(datum.value, ',.0f')")


def padded(max_value):
    """長條圖右側留空間給數值標籤。"""
    return alt.Scale(domain=[0, max_value * 1.2 if max_value > 0 else 1], nice=False)
color = alt.Color('通路:N', title='通路',
                  scale=alt.Scale(domain=channels_all, range=[CHANNEL_COLORS.get(c, '#898781') for c in channels_all]),
                  legend=alt.Legend(orient='top', title=None))

# 月銷售趨勢（折線）
monthly = (df.groupby(['月份', '通路'])
           .agg(銷售金額=('銷售金額', 'sum'), 銷售數量=('銷售數量', 'sum'), 天數=('日期', 'nunique'))
           .reset_index())
st.subheader('月銷售趨勢')
st.caption('每個通路一條線。滑鼠移到點上可看天數；不滿整月的月份合計會偏低。')
trend = alt.Chart(monthly).encode(
    x=alt.X('月份:O', title=None, axis=alt.Axis(labelAngle=0)),
    y=alt.Y(f'{metric}:Q', title=None, axis=WAN),
    color=color,
    tooltip=[alt.Tooltip('月份:O'), alt.Tooltip('通路:N'), alt.Tooltip(f'{metric}:Q', format=fmt),
             alt.Tooltip('天數:Q')])
st.altair_chart((trend.mark_line(strokeWidth=2) + trend.mark_point(filled=True, size=80, opacity=1))
                .properties(height=300), width='stretch')

g1, g2 = st.columns([2, 3])

# 通路排行（長條）
by_channel = df.groupby('通路')[metric].sum().reset_index().sort_values(metric, ascending=False)
with g1:
    st.subheader('通路排行')
    st.caption('篩選期間合計。各通路的資料期間可能不同。')
    bars = alt.Chart(by_channel).encode(
        y=alt.Y('通路:N', sort='-x', title=None),
        x=alt.X(f'{metric}:Q', title=None, axis=WAN, scale=padded(by_channel[metric].max())),
        tooltip=[alt.Tooltip('通路:N'), alt.Tooltip(f'{metric}:Q', format=fmt)])
    st.altair_chart((bars.mark_bar(size=24, cornerRadiusEnd=4).encode(color=color.legend(None)) +
                     bars.mark_text(align='left', dx=6).encode(text=alt.Text(f'{metric}:Q', format=fmt)))
                    .properties(height=max(120, 56 * len(by_channel))), width='stretch')

# 熱銷品項前 10 名
by_item = df.groupby(['品項', '通路'])[metric].sum().reset_index()
top = by_item.groupby('品項')[metric].sum().nlargest(10)
top_rows = by_item[by_item['品項'].isin(top.index)]
totals = top.rename('合計').reset_index()
with g2:
    st.subheader('熱銷品項前 10 名')
    st.caption('有條碼的品項會合併兩個通路，顏色區分各通路的貢獻。')
    order = list(top.index)
    seg = alt.Chart(top_rows).mark_bar(size=18).encode(
        y=alt.Y('品項:N', sort=order, title=None, axis=alt.Axis(labelLimit=280)),
        x=alt.X(f'sum({metric}):Q', title=None, axis=WAN, scale=padded(top.max())),
        color=color,
        tooltip=[alt.Tooltip('品項:N'), alt.Tooltip('通路:N'), alt.Tooltip(f'{metric}:Q', format=fmt)])
    labels = alt.Chart(totals).mark_text(align='left', dx=6).encode(
        y=alt.Y('品項:N', sort=order), x=alt.X('合計:Q', scale=padded(top.max())), text=alt.Text('合計:Q', format=fmt))
    st.altair_chart((seg + labels).properties(height=34 * len(order) + 40), width='stretch')

# ---------------------------------------------------------------- 6. 下載

st.divider()
st.subheader('下載整理後的 Excel')
scope = st.radio('匯出範圍', ['全部整理結果', '目前篩選範圍'], horizontal=True,
                 help='Excel 內含三個工作表：銷售明細、問題列、康是美品號對照')
detail = (df_all if scope == '全部整理結果' else df)[cs.COLUMNS].copy()
mapping_df = pd.DataFrame(list(res.mapping.values()), columns=cs.MAP_COLUMNS)
st.download_button(f'下載 Excel（{len(detail):,} 筆）', lambda: build_excel(detail, issues_df, mapping_df),
                   file_name='sellout_整理結果.xlsx', type='primary', on_click='ignore',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
