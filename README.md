# Sellout Dashboard

整理各通路的 sellout（銷售出貨）Excel 檔案，彙整成統一格式，並透過 Dashboard 視覺化呈現銷售狀況。

## 專案目的

各通路提供的 sellout 報表格式不一（欄位名稱、日期格式、產品編號、檔案結構都可能不同），人工整理既耗時又容易出錯。本專案希望：

1. **自動整理資料**：讀取各通路的 sellout Excel，清理並轉換成一致的欄位與格式。
2. **彙整成單一資料集**：將所有通路的資料合併，方便跨通路比較與分析。
3. **Dashboard 呈現**：以圖表呈現銷售趨勢、通路表現、產品排行等關鍵指標，讓團隊快速掌握狀況。

## 處理流程

```
各通路 Excel  →  資料清理與欄位對應  →  合併成統一資料集  →  Dashboard
```

1. **收集**：將各通路的原始 Excel 放入指定資料夾。
2. **清理**：依各通路的格式做欄位對應、日期與數值格式統一、去除空白或重複資料。
3. **合併**：輸出整合後的資料檔，作為 Dashboard 的資料來源。
4. **呈現**：Dashboard 讀取整合資料並產生圖表。

## Dashboard 預計呈現內容

- 整體銷售金額與數量（依日、週、月）
- 各通路銷售比較
- 產品／品項銷售排行
- 期間比較（例如月增率、年增率）

## 使用方式

1. 把各通路的 Excel 放進 `data/raw`。
2. 執行 `python scripts\clean_sellout.py`（第一次使用先執行 `python -m pip install -r requirements.txt`）。
3. 雙擊 `dashboard/sellout_dashboard.html`，用瀏覽器開啟 Dashboard。

或使用網頁版（可直接上傳 Excel、下載整理後的 Excel）：

```
python -m streamlit run app.py
```

瀏覽器會開啟 http://localhost:8501。

## 資料夾結構

```
sellout/
├── data/
│   ├── raw/          # 各通路原始 Excel
│   └── processed/    # 整理後的統一資料
├── scripts/
│   ├── clean_sellout.py     # 資料清理與合併
│   └── build_dashboard.py   # 產生 Dashboard
├── dashboard/
│   ├── template.html            # Dashboard 樣板
│   └── sellout_dashboard.html   # 產生出來的 Dashboard（含資料，不放進 Git）
├── notes/format.md   # 統一格式與處理規則
├── docs/             # 網站（只放網頁和加密後的資料檔）
├── app.py            # Streamlit 網頁版
└── README.md
```

## 注意事項

- 原始 Excel 可能含有敏感的銷售資料，請勿上傳至公開的版本控制或分享空間。
- 新增通路時，需補上該通路的欄位對應規則。
