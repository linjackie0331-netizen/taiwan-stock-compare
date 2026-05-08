# Taiwan Stock Compare 🇹🇼

台灣股票多股財務比較分析工具 — 自動從 Goodinfo.tw 抓取財報、計算進階指標、產生互動 HTML 儀表板 + Excel 報表。

## 功能特色

| 功能 | 說明 |
|------|------|
| **多股同業比較** | 1~5 支股票同時分析，綠底＝最佳、紅底＝最差（數值相同時不標色） |
| **5 年歷史資料** | 看完整景氣循環（可調整） |
| **進階財務指標** | ROIC、DuPont 三因子、FCF Margin、營收 CAGR |
| **🛡️ 存活評估** | 求職專用：自動判斷財務危機風險（低/中/高），含燈號＋中文結論 |
| **趨勢箭頭** | Excel 每項指標自動標示 ↑↓→ 方向 |
| **本地快取** | 同日不重複爬蟲（SQLite），節省時間 |
| **互動 HTML** | Chart.js 圖表，6 個分頁，直接開啟不需安裝 |
| **Excel 輸出** | 求職評估（第一頁）+ 比較摘要 + 各股詳細數據 |

## 快速開始

```bash
# 安裝依賴
pip install -r requirements.txt

# 單股分析
python stock_analyze.py 2330

# 多股同業比較（推薦）
python stock_analyze.py 2317 2330 2454

# 更長歷史
python stock_analyze.py 2330 --years 7

# 強制更新快取
python stock_analyze.py 2330 --no-cache
```

報告輸出在 `output/` 目錄。

## 分析指標說明

### 基本指標
- 毛利率、營業利益率、稅後淨利率、EPS

### 報酬率指標
- **ROE**：股東權益報酬率
- **ROA**：資產報酬率
- **ROIC**：投入資本報酬率（衡量核心業務效率）

### DuPont 三因子拆解
```
ROE = 淨利率 × 資產周轉率 × 財務槓桿
```
- 淨利率高 → 獲利能力強（品牌/技術護城河）
- 資產周轉率高 → 資產使用效率高
- 財務槓桿高 → 借款放大報酬（注意風險）

### 現金流指標
- **FCF（自由現金流）**：營業現金流 − 資本支出
- **FCF Margin**：FCF / 營收（越高代表賺現金能力越強）
- **營收 CAGR**：複合年均成長率

## 比較表說明

| 顏色 | 意義 |
|------|------|
| 🟢 綠底粗體 | 同期所有比較標的中最佳 |
| 🔴 紅底 | 同期所有比較標的中最差 |
| — | 資料不可用 |

## 常見使用場景

**求職（投資研究/財務分析）**
```bash
# 半導體三雄比較
python stock_analyze.py 2330 2454 3711

# 電子五哥比較
python stock_analyze.py 2317 2354 2382 2395 2412
```

**個人投資**
```bash
# 存股清單評估
python stock_analyze.py 0050 2330 2412 2882

# 金融股比較
python stock_analyze.py 2881 2882 2884 2886
```

## 資料來源

- **財報數據**：[Goodinfo.tw](https://goodinfo.tw/)（台灣上市/上櫃公司）
- **官方申報驗證**：[MOPS 公開資訊觀測站](https://mops.twse.com.tw/)

> ⚠️ 本工具僅供學習與研究，不構成投資建議。資料準確性取決於 Goodinfo.tw。

## 技術架構

```
stock_analyze.py
├── 快取層（SQLite）     ← 避免重複爬蟲
├── 抓取層（Goodinfo）   ← requests + BeautifulSoup
├── 指標計算層           ← 基本 + DuPont + ROIC + CAGR
├── HTML 生成            ← Chart.js 互動儀表板
└── Excel 生成           ← openpyxl 格式化報表
```

## 授權

MIT License — 自由使用、修改、分享。
