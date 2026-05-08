# Taiwan Stock Compare 🇹🇼

台灣股票多股財務健診工具 — 輸入股票代碼，即時從 Goodinfo.tw 抓取財報、計算財務指標、金融檢察官偵查財報異常。

## 功能特色

| 功能 | 說明 |
|------|------|
| **🌐 Web App** | 瀏覽器輸入代碼即時分析，支援上市/上櫃/已下市歷史財報 |
| **多股同業比較** | 1~5 支股票同時分析，最佳/最差自動標色 |
| **📖 白話指標說明** | 每個財務指標都有公式＋白話說明＋求職者視角（點 ❓ 即看） |
| **🛡️ 存活評估** | 這公司 2 年後還在嗎？自動燈號判斷（低/中/高風險）|
| **⚖️ 金融檢察官** | 財報異常偵測：盈餘品質、資產膨脹、掏空警示 |
| **🔍 跨公司對比裁定** | 比較兩間公司的財報可疑程度，給出求職建議 |
| **本地快取** | 同日不重複爬蟲（SQLite），節省時間 |
| **Excel 輸出** | 含金融檢察官、求職評估、趨勢箭頭的完整報表 |

## 快速開始（Web App — 推薦）

```bash
# 1. 安裝依賴
pip install -r requirements.txt

# 2. 啟動 Web 服務
python app.py

# 3. 瀏覽器開啟
#    http://localhost:5000
#    輸入股票代碼（如 2330、2454）→ 點「開始分析」
```

> ⚠️ **架構說明**：瀏覽器無法直接爬 Goodinfo.tw（CORS 限制），所以需要本機執行 `python app.py` 作為代理後端，再用瀏覽器操作。

## CLI 模式（產生 HTML + Excel 檔案）

```bash
python stock_analyze.py 2330 2454          # 比較兩間公司
python stock_analyze.py 2317 2330 2454     # 最多 5 間同時比較
python stock_analyze.py 2330 --years 7    # 拉長歷史到 7 年
python stock_analyze.py 2330 --no-cache   # 強制重新抓取
# 報告輸出在 output/ 目錄
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
