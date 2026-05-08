#!/usr/bin/env python3
"""
taiwan-stock-compare: 台灣股票多股財務比較分析工具
用法：
  python stock_analyze.py 2317              # 單股詳細分析
  python stock_analyze.py 2317 2330 2454    # 多股同業比較
  python stock_analyze.py 2317 --years 5   # 指定年數
  python stock_analyze.py 2317 --no-cache  # 強制重新抓取
"""

import sys
import time
import json
import sqlite3
import argparse
from datetime import date
from pathlib import Path

import requests
from bs4 import BeautifulSoup

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

BASE_DIR = Path(__file__).parent
CACHE_DB  = BASE_DIR / "cache.db"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# ─── 快取層 ────────────────────────────────────────────────

def init_cache():
    conn = sqlite3.connect(CACHE_DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS stock_cache (
            stock_id TEXT, rpt_cat TEXT, fetch_date TEXT, html TEXT,
            PRIMARY KEY (stock_id, rpt_cat, fetch_date)
        )
    """)
    conn.commit()
    return conn

# ─── 抓取層 ────────────────────────────────────────────────

def get_client_key():
    tz_offset = -480
    now_ms = time.time() * 1000
    days_since_epoch = now_ms / 86400000
    days_adjusted = days_since_epoch - tz_offset / 1440
    client_key = f"2.8|38057.1435627105|46946.0324515993|{tz_offset}|{days_adjusted}|{days_adjusted}"
    return client_key, days_adjusted

def _fetch_from_goodinfo(stock_id, rpt_cat, days_adjusted, client_key):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Referer': 'https://goodinfo.tw/'
    }
    cookies = {'CLIENT_KEY': client_key}
    url = (f"https://goodinfo.tw/tw/StockFinDetail.asp"
           f"?RPT_CAT={rpt_cat}&STOCK_ID={stock_id}&REINIT={days_adjusted:.10f}")
    r = requests.get(url, headers=headers, cookies=cookies, timeout=20)
    r.encoding = 'utf-8'
    return r.text

def fetch_report(stock_id, rpt_cat, days_adjusted, client_key, conn, use_cache=True):
    today = date.today().isoformat()
    if use_cache:
        row = conn.execute(
            "SELECT html FROM stock_cache WHERE stock_id=? AND rpt_cat=? AND fetch_date=?",
            (stock_id, rpt_cat, today)
        ).fetchone()
        if row:
            return BeautifulSoup(row[0], 'html.parser')

    html = _fetch_from_goodinfo(stock_id, rpt_cat, days_adjusted, client_key)
    conn.execute(
        "INSERT OR REPLACE INTO stock_cache VALUES (?,?,?,?)",
        (stock_id, rpt_cat, today, html)
    )
    conn.commit()
    return BeautifulSoup(html, 'html.parser')

# ─── 解析層 ────────────────────────────────────────────────

def parse_table(soup, max_years=5):
    tables = soup.find_all('table')
    if len(tables) < 7:
        return {}, []
    t = tables[6]
    rows = t.find_all('tr')
    years, data = [], {}

    for i, row in enumerate(rows):
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        row_data = [c.get_text(strip=True) for c in cells]

        if i == 0 and any(y in row_data for y in [str(y) for y in range(2018, 2026)]):
            for val in row_data[1:]:
                if len(val) == 4 and val.isdigit() and len(years) < max_years:
                    years.append(val)
            continue

        if len(row_data) >= 3 and row_data[0]:
            field_name = row_data[0]
            values = {}
            val_cols = row_data[1:]
            for j, yr in enumerate(years):
                idx = j * 2
                if idx < len(val_cols):
                    raw = val_cols[idx]
                    try:
                        values[yr] = float(raw.replace(',', ''))
                    except Exception:
                        values[yr] = None
            if values:
                data[field_name] = values

    return data, years

def get_stock_name(soup):
    for tag in soup.find_all(['h1', 'h2', 'title']):
        txt = tag.get_text(strip=True)
        if txt:
            return txt[:20]
    return ''

# ─── 指標計算層 ────────────────────────────────────────────

def _g(table, key, yr):
    return table.get(key, {}).get(yr)

def _find(table, *keywords):
    for key in table:
        if all(kw in key for kw in keywords):
            return key
    return None

def safe_div(a, b, pct=True):
    if a is not None and b and b != 0:
        return a / b * (100 if pct else 1)
    return None

def calc_cagr(v_start, v_end, n_years):
    if v_start and v_end and v_start > 0 and n_years > 0:
        return ((v_end / v_start) ** (1 / n_years) - 1) * 100
    return None

def calculate_metrics(is_d, bs_d, cf_d, years):
    metrics = {}
    for yr in years:
        rev_k  = _find(is_d, '營業收入')
        gp_k   = _find(is_d, '毛利') or _find(is_d, '營業毛利')
        op_k   = _find(is_d, '營業利益')
        ni_k   = _find(is_d, '稅後淨利') or _find(is_d, '本期淨利')
        eps_k  = _find(is_d, '每股') or _find(is_d, 'EPS')
        ca_k   = _find(bs_d, '流動資產合計') or _find(bs_d, '流動資產總額')
        cl_k   = _find(bs_d, '流動負債合計') or _find(bs_d, '流動負債總額')
        tl_k   = _find(bs_d, '負債總額') or _find(bs_d, '負債合計')
        ta_k   = _find(bs_d, '資產總額') or _find(bs_d, '資產合計')
        eq_k   = _find(bs_d, '股東權益總額') or _find(bs_d, '權益總額')
        cash_k = _find(bs_d, '現金') or _find(bs_d, '約當現金')
        ocf_k  = _find(cf_d, '營業活動') or _find(cf_d, '來自營運')
        capex_k= _find(cf_d, '資本支出') or _find(cf_d, '取得不動產')
        div_k  = _find(cf_d, '股利') or _find(cf_d, '發放現金股利')

        rev  = _g(is_d, rev_k,  yr) if rev_k  else None
        gp   = _g(is_d, gp_k,   yr) if gp_k   else None
        op   = _g(is_d, op_k,   yr) if op_k   else None
        ni   = _g(is_d, ni_k,   yr) if ni_k   else None
        eps  = _g(is_d, eps_k,  yr) if eps_k  else None
        ca   = _g(bs_d, ca_k,   yr) if ca_k   else None
        cl   = _g(bs_d, cl_k,   yr) if cl_k   else None
        tl   = _g(bs_d, tl_k,   yr) if tl_k   else None
        ta   = _g(bs_d, ta_k,   yr) if ta_k   else None
        eq   = _g(bs_d, eq_k,   yr) if eq_k   else None
        cash = _g(bs_d, cash_k, yr) if cash_k else None
        ocf  = _g(cf_d, ocf_k,  yr) if ocf_k  else None
        capex= _g(cf_d, capex_k,yr) if capex_k else None

        fcf = (ocf + capex) if (ocf is not None and capex is not None) else (
              ocf if ocf is not None else None)

        # DuPont 三因子
        net_margin   = safe_div(ni, rev)
        asset_turnover = safe_div(rev, ta, pct=False)  # not percentage
        equity_mult  = safe_div(ta, eq, pct=False)
        roe_dupont   = None
        if all(x is not None for x in [net_margin, asset_turnover, equity_mult]):
            roe_dupont = net_margin / 100 * asset_turnover * equity_mult * 100

        # ROIC = NOPAT / Invested Capital (simplified)
        invested_cap = (ta - cl) if (ta is not None and cl is not None) else None
        nopat = op * 0.8 if op is not None else None  # assume 20% tax
        roic = safe_div(nopat, invested_cap)

        metrics[yr] = {
            'revenue': rev, 'gross_profit': gp, 'op_income': op,
            'net_income': ni, 'eps': eps,
            'cash': cash, 'current_assets': ca, 'current_liabilities': cl,
            'total_liabilities': tl, 'total_assets': ta, 'equity': eq,
            'op_cf': ocf, 'fcf': fcf,
            'gross_margin':    safe_div(gp, rev),
            'op_margin':       safe_div(op, rev),
            'net_margin':      net_margin,
            'current_ratio':   safe_div(ca, cl),
            'debt_ratio':      safe_div(tl, ta),
            'roe':             safe_div(ni, eq),
            'roa':             safe_div(ni, ta),
            'roic':            roic,
            'fcf_margin':      safe_div(fcf, rev),
            'ocf_margin':      safe_div(ocf, rev),
            'roe_dupont':      roe_dupont,
            'asset_turnover':  asset_turnover,
            'equity_mult':     equity_mult,
        }

    # Revenue CAGR
    rev_vals = [(yr, metrics[yr]['revenue']) for yr in years if metrics[yr]['revenue']]
    if len(rev_vals) >= 2:
        yr_start, v_start = rev_vals[-1]
        yr_end,   v_end   = rev_vals[0]
        n = int(yr_end) - int(yr_start)
        cagr = calc_cagr(v_start, v_end, n)
        for yr in years:
            metrics[yr]['rev_cagr'] = cagr
    return metrics

# ─── 主要抓取流程 ──────────────────────────────────────────

def fetch_stock(stock_id, max_years=5, use_cache=True):
    conn = init_cache()
    client_key, days_adjusted = get_client_key()

    print(f"  [{stock_id}] 損益表...", end=' ', flush=True)
    is_soup = fetch_report(stock_id, 'IS_YEAR', days_adjusted, client_key, conn, use_cache)
    is_d, years = parse_table(is_soup, max_years)
    print("✓")

    time.sleep(0.8)
    print(f"  [{stock_id}] 資產負債表...", end=' ', flush=True)
    bs_soup = fetch_report(stock_id, 'BS_YEAR', days_adjusted, client_key, conn, use_cache)
    bs_d, _ = parse_table(bs_soup, max_years)
    print("✓")

    time.sleep(0.8)
    print(f"  [{stock_id}] 現金流量表...", end=' ', flush=True)
    cf_soup = fetch_report(stock_id, 'CF_YEAR', days_adjusted, client_key, conn, use_cache)
    cf_d, _ = parse_table(cf_soup, max_years)
    print("✓")

    conn.close()
    years = years[:max_years]
    metrics = calculate_metrics(is_d, bs_d, cf_d, years)

    return {
        'stock_id': stock_id,
        'years': years,
        'income_statement': is_d,
        'balance_sheet': bs_d,
        'cash_flow': cf_d,
        'metrics': metrics,
    }

# ─── 輔助格式化 ────────────────────────────────────────────

def fmt(v, decimals=1, suffix=''):
    if v is None:
        return '—'
    return f"{v:,.{decimals}f}{suffix}"

def fmt_b(v):  # billions with unit
    return fmt(v, 1, ' 億')

def pct_color(v, higher_better=True):
    if v is None:
        return '#718096'
    if higher_better:
        return '#38a169' if v > 0 else '#e53e3e'
    return '#e53e3e' if v > 0 else '#38a169'

def rank_values(stocks_metrics, yr, key, higher_better=True):
    vals = [(sid, stocks_metrics[sid].get(yr, {}).get(key)) for sid in stocks_metrics]
    vals = [(sid, v) for sid, v in vals if v is not None]
    if not vals:
        return {}
    vals.sort(key=lambda x: x[1], reverse=higher_better)
    colors = {vals[0][0]: '#276749', vals[-1][0]: '#9b2c2c'}
    return colors

# ─── HTML 生成：比較儀表板 ────────────────────────────────

CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Microsoft JhengHei', 'Noto Sans TC', sans-serif; background: #f0f4f8; color: #2d3748; }
.header { background: linear-gradient(135deg, #1a365d 0%, #2b6cb0 50%, #3182ce 100%);
  color: white; padding: 20px 32px; }
.header h1 { font-size: 1.4rem; margin-bottom: 4px; }
.header .sub { font-size: 0.82rem; opacity: 0.85; }
.tabs { display: flex; background: white; border-bottom: 2px solid #e2e8f0; padding: 0 24px; }
.tab { padding: 12px 20px; cursor: pointer; font-size: 0.9rem; font-weight: 600;
  color: #718096; border-bottom: 3px solid transparent; margin-bottom: -2px; }
.tab.active { color: #2b6cb0; border-bottom-color: #2b6cb0; }
.tab-content { display: none; padding: 24px 32px; }
.tab-content.active { display: block; }
.section-title { font-size: 1rem; font-weight: 700; color: #2d3748; margin: 20px 0 12px; padding-left: 10px; border-left: 4px solid #3182ce; }
.kpi-row { display: flex; gap: 14px; margin-bottom: 20px; flex-wrap: wrap; }
.kpi-card { flex: 1; min-width: 160px; background: white; border-radius: 10px; padding: 16px 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); border-left: 4px solid #3182ce; }
.kpi-card.green { border-left-color: #38a169; }
.kpi-card.orange { border-left-color: #dd6b20; }
.kpi-card.red { border-left-color: #e53e3e; }
.kpi-card.purple { border-left-color: #805ad5; }
.kpi-label { font-size: 0.74rem; color: #718096; margin-bottom: 4px; font-weight: 500; }
.kpi-value { font-size: 1.5rem; font-weight: 700; color: #2d3748; }
.kpi-change { font-size: 0.78rem; margin-top: 3px; }
.up { color: #38a169; } .down { color: #e53e3e; } .neutral { color: #718096; }
.charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-bottom: 20px; }
.chart-card { background: white; border-radius: 10px; padding: 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }
.chart-card.full { grid-column: 1 / -1; }
.chart-title { font-size: 0.88rem; font-weight: 700; color: #4a5568; margin-bottom: 12px; }
.chart-container { position: relative; height: 220px; }
.cmp-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }
.cmp-table th { background: #2b6cb0; color: white; padding: 10px 14px; text-align: center; }
.cmp-table th:first-child { text-align: left; }
.cmp-table td { padding: 9px 14px; text-align: right; border-bottom: 1px solid #e2e8f0; }
.cmp-table td:first-child { text-align: left; font-weight: 600; color: #4a5568; background: #f7fafc; }
.cmp-table tr:last-child td { border-bottom: none; }
.cmp-table .cat-row td { background: #ebf8ff; color: #2b6cb0; font-weight: 700; font-size: 0.82rem; }
.best { background: #f0fff4 !important; color: #276749 !important; font-weight: 700; }
.worst { background: #fff5f5 !important; color: #9b2c2c !important; }
.stock-badge { display: inline-block; padding: 2px 8px; border-radius: 99px; font-size: 0.75rem; font-weight: 700; margin-right: 4px; }
.insight-box { background: linear-gradient(135deg, #ebf8ff, #e6fffa); border: 1px solid #bee3f8;
  border-radius: 10px; padding: 16px 20px; margin-bottom: 18px; }
.insight-box h3 { color: #2b6cb0; font-size: 0.88rem; margin-bottom: 8px; }
.insight-box li { font-size: 0.84rem; color: #4a5568; padding: 2px 0 2px 16px; position: relative; list-style: none; }
.insight-box li::before { content: '▸'; position: absolute; left: 0; color: #3182ce; }
.stock-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 20px; }
.stock-summary-card { background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); border-top: 4px solid #3182ce; }
.stock-summary-card h2 { font-size: 1.1rem; color: #2d3748; margin-bottom: 12px; }
.metric-row { display: flex; justify-content: space-between; padding: 4px 0; border-bottom: 1px solid #f0f4f8; font-size: 0.84rem; }
.metric-row:last-child { border-bottom: none; }
.metric-name { color: #718096; }
.metric-val { font-weight: 600; color: #2d3748; }
@media (max-width: 768px) { .charts-grid { grid-template-columns: 1fr; } .kpi-row { flex-direction: column; } }
"""

STOCK_COLORS = ['#3182ce', '#38a169', '#dd6b20', '#805ad5', '#e53e3e']

def js_array(data_list):
    return '[' + ', '.join(str(v) if v is not None else 'null' for v in data_list) + ']'

def build_comparison_html(all_stocks):
    stock_ids = list(all_stocks.keys())
    n = len(stock_ids)
    # Use first stock's years as reference, intersect
    all_years = all_stocks[stock_ids[0]]['years']

    latest_yr = all_years[0] if all_years else '—'
    color_map = {sid: STOCK_COLORS[i % len(STOCK_COLORS)] for i, sid in enumerate(stock_ids)}

    def m(sid, yr, key):
        return all_stocks[sid]['metrics'].get(yr, {}).get(key)

    def best_worst_tds(metric_key, yr, higher_better=True, fmt_fn=None):
        vals = {sid: m(sid, yr, metric_key) for sid in stock_ids}
        valid = {sid: v for sid, v in vals.items() if v is not None}
        best = max(valid, key=valid.get) if valid else None
        worst = min(valid, key=valid.get) if valid else None
        if not higher_better:
            best, worst = worst, best
        tds = []
        for sid in stock_ids:
            v = vals[sid]
            if fmt_fn:
                display = fmt_fn(v)
            else:
                display = fmt(v) if v is not None else '—'
            cls = ' class="best"' if sid == best else (' class="worst"' if sid == worst else '')
            tds.append(f'<td{cls}>{display}</td>')
        return ''.join(tds)

    # Build comparison table rows
    def cmp_row(label, key, yr, higher_better=True, fmt_fn=None):
        return f'<tr><td>{label}</td>{best_worst_tds(key, yr, higher_better, fmt_fn)}</tr>'

    headers = ''.join(f'<th style="color:{color_map[sid]};background:#1a365d;">{sid}</th>' for sid in stock_ids)

    # Stock summary cards
    stock_cards_html = ''
    for i, sid in enumerate(stock_ids):
        color = color_map[sid]
        yr = latest_yr
        cards_data = [
            ('營業收入', fmt_b(m(sid, yr, 'revenue'))),
            ('毛利率', fmt(m(sid, yr, 'gross_margin'), suffix='%')),
            ('稅後淨利率', fmt(m(sid, yr, 'net_margin'), suffix='%')),
            ('EPS (元)', fmt(m(sid, yr, 'eps'), 2)),
            ('ROE', fmt(m(sid, yr, 'roe'), suffix='%')),
            ('負債比率', fmt(m(sid, yr, 'debt_ratio'), suffix='%')),
        ]
        rows_html = ''.join(
            f'<div class="metric-row"><span class="metric-name">{name}</span><span class="metric-val">{val}</span></div>'
            for name, val in cards_data
        )
        stock_cards_html += f'''
<div class="stock-summary-card" style="border-top-color:{color}">
  <h2 style="color:{color}">🏢 {sid}</h2>
  {rows_html}
</div>'''

    # Chart datasets for revenue
    rev_datasets = ''
    for i, sid in enumerate(stock_ids):
        rev_data = js_array([m(sid, yr, 'revenue') for yr in all_years])
        rev_datasets += f'''{{"label":"{sid} 營收","data":{rev_data},"backgroundColor":"{'rgba('+','.join(str(int(color_map[sid].lstrip('#')[j:j+2],16)) for j in (0,2,4))+',0.2)'}" if False else "rgba(49,130,206,0.15)","borderColor":"{color_map[sid]}","borderWidth":2,"type":"bar","yAxisID":"y"}},'''

    margin_datasets = ''
    for i, sid in enumerate(stock_ids):
        gm_data = js_array([m(sid, yr, 'gross_margin') for yr in all_years])
        margin_datasets += f'{{"label":"{sid} 毛利率","data":{gm_data},"borderColor":"{color_map[sid]}","backgroundColor":"{color_map[sid]}","pointRadius":5,"tension":0.3,"type":"line","yAxisID":"y2"}},'

    roe_datasets = ''
    for i, sid in enumerate(stock_ids):
        roe_data = js_array([m(sid, yr, 'roe') for yr in all_years])
        roe_datasets += f'{{"label":"{sid} ROE","data":{roe_data},"borderColor":"{color_map[sid]}","backgroundColor":"{color_map[sid]}","pointRadius":5,"tension":0.3,"fill":false}},'

    nm_datasets = ''
    for i, sid in enumerate(stock_ids):
        nm_data = js_array([m(sid, yr, 'net_margin') for yr in all_years])
        nm_datasets += f'{{"label":"{sid} 淨利率","data":{nm_data},"borderColor":"{color_map[sid]}","backgroundColor":"{color_map[sid]}","pointRadius":5,"tension":0.3,"fill":false}},'

    dr_datasets = ''
    for i, sid in enumerate(stock_ids):
        dr_data = js_array([m(sid, yr, 'debt_ratio') for yr in all_years])
        dr_datasets += f'{{"label":"{sid} 負債比率","data":{dr_data},"borderColor":"{color_map[sid]}","backgroundColor":"{color_map[sid]}","pointRadius":5,"tension":0.3,"fill":false}},'

    fcf_datasets = ''
    for i, sid in enumerate(stock_ids):
        fcf_data = js_array([m(sid, yr, 'fcf') for yr in all_years])
        fcf_datasets += f'{{"label":"{sid} FCF","data":{fcf_data},"backgroundColor":"{color_map[sid]}","borderColor":"{color_map[sid]}","borderWidth":2}},'

    labels_js = str(all_years).replace("'", '"')

    # Comparison table for all years
    def all_years_rows(label, key, higher_better=True, fmt_fn=None):
        rows = ''
        for yr in all_years:
            yr_label = yr if yr == all_years[0] else f'  {yr}'
            rows += f'<tr><td style="padding-left:24px;color:#718096;">{yr}</td>{best_worst_tds(key, yr, higher_better, fmt_fn)}</tr>'
        return f'<tr class="cat-row"><td colspan="{n+1}">📅 {label}</td></tr>' + rows

    cmp_table_html = f'''
<table class="cmp-table">
<thead><tr><th>指標 / 年度</th>{headers}</tr></thead>
<tbody>
{all_years_rows("營業收入 (億元)", "revenue", higher_better=True, fmt_fn=lambda v: fmt(v, 1))}
{all_years_rows("毛利率 (%)", "gross_margin", higher_better=True, fmt_fn=lambda v: fmt(v, 1, '%'))}
{all_years_rows("營業利益率 (%)", "op_margin", higher_better=True, fmt_fn=lambda v: fmt(v, 1, '%'))}
{all_years_rows("稅後淨利率 (%)", "net_margin", higher_better=True, fmt_fn=lambda v: fmt(v, 1, '%'))}
{all_years_rows("EPS (元)", "eps", higher_better=True, fmt_fn=lambda v: fmt(v, 2))}
{all_years_rows("ROE (%)", "roe", higher_better=True, fmt_fn=lambda v: fmt(v, 1, '%'))}
{all_years_rows("ROIC (%)", "roic", higher_better=True, fmt_fn=lambda v: fmt(v, 1, '%'))}
{all_years_rows("流動比率 (%)", "current_ratio", higher_better=True, fmt_fn=lambda v: fmt(v, 0, '%'))}
{all_years_rows("負債比率 (%)", "debt_ratio", higher_better=False, fmt_fn=lambda v: fmt(v, 1, '%'))}
{all_years_rows("自由現金流 (億元)", "fcf", higher_better=True, fmt_fn=lambda v: fmt(v, 1))}
</tbody>
</table>'''

    # DuPont table (latest year)
    dupont_rows = ''
    for sid in stock_ids:
        yr = latest_yr
        nm   = fmt(m(sid, yr, 'net_margin'), 1, '%')
        at   = fmt(m(sid, yr, 'asset_turnover'), 3, 'x') if m(sid, yr, 'asset_turnover') else '—'
        em   = fmt(m(sid, yr, 'equity_mult'), 2, 'x') if m(sid, yr, 'equity_mult') else '—'
        roe  = fmt(m(sid, yr, 'roe'), 1, '%')
        cagr = fmt(m(sid, yr, 'rev_cagr'), 1, '%')
        dupont_rows += f'<tr><td style="color:{color_map[sid]};font-weight:700;">{sid}</td><td>{nm}</td><td>{at}</td><td>{em}</td><td>{roe}</td><td>{cagr}</td></tr>'

    title = ' vs '.join(stock_ids)
    fetched = date.today().isoformat()

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>財務比較: {title}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.5.0/dist/chart.umd.min.js"></script>
<style>{CSS}</style>
</head>
<body>
<div class="header">
  <div>
    <h1>📊 台灣股票財務比較分析</h1>
    <div class="sub">分析標的：{title}｜資料來源：Goodinfo.tw｜分析日期：{fetched}｜金額單位：億元 (TWD)</div>
  </div>
</div>
<div class="tabs">
  <div class="tab active" onclick="switchTab('overview',this)">🏠 概覽</div>
  <div class="tab" onclick="switchTab('ops',this)">📈 營運</div>
  <div class="tab" onclick="switchTab('profit',this)">💰 獲利</div>
  <div class="tab" onclick="switchTab('finance',this)">🏦 財務健全度</div>
  <div class="tab" onclick="switchTab('dupont',this)">🔬 進階分析</div>
</div>

<div id="overview" class="tab-content active">
  <div class="section-title">各股 {latest_yr} 年度快速摘要</div>
  <div class="stock-cards">{stock_cards_html}</div>
  <div class="section-title">全期間完整比較表</div>
  {cmp_table_html}
  <p style="font-size:0.75rem;color:#718096;margin-top:8px;">🟢 綠底 = 同期最佳　🔴 紅底 = 同期最差</p>
</div>

<div id="ops" class="tab-content">
  <div class="charts-grid">
    <div class="chart-card"><div class="chart-title">營業收入趨勢 (億元)</div><div class="chart-container"><canvas id="revChart"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">毛利率比較 (%)</div><div class="chart-container"><canvas id="gmChart"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">營業利益率 (%)</div><div class="chart-container"><canvas id="opChart"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">稅後淨利率 (%)</div><div class="chart-container"><canvas id="nmChart"></canvas></div></div>
  </div>
</div>

<div id="profit" class="tab-content">
  <div class="charts-grid">
    <div class="chart-card"><div class="chart-title">ROE 趨勢 (%)</div><div class="chart-container"><canvas id="roeChart"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">自由現金流 (億元)</div><div class="chart-container"><canvas id="fcfChart"></canvas></div></div>
  </div>
</div>

<div id="finance" class="tab-content">
  <div class="charts-grid">
    <div class="chart-card"><div class="chart-title">負債比率 (%)</div><div class="chart-container"><canvas id="drChart"></canvas></div></div>
  </div>
</div>

<div id="dupont" class="tab-content">
  <div class="section-title">DuPont 三因子拆解（{latest_yr}）</div>
  <div class="insight-box">
    <h3>📖 DuPont 公式說明</h3>
    <ul>
      <li>ROE = 淨利率 × 資產周轉率 × 財務槓桿倍數</li>
      <li>淨利率高 → 獲利能力強（品牌/技術護城河）</li>
      <li>資產周轉率高 → 資產使用效率高（零售/製造業優勢）</li>
      <li>財務槓桿高 → 使用更多借款放大獲利（金融業常見，需注意風險）</li>
    </ul>
  </div>
  <table class="cmp-table">
    <thead><tr><th>股票</th><th>淨利率</th><th>資產周轉率</th><th>財務槓桿</th><th>ROE（實際）</th><th>營收CAGR</th></tr></thead>
    <tbody>{dupont_rows}</tbody>
  </table>
</div>

<script>
function switchTab(name, el) {{
  document.querySelectorAll('.tab-content').forEach(e => e.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(e => e.classList.remove('active'));
  document.getElementById(name).classList.add('active');
  el.classList.add('active');
}}

const labels = {labels_js};
const chartOpts = (yLabel, y2Label) => ({{
  responsive: true, maintainAspectRatio: false,
  plugins: {{ legend: {{ position: 'bottom', labels: {{ font: {{ size: 11 }} }} }} }},
  scales: {{
    x: {{ grid: {{ display: false }} }},
    y: {{ grid: {{ color: 'rgba(0,0,0,0.05)' }}, title: {{ display: !!yLabel, text: yLabel }} }},
  }}
}});

new Chart(document.getElementById('revChart'), {{
  type: 'bar', data: {{ labels, datasets: [{rev_datasets.rstrip(',')}] }},
  options: chartOpts('億元', '')
}});

const lineOpts = {{ responsive: true, maintainAspectRatio: false,
  plugins: {{ legend: {{ position: 'bottom', labels: {{ font: {{ size: 11 }} }} }} }},
  scales: {{ x: {{ grid: {{ display: false }} }}, y: {{ ticks: {{ callback: v => v + '%' }} }} }} }};

new Chart(document.getElementById('gmChart'), {{
  type: 'line', data: {{ labels, datasets: [{margin_datasets.replace(',type:bar', '').rstrip(',')}] }},
  options: lineOpts
}});

new Chart(document.getElementById('opChart'), {{
  type: 'line', data: {{ labels, datasets: [{
    ''.join(f'{{"label":"{sid} 營業利益率","data":{js_array([m(sid, yr, "op_margin") for yr in all_years])},"borderColor":"{color_map[sid]}","backgroundColor":"{color_map[sid]}","pointRadius":5,"tension":0.3,"fill":false}},' for sid in stock_ids).rstrip(',')
  }] }},
  options: lineOpts
}});

new Chart(document.getElementById('nmChart'), {{
  type: 'line', data: {{ labels, datasets: [{nm_datasets.rstrip(',')}] }},
  options: lineOpts
}});

new Chart(document.getElementById('roeChart'), {{
  type: 'line', data: {{ labels, datasets: [{roe_datasets.rstrip(',')}] }},
  options: lineOpts
}});

new Chart(document.getElementById('fcfChart'), {{
  type: 'bar', data: {{ labels, datasets: [{fcf_datasets.rstrip(',')}] }},
  options: chartOpts('億元', '')
}});

new Chart(document.getElementById('drChart'), {{
  type: 'line', data: {{ labels, datasets: [{dr_datasets.rstrip(',')}] }},
  options: lineOpts
}});
</script>
</body>
</html>"""
    return html

# ─── Excel 生成 ────────────────────────────────────────────

def build_excel(all_stocks):
    if not HAS_EXCEL:
        print("⚠️  openpyxl 未安裝，跳過 Excel 輸出。執行: pip install openpyxl")
        return None

    wb = openpyxl.Workbook()
    stock_ids = list(all_stocks.keys())

    # Styles
    hdr_fill  = PatternFill("solid", fgColor="2B6CB0")
    cat_fill  = PatternFill("solid", fgColor="EBF8FF")
    best_fill = PatternFill("solid", fgColor="F0FFF4")
    bold = Font(bold=True)
    hdr_font = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal='center')
    right  = Alignment(horizontal='right')

    def hdr_cell(ws, row, col, val):
        c = ws.cell(row=row, column=col, value=val)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = center

    def set_border(ws, row, col):
        thin = Side(style='thin', color='E2E8F0')
        ws.cell(row, col).border = Border(bottom=thin)

    # ── Sheet 1: 比較摘要 ──
    ws = wb.active
    ws.title = "比較摘要"

    latest_yr = all_stocks[stock_ids[0]]['years'][0]

    ws.cell(1, 1, f"台灣股票財務比較分析 — {' vs '.join(stock_ids)} — {latest_yr}年度").font = Font(bold=True, size=13)
    ws.merge_cells(f'A1:{get_column_letter(len(stock_ids)+1)}1')

    hdr_cell(ws, 2, 1, "指標")
    for i, sid in enumerate(stock_ids, 2):
        hdr_cell(ws, 2, i, sid)

    metrics_to_show = [
        ("─ 規模 ─", None, None, None),
        ("營業收入 (億元)", "revenue", True, lambda v: fmt(v, 1)),
        ("稅後淨利 (億元)", "net_income", True, lambda v: fmt(v, 1)),
        ("自由現金流 (億元)", "fcf", True, lambda v: fmt(v, 1)),
        ("─ 獲利能力 ─", None, None, None),
        ("毛利率 (%)", "gross_margin", True, lambda v: fmt(v, 1)),
        ("營業利益率 (%)", "op_margin", True, lambda v: fmt(v, 1)),
        ("稅後淨利率 (%)", "net_margin", True, lambda v: fmt(v, 1)),
        ("EPS (元)", "eps", True, lambda v: fmt(v, 2)),
        ("─ 報酬率 ─", None, None, None),
        ("ROE (%)", "roe", True, lambda v: fmt(v, 1)),
        ("ROA (%)", "roa", True, lambda v: fmt(v, 1)),
        ("ROIC (%)", "roic", True, lambda v: fmt(v, 1)),
        ("─ 財務健全 ─", None, None, None),
        ("流動比率 (%)", "current_ratio", True, lambda v: fmt(v, 0)),
        ("負債比率 (%)", "debt_ratio", False, lambda v: fmt(v, 1)),
        ("─ 成長 ─", None, None, None),
        ("營收CAGR (%)", "rev_cagr", True, lambda v: fmt(v, 1)),
    ]

    row = 3
    for label, key, higher_better, fmt_fn in metrics_to_show:
        if key is None:
            c = ws.cell(row, 1, label)
            c.font = Font(bold=True, color="2B6CB0")
            c.fill = cat_fill
            ws.merge_cells(f'A{row}:{get_column_letter(len(stock_ids)+1)}{row}')
            row += 1
            continue

        ws.cell(row, 1, label).font = Font(bold=True)
        vals = {sid: all_stocks[sid]['metrics'].get(latest_yr, {}).get(key) for sid in stock_ids}
        valid_vals = {sid: v for sid, v in vals.items() if v is not None}
        best = max(valid_vals, key=valid_vals.get) if valid_vals else None
        worst = min(valid_vals, key=valid_vals.get) if valid_vals else None
        if not higher_better:
            best, worst = worst, best

        for i, sid in enumerate(stock_ids, 2):
            v = vals[sid]
            c = ws.cell(row, i, fmt_fn(v) if fmt_fn and v is not None else '—')
            c.alignment = right
            if sid == best:
                c.fill = best_fill; c.font = Font(bold=True, color="276749")

        row += 1

    ws.column_dimensions['A'].width = 22
    for i in range(2, len(stock_ids) + 2):
        ws.column_dimensions[get_column_letter(i)].width = 14

    # ── One sheet per stock ──
    for sid in stock_ids:
        data = all_stocks[sid]
        years = data['years']
        ws2 = wb.create_sheet(f"{sid}詳細")

        ws2.cell(1, 1, f"{sid} 詳細財務數據").font = Font(bold=True, size=12)
        ws2.merge_cells(f'A1:{get_column_letter(len(years)+1)}1')

        hdr_cell(ws2, 2, 1, "項目")
        for i, yr in enumerate(years, 2):
            hdr_cell(ws2, 2, i, yr)

        sections = [
            ("損益表", data['income_statement']),
            ("資產負債表", data['balance_sheet']),
            ("現金流量表", data['cash_flow']),
        ]
        row = 3
        for sec_name, table in sections:
            c = ws2.cell(row, 1, f"── {sec_name} ──")
            c.font = Font(bold=True, color="2B6CB0"); c.fill = cat_fill
            ws2.merge_cells(f'A{row}:{get_column_letter(len(years)+1)}{row}')
            row += 1
            for field, yr_vals in list(table.items())[:40]:
                ws2.cell(row, 1, field).font = Font(bold=True)
                for i, yr in enumerate(years, 2):
                    v = yr_vals.get(yr)
                    if v is not None:
                        c = ws2.cell(row, i, v)
                        c.alignment = right
                        c.number_format = '#,##0.00'
                row += 1

        # Computed metrics
        c = ws2.cell(row, 1, "── 計算指標 ──")
        c.font = Font(bold=True, color="2B6CB0"); c.fill = cat_fill
        ws2.merge_cells(f'A{row}:{get_column_letter(len(years)+1)}{row}')
        row += 1
        metric_labels = [
            ("毛利率 (%)", "gross_margin"), ("營業利益率 (%)", "op_margin"),
            ("淨利率 (%)", "net_margin"), ("EPS (元)", "eps"),
            ("ROE (%)", "roe"), ("ROA (%)", "roa"), ("ROIC (%)", "roic"),
            ("流動比率 (%)", "current_ratio"), ("負債比率 (%)", "debt_ratio"),
            ("自由現金流 (億)", "fcf"), ("FCF Margin (%)", "fcf_margin"),
        ]
        for label, key in metric_labels:
            ws2.cell(row, 1, label).font = Font(bold=True)
            for i, yr in enumerate(years, 2):
                v = data['metrics'].get(yr, {}).get(key)
                if v is not None:
                    c = ws2.cell(row, i, round(v, 2))
                    c.alignment = right
                    c.number_format = '#,##0.00'
            row += 1

        ws2.column_dimensions['A'].width = 24
        for i in range(2, len(years) + 2):
            ws2.column_dimensions[get_column_letter(i)].width = 12

    out_path = OUTPUT_DIR / f"{'_vs_'.join(stock_ids)}_analysis.xlsx"
    wb.save(out_path)
    return out_path

# ─── CLI 入口 ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='台灣股票財務分析工具')
    parser.add_argument('stocks', nargs='+', help='股票代碼 (1-5支)')
    parser.add_argument('--years', type=int, default=5, help='分析年數 (預設5年)')
    parser.add_argument('--no-cache', action='store_true', help='強制重新抓取，不使用快取')
    parser.add_argument('--no-excel', action='store_true', help='不產生 Excel 檔案')
    args = parser.parse_args()

    stock_ids = args.stocks[:5]
    use_cache = not args.no_cache

    print(f"\n🚀 台灣股票財務分析 — {' / '.join(stock_ids)}")
    print(f"   分析年數: {args.years} 年  |  快取: {'開啟' if use_cache else '關閉'}\n")

    all_stocks = {}
    for sid in stock_ids:
        print(f"📥 抓取 {sid} 資料...")
        try:
            all_stocks[sid] = fetch_stock(sid, max_years=args.years, use_cache=use_cache)
        except Exception as e:
            print(f"  ❌ 抓取 {sid} 失敗: {e}")

    if not all_stocks:
        print("❌ 無資料可分析"); sys.exit(1)

    tag = '_vs_'.join(all_stocks.keys())
    today = date.today().isoformat()

    # HTML
    print(f"\n📊 生成 HTML 報告...")
    html = build_comparison_html(all_stocks)
    html_path = OUTPUT_DIR / f"{tag}_{today}.html"
    html_path.write_text(html, encoding='utf-8')
    print(f"   ✅ {html_path}")

    # Excel
    if not args.no_excel:
        print(f"📋 生成 Excel 報表...")
        xlsx_path = build_excel(all_stocks)
        if xlsx_path:
            print(f"   ✅ {xlsx_path}")

    print(f"\n✨ 完成！檔案儲存於 output/ 目錄")

if __name__ == '__main__':
    main()
