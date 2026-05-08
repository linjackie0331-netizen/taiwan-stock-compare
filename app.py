#!/usr/bin/env python3
"""
Flask 後端：作為瀏覽器與 Goodinfo.tw 之間的代理
瀏覽器無法直接爬 Goodinfo（CORS 限制），所以需要這個後端
"""
import os
from flask import Flask, request, jsonify, send_from_directory
from pathlib import Path
import stock_analyze as sa

app = Flask(__name__)
TMPL = Path(__file__).parent / 'templates'


@app.after_request
def add_cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response


@app.route('/')
def index():
    return send_from_directory(TMPL, 'index.html')


@app.route('/api/ping')
def ping():
    token = os.environ.get('FINMIND_TOKEN', '')
    return jsonify({'ok': True, 'token_set': bool(token), 'token_prefix': token[:10] if token else ''})


@app.route('/api/data', methods=['POST', 'OPTIONS'])
def data():
    """輕量端點：只回傳 metrics（前端自行計算存活評估 & 財務偵測）"""
    if request.method == 'OPTIONS':
        return '', 200
    body      = request.get_json(force=True) or {}
    raw       = body.get('stocks', [])
    stock_ids = [str(s).strip() for s in raw if str(s).strip()][:3]
    years     = max(3, min(10, int(body.get('years', 5))))
    no_cache  = bool(body.get('no_cache', False))

    if not stock_ids:
        return jsonify({'error': '請至少輸入一個股票代碼'}), 400

    result, errors = {}, {}
    for sid in stock_ids:
        try:
            d = sa.fetch_stock(sid, max_years=years, use_cache=not no_cache)
            result[sid] = {'years': d['years'], 'metrics': d['metrics']}
        except Exception as e:
            errors[sid] = str(e)

    if not result:
        return jsonify({'error': '無法取得任何資料', 'details': errors}), 502

    return jsonify({'ok': True, 'stocks': result, 'errors': errors})


@app.route('/api/analyze', methods=['POST', 'OPTIONS'])
def analyze():
    """完整端點：包含存活評估 & 財務偵測（供本地 templates/index.html 使用）"""
    if request.method == 'OPTIONS':
        return '', 200
    body      = request.get_json(force=True) or {}
    raw       = body.get('stocks', [])
    stock_ids = [str(s).strip() for s in raw if str(s).strip()][:5]
    years     = max(3, min(10, int(body.get('years', 5))))
    no_cache  = bool(body.get('no_cache', False))

    if not stock_ids:
        return jsonify({'error': '請至少輸入一個股票代碼'}), 400

    all_stocks, errors = {}, {}
    for sid in stock_ids:
        try:
            all_stocks[sid] = sa.fetch_stock(sid, max_years=years, use_cache=not no_cache)
        except Exception as e:
            errors[sid] = str(e)

    if not all_stocks:
        return jsonify({'error': '無法取得任何資料', 'details': errors}), 502

    result = {}
    for sid, data in all_stocks.items():
        result[sid] = {
            'years':    data['years'],
            'metrics':  data['metrics'],
            'survival': sa.calc_survival_score(data),
            'forensic': sa.calc_forensic_score(data),
        }

    if len(all_stocks) >= 2:
        result['_cross'] = sa.build_cross_forensic(all_stocks)

    return jsonify({'ok': True, 'stocks': result, 'errors': errors})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print('\n🚀 台灣股票財務健診工具')
    print(f'   瀏覽器開啟: http://localhost:{port}\n')
    app.run(host='0.0.0.0', port=port, debug=False)
