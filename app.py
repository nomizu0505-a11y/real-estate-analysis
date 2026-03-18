#!/usr/bin/env python3
"""
不動産投資分析レポート Webアプリ
Flask サーバー（HTML配信 + Excel生成 + AI プロキシ）

起動方法:
  python3 app.py

アクセス:
  http://localhost:5050/  （ローカル）
  または公開URLでどの端末からもアクセス可能
"""
import os
import io
import sys
import json
import re
import tempfile
import traceback
from datetime import datetime
from flask import Flask, request, send_file, jsonify, render_template_string, send_from_directory
from flask_cors import CORS

# maisoku_gen.py を同ディレクトリから読み込む
sys.path.insert(0, os.path.dirname(__file__))
from maisoku_gen import generate_maisoku

app = Flask(__name__, static_folder=os.path.dirname(__file__))
CORS(app)

# ===== OpenAI クライアント（Manus APIプロキシ経由） =====
try:
    from openai import OpenAI
    ai_client = OpenAI()  # OPENAI_API_KEY と base_url は環境変数から自動取得
    AI_AVAILABLE = True
    print("[INFO] Manus AI client initialized (gpt-4.1-mini / gemini-2.5-flash)")
except Exception as e:
    print(f"[WARNING] OpenAI client init failed: {e}")
    AI_AVAILABLE = False


# ===== メインページ配信 =====
@app.route('/')
def index():
    html_path = os.path.join(os.path.dirname(__file__), 'index.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        return f.read(), 200, {'Content-Type': 'text/html; charset=utf-8'}


# ===== Excel生成 =====
@app.route('/generate-maisoku', methods=['POST'])
def api_generate_maisoku():
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({"error": "No data provided"}), 400

        file_data = generate_maisoku(data)

        prop_name = data.get('prop_name', '物件')
        floor_str = data.get('floor', '')
        safe_name = prop_name.replace(' ', '_').replace('/', '_')[:20]
        safe_floor = floor_str.replace(' ', '_').replace('/', '_')[:10]
        date_str = datetime.now().strftime('%Y%m%d')
        if safe_floor:
            filename = f"{safe_name}_{safe_floor}_マイソク_{date_str}.xlsx"
        else:
            filename = f"{safe_name}_マイソク_{date_str}.xlsx"

        return send_file(
            io.BytesIO(file_data),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


# ===== AI：アピールポイント・市場分析生成 =====
@app.route('/ai-generate', methods=['POST'])
def api_ai_generate():
    if not AI_AVAILABLE:
        return jsonify({"error": "AI client not available"}), 503

    try:
        data   = request.get_json(force=True)
        prompt = data.get('prompt', '')
        model  = data.get('model', 'gpt-4.1-mini')

        if not prompt:
            return jsonify({"error": "prompt is required"}), 400

        resp = ai_client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=4096,
            temperature=0.7
        )

        text = resp.choices[0].message.content or ''
        return jsonify({"text": text})

    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


# ===== AI：販売図面テキスト解析 =====
@app.route('/ai-maisoku-parse', methods=['POST'])
def api_ai_maisoku_parse():
    if not AI_AVAILABLE:
        return jsonify({"error": "AI client not available"}), 503

    try:
        data      = request.get_json(force=True)
        mode      = data.get('mode', 'text')
        content   = data.get('content', '')
        mime_type = data.get('mime_type', '')
        model     = data.get('model', 'gpt-4.1-mini')

        SYSTEM_PROMPT = """あなたは不動産の販売図面（マイソク）から物件情報を抽出するAIです。
以下のJSON形式で情報を返してください。不明な項目は空文字にしてください。
{
  "name": "物件名",
  "floor": "階数（数字のみ）",
  "area": "専有面積（数字のみ、㎡不要）",
  "balcony": "バルコニー面積（数字のみ、不明なら空）",
  "price": "販売価格（万円単位の数字のみ）",
  "address": "所在地",
  "station": "最寄り駅と徒歩分数",
  "structure": "構造（RC造など）",
  "total_units": "総戸数（数字のみ）",
  "floors_total": "地上階数（数字のみ）",
  "built_date": "築年月（例：2014年1月）",
  "mgmt_company": "管理会社",
  "mgmt_fee": "管理費（円/月、数字のみ）",
  "repair_fund": "修繕積立金（円/月、数字のみ）",
  "other_fee": "その他費用（円/月、数字のみ、なければ0）",
  "rental": "賃料（円/月、数字のみ）",
  "status": "現況（賃貸中など）",
  "delivery": "引渡日",
  "washer": "室内洗濯機置場（有/無）",
  "landrights": "土地権利（所有権など）",
  "notes": "その他特記事項"
}
JSONのみを返してください。"""

        if mode == 'text':
            messages = [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": f"以下の物件情報テキストから情報を抽出してください:\n\n{content}"}
            ]
        elif mode in ('image', 'pdf'):
            if ('gpt-4' in model or 'gemini' in model) and mime_type.startswith('image/'):
                messages = [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": [
                        {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{content}"}},
                        {"type": "text", "text": "この販売図面から物件情報を抽出してください。"}
                    ]}
                ]
            else:
                messages = [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": f"以下はファイル（{mime_type}）の内容です。物件情報を抽出してください:\n\n{content[:2000]}"}
                ]
        else:
            return jsonify({"error": f"Unknown mode: {mode}"}), 400

        resp = ai_client.chat.completions.create(
            model=model,
            messages=messages,
            max_tokens=1500,
            temperature=0
        )

        text = resp.choices[0].message.content or ''

        # JSONパース
        json_str = re.sub(r'```json\n?', '', text)
        json_str = re.sub(r'```\n?', '', json_str).strip()
        try:
            result = json.loads(json_str)
        except Exception:
            m = re.search(r'\{[\s\S]+\}', json_str)
            if m:
                try:
                    result = json.loads(m.group(0))
                except Exception:
                    result = {"raw": text}
            else:
                result = {"raw": text}

        return jsonify(result)

    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


# ===== ヘルスチェック =====
@app.route('/health', methods=['GET'])
def health():
    return jsonify({
        "status": "ok",
        "ai_available": AI_AVAILABLE,
        "ai_provider": "Manus OpenAI Proxy (gpt-4.1-mini / gemini-2.5-flash)",
        "version": "9.0"
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    print(f"\n{'='*50}")
    print(f"  不動産投資分析レポート v9.0 起動中")
    print(f"  URL: http://localhost:{port}/")
    print(f"  AI: {'✅ 利用可能' if AI_AVAILABLE else '❌ 利用不可'}")
    print(f"{'='*50}\n")
    app.run(host='0.0.0.0', port=port, debug=False)
