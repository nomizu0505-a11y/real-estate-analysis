import os
import requests
from flask import Flask, request, jsonify, send_from_directory

app = Flask(__name__, static_folder="public", static_url_path="")

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")


@app.route("/api/claude", methods=["POST"])
def claude_proxy():
    if not ANTHROPIC_API_KEY:
        return jsonify({"error": "APIキーが設定されていません"}), 500
    try:
        payload = request.get_json()
        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "Content-Type": "application/json",
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
            },
            json=payload,
            timeout=120,
        )
        return jsonify(response.json()), response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/")
def index():
    return send_from_directory("public", "index.html")


@app.route("/<path:path>")
def static_files(path):
    return send_from_directory("public", path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
