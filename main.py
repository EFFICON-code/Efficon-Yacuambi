import os
import json
import requests
from flask import Flask, request, jsonify, send_from_directory, abort
from flask_cors import CORS

# -------------------- App base --------------------
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# -------------------- Config ----------------------
ENTITY_NAME = os.environ.get("EFFICON_ENTITY_NAME", "ENTIDAD-NO-SET")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
MODEL_ID = os.environ.get("EFFICON_MODEL", "gpt-4o")  # <-- puedes cambiarlo en Railway sin tocar código

# Rutas y archivos de módulos
MODULE_DIR = os.path.join(os.getcwd(), "modules", "pack")
MANIFEST_PATH = os.path.join(MODULE_DIR, "manifest.json")

# -------------------- Health ----------------------
@app.get("/healthz")
def healthz():
    has_manifest = os.path.isfile(MANIFEST_PATH)
    return jsonify(status="ok", entity=ENTITY_NAME, manifest=has_manifest), 200

# -------------------- Módulos ---------------------
@app.get("/modules/manifest")
def get_manifest():
    if not os.path.isfile(MANIFEST_PATH):
        return jsonify(error="manifest not found"), 404
    with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
        mf = json.load(f)
    # fuerza el nombre de entidad del entorno
    mf["entity"] = ENTITY_NAME
    return jsonify(mf), 200

@app.get("/modules/download/<path:fname>")
def download_bas(fname):
    if not fname.lower().endswith(".bas"):
        abort(400)
    fullpath = os.path.join(MODULE_DIR, fname)
    if not os.path.isfile(fullpath):
        abort(404)
    return send_from_directory(MODULE_DIR, fname, as_attachment=True)

# -------------------- ChatGPT (con token opcional) ---------------------
@app.post("/chatgpt")
def chatgpt():
    data = request.get_json(silent=True) or {}
    prompt = data.get("prompt", "")

    # Token opcional: si EFFICON_TOKEN existe, lo exigimos
    token_expected = os.environ.get("EFFICON_TOKEN", "").strip()
    token_recv = (request.headers.get("X-EFFICON-TOKEN", "") or data.get("token", "")).strip()
    if token_expected and token_recv != token_expected:
        return jsonify(error="Auth failed: invalid token"), 403

    if not prompt:
        return jsonify(error="prompt requerido"), 400

    if not OPENAI_API_KEY:
        return jsonify(error="OPENAI_API_KEY not configured"), 500

    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": MODEL_ID,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.2
    }

    try:
        r = requests.post("https://api.openai.com/v1/chat/completions",
                          headers=headers, json=payload, timeout=120)
        r.raise_for_status()
        out = r.json()
        text = out.get("choices", [{}])[0].get("message", {}).get("content", "").strip()
        return jsonify(ok=True, entity=ENTITY_NAME, answer=text), 200
    except requests.exceptions.RequestException as e:
        status = getattr(e.response, "status_code", 502)
        return jsonify(error="OpenAI request failed",
                       details=str(e),
                       body=getattr(e.response, "text", "")), status
# -------------------- Run local -------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
    
