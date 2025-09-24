import os
import json
import requests
from flask import Flask, request, jsonify, abort
from flask_cors import CORS

# -------------------- App base --------------------
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# -------------------- Config ----------------------
ENTITY_NAME = os.environ.get("EFFICON_ENTITY_NAME", "ENTIDAD-NO-SET")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
MODEL_ID = os.environ.get("EFFICON_MODEL", "gpt-4o")

# Obtener la lista de tokens válidos desde el entorno
# Si la variable no existe, se usa una cadena vacía
# Se divide la cadena por comas para crear una lista
# El .strip() limpia los espacios en blanco de cada token
valid_tokens_str = os.environ.get("EFFICON_TOKENS", "")
VALID_TOKENS = [token.strip() for token in valid_tokens_str.split(',') if token.strip()]

# -------------------- Health ----------------------
@app.get("/healthz")
def healthz():
    """
    Endpoint para verificar el estado de la aplicación.
    """
    return jsonify(status="ok", entity=ENTITY_NAME), 200

# -------------------- ChatGPT (con múltiples tokens) ---------------------
@app.post("/chatgpt")
def chatgpt():
    """
    Endpoint principal para procesar prompts con la API de OpenAI.
    Requiere autenticación con un token.
    """
    data = request.get_json(silent=True) or {}
    prompt = data.get("prompt", "")
    
    # 1. Obtener el token del request
    token_recv = (request.headers.get("X-EFFICON-TOKEN", "") or data.get("token", "")).strip()

    # 2. Validar el token
    # Si la lista de tokens válidos no está vacía Y el token recibido NO está en la lista
    if VALID_TOKENS and token_recv not in VALID_TOKENS:
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
