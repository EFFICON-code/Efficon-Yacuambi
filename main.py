import os
import json
import traceback
from flask import Flask, request, jsonify
from flask_cors import CORS
from google import genai
from google.genai import types

# -------------------- App base --------------------
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# -------------------- Config ----------------------
ENTITY_NAME = os.environ.get("EFFICON_ENTITY_NAME", "ENTIDAD-NO-SET")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# Actualizado por defecto a gemini-2.5-pro para máxima compatibilidad
MODEL_ID = os.environ.get("EFFICON_MODEL", "gemini-2.5-pro")

valid_tokens_str = os.environ.get("EFFICON_TOKENS", "")
VALID_TOKENS = [token.strip() for token in valid_tokens_str.split(',') if token.strip()]

# Inicializar cliente de Gemini con el nuevo SDK
client = None
if GEMINI_API_KEY:
    client = genai.Client(api_key=GEMINI_API_KEY)

# -------------------- Health ----------------------
@app.get("/healthz")
def healthz():
    """Endpoint para verificar el estado de la aplicación."""
    return jsonify(status="ok", entity=ENTITY_NAME), 200

# -------------------- Endpoint Principal (Gemini) ---------------------
@app.post("/chatgpt")
def procesar_prompt():
    data = request.get_json(silent=True) or {}
    prompt = data.get("prompt", "")
    
    # 1. Obtener y validar el token
    token_recv = (request.headers.get("X-EFFICON-TOKEN", "") or data.get("token", "")).strip()

    if VALID_TOKENS and token_recv not in VALID_TOKENS:
        return jsonify(error="Auth failed: invalid token"), 403

    if not prompt:
        return jsonify(error="prompt requerido"), 400

    # Si falta la API Key, lanza el famoso Error 500
    if not GEMINI_API_KEY or not client:
        return jsonify(error="GEMINI_API_KEY not configured in environment"), 500

    try:
        # Generar contenido usando la nueva librería genai
        response = client.models.generate_content(
            model=MODEL_ID,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2, # Respuestas precisas y formales
            ),
        )
        
        # Extraer el texto limpio
        text = response.text.strip()
        
        # Devolver el formato exacto que Excel espera
        return jsonify(ok=True, entity=ENTITY_NAME, answer=text), 200

    except Exception as e:
        # Imprimir el error exacto en los logs de Railway
        print("========== ERROR DE GEMINI ==========", flush=True)
        traceback.print_exc()
        print("=====================================", flush=True)
        return jsonify(error="Gemini request failed", details=str(e)), 502
                       
# -------------------- Run local -------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)