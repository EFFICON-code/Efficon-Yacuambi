import os
import json
from flask import Flask, request, jsonify
from flask_cors import CORS
import google.generativeai as genai

# -------------------- App base --------------------
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# -------------------- Config ----------------------
ENTITY_NAME = os.environ.get("EFFICON_ENTITY_NAME", "ENTIDAD-NO-SET")
# Cambiamos OPENAI_API_KEY por GEMINI_API_KEY
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
# Cambiamos el modelo por defecto al más potente para textos largos
MODEL_ID = os.environ.get("EFFICON_MODEL", "gemini-1.5-pro")

# Obtener la lista de tokens válidos desde el entorno
valid_tokens_str = os.environ.get("EFFICON_TOKENS", "")
VALID_TOKENS = [token.strip() for token in valid_tokens_str.split(',') if token.strip()]

# Configurar el cliente de Google Gemini globalmente
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

# -------------------- Health ----------------------
@app.get("/healthz")
def healthz():
    """
    Endpoint para verificar el estado de la aplicación.
    """
    return jsonify(status="ok", entity=ENTITY_NAME), 200

# -------------------- Endpoint Principal (Gemini) ---------------------
# Mantenemos la ruta /chatgpt para NO romper tu código VBA en Excel
@app.post("/chatgpt")
def procesar_prompt():
    """
    Endpoint principal para procesar prompts con Gemini 1.5 Pro.
    Requiere autenticación con un token.
    """
    data = request.get_json(silent=True) or {}
    prompt = data.get("prompt", "")
    
    # 1. Obtener el token del request
    token_recv = (request.headers.get("X-EFFICON-TOKEN", "") or data.get("token", "")).strip()

    # 2. Validar el token
    if VALID_TOKENS and token_recv not in VALID_TOKENS:
        return jsonify(error="Auth failed: invalid token"), 403

    if not prompt:
        return jsonify(error="prompt requerido"), 400

    if not GEMINI_API_KEY:
        return jsonify(error="GEMINI_API_KEY not configured in environment"), 500

    try:
        # Inicializar el modelo
        model = genai.GenerativeModel(MODEL_ID)
        
        # Configuramos la temperatura baja (0.2) para respuestas formales y precisas
        config = genai.types.GenerationConfig(temperature=0.2)
        
        # Generar la respuesta
        response = model.generate_content(prompt, generation_config=config)
        
        # Extraer el texto limpio
        text = response.text.strip()
        
        # IMPORTANTE: Devolvemos el mismo JSON que esperaba tu Excel
        # Si tu Excel espera "choices" y "message" (formato OpenAI viejo), lo emulamos.
        # En tu código de Excel original pedías: jsonResp("choices")(1)("message")("content")
        # Aquí te devuelvo exactamente ese formato para que tu macro funcione al instante.
        emulated_response = {
            "choices": [
                {
                    "message": {
                        "content": text
                    }
                }
            ]
        }
        
        # Ojo: Devuelvo tu formato base ("ok", "entity", "answer") y también "choices" para Excel
        return jsonify(
            ok=True, 
            entity=ENTITY_NAME, 
            answer=text,
            choices=emulated_response["choices"]
        ), 200

    except Exception as e:
        return jsonify(error="Gemini request failed", details=str(e)), 502
                       
# -------------------- Run local -------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)