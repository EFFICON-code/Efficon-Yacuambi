import os
import json
import traceback
import unicodedata
from flask import Flask, request, jsonify
from flask_cors import CORS
from google import genai
from google.genai import types


CONTRACT_EXPIRED_MESSAGE = (
    "Se ha culminado el plazo del contrato. Contacte con PROESTRATEGIA para "
    "renovar el contrato.\n\n"
    "Ing. Carlos Antonio Salinas Coronel\n"
    "Celular: 0967314512\n\n"
    "Gracias por utilizar nuestros servicios."
)


def normalize_entity_name(value):
    """Normaliza una entidad para comparaciones tolerantes a formato."""
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(text.casefold().split())


def parse_expired_entities(raw_value):
    """Acepta una lista JSON o una lista de nombres separada por comas."""
    if not raw_value or not str(raw_value).strip():
        return []

    raw_text = str(raw_value).strip()
    try:
        parsed = json.loads(raw_text)
    except (json.JSONDecodeError, TypeError):
        parsed = raw_text.split(",")

    if not isinstance(parsed, list):
        parsed = raw_text.split(",")

    return [item.strip() for item in parsed if isinstance(item, str) and item.strip()]

# -------------------- App base --------------------
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

# -------------------- Config ----------------------
ENTITY_NAME = os.environ.get("EFFICON_ENTITY_NAME", "ENTIDAD-NO-SET")
EXPIRED_ENTITIES = parse_expired_entities(os.environ.get("EFFICON_EXPIRED_ENTITIES", ""))

# [DIAGNOSTIC] Logs temporales para depurar EFFICON_EXPIRED_ENTITIES
print(
    "[STARTUP] EFFICON_EXPIRED_ENTITIES raw value:",
    os.environ.get("EFFICON_EXPIRED_ENTITIES", "NO_SET"),
    flush=True,
)
print("[STARTUP] EXPIRED_ENTITIES parsed list:", EXPIRED_ENTITIES, flush=True)

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# Actualizado por defecto a gemini-2.5-pro para máxima compatibilidad
MODEL_ID = os.environ.get("EFFICON_MODEL", "gemini-2.5-pro")

valid_tokens_str = os.environ.get("EFFICON_TOKENS", "")
VALID_TOKENS = [token.strip() for token in valid_tokens_str.split(',') if token.strip()]

# Inicializar cliente de Gemini con el nuevo SDK
client = None
if GEMINI_API_KEY:
    client = genai.Client(api_key=GEMINI_API_KEY)


def is_contract_expired(entity_name):
    """Valida la entidad recibida en una solicitud contra la lista configurada."""
    normalized_entity = normalize_entity_name(entity_name)
    normalized_expired = {
        normalize_entity_name(entity)
        for entity in EXPIRED_ENTITIES
        if normalize_entity_name(entity)
    }
    return bool(normalized_entity and normalized_entity in normalized_expired)

# -------------------- Health ----------------------
@app.get("/healthz")
def healthz():
    """Endpoint para verificar el estado de la aplicación."""
    return jsonify(status="ok", entity=ENTITY_NAME), 200

# -------------------- Endpoint Principal (Gemini) ---------------------
@app.post("/chatgpt")
def procesar_prompt():
    data = request.get_json(silent=True) or {}

    # [REQUEST-FULL] Logs de diagnóstico para identificar qué datos/contexto
    # envía VBA que permitan mapear la solicitud a una entidad, dado que VBA
    # no envía el campo "entity" en el JSON.
    try:
        raw_body = request.get_data(as_text=True)
    except Exception as exc:
        raw_body = f"<error reading raw body: {exc}>"
    print("[REQUEST-FULL] Raw JSON body:", raw_body, flush=True)
    print("[REQUEST-FULL] Parsed JSON data:", data, flush=True)
    print(
        "[REQUEST-FULL] Headers received:",
        dict(request.headers),
        flush=True,
    )
    print(
        "[REQUEST-FULL] Query parameters:",
        request.args.to_dict(flat=False),
        flush=True,
    )
    print(
        "[REQUEST-FULL] Full request dump:",
        {
            "method": request.method,
            "path": request.path,
            "full_path": request.full_path,
            "url": request.url,
            "remote_addr": request.remote_addr,
            "content_type": request.content_type,
            "content_length": request.content_length,
            "form": request.form.to_dict(flat=False),
            "cookies": request.cookies,
        },
        flush=True,
    )

    request_entity = data.get("entity", "")

    # [DIAGNOSTIC] Logs temporales para depurar EFFICON_EXPIRED_ENTITIES
    print(
        "[REQUEST] Received entity:",
        request_entity if request_entity else "empty_string",
        flush=True,
    )
    print("[REQUEST] Normalized entity:", normalize_entity_name(request_entity), flush=True)
    print(
        "[REQUEST] Normalized EXPIRED_ENTITIES set:",
        {
            normalize_entity_name(entity)
            for entity in EXPIRED_ENTITIES
            if normalize_entity_name(entity)
        },
        flush=True,
    )
    print(
        "[REQUEST] is_contract_expired() result:",
        is_contract_expired(request_entity),
        flush=True,
    )

    # Se valida antes de autenticar o invocar cualquier servicio con costo.
    # En este backend compartido, la entidad se obtiene de cada solicitud.
    if is_contract_expired(request_entity):
        return jsonify(
            error="contract_expired",
            contract_expired=True,
            entity=request_entity,
            message=CONTRACT_EXPIRED_MESSAGE,
        ), 403

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
