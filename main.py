import os
import json
import logging
import traceback
import unicodedata
from flask import Flask, request, jsonify
from flask_cors import CORS
from google import genai
from google.genai import types
from openai import APIConnectionError, APITimeoutError, OpenAI


CONTRACT_EXPIRED_MESSAGE = (
    "Se ha culminado el plazo del contrato. Contacte con PROESTRATEGIA para "
    "renovar el contrato.\n\n"
    "Ing. Carlos Antonio Salinas Coronel\n"
    "Celular: 0967314512\n\n"
    "Gracias por utilizar nuestros servicios."
)

logger = logging.getLogger(__name__)


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
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# Actualizado por defecto a gemini-2.5-pro para máxima compatibilidad
MODEL_ID = os.environ.get("EFFICON_MODEL", "gemini-2.5-pro")

LLM_PRIMARY_PROVIDER = os.environ.get("LLM_PRIMARY_PROVIDER", "openai").strip().lower()
LLM_FALLBACK_PROVIDER = os.environ.get("LLM_FALLBACK_PROVIDER", "gemini").strip().lower()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-5.6-luna")

valid_tokens_str = os.environ.get("EFFICON_TOKENS", "")
VALID_TOKENS = [token.strip() for token in valid_tokens_str.split(',') if token.strip()]

# Inicializar cliente de Gemini con el nuevo SDK
client = None
if GEMINI_API_KEY:
    client = genai.Client(api_key=GEMINI_API_KEY)

# El SDK de OpenAI aplica retries automáticos por defecto. Se desactivan para
# activar inmediatamente el fallback controlado por esta aplicación.
openai_client = None
if OPENAI_API_KEY:
    openai_client = OpenAI(
        api_key=OPENAI_API_KEY,
        timeout=120.0,
        max_retries=0,
    )


class LLMConfigurationError(RuntimeError):
    """Configuración incompleta o inválida de proveedores LLM."""


class LLMResponseError(RuntimeError):
    """Respuesta del proveedor sin texto utilizable."""


def _openai_error_code(error):
    """Obtiene un código seguro de un error del SDK sin registrar su mensaje."""
    code = getattr(error, "code", None)
    body = getattr(error, "body", None)
    if isinstance(body, dict):
        code = body.get("code", code)
        nested_error = body.get("error")
        if isinstance(nested_error, dict):
            code = nested_error.get("code", code)
    return str(code) if code else None


def openai_fallback_reason(error):
    """Clasifica exclusivamente fallos recuperables de OpenAI."""
    if isinstance(error, APITimeoutError):
        return "timeout"
    if isinstance(error, APIConnectionError):
        return "connection_error"

    status_code = getattr(error, "status_code", None)
    error_code = _openai_error_code(error)
    if error_code in {
        "insufficient_quota",
        "credit_balance_exhausted",
        "organization_spend_limit_exceeded",
        "project_spend_limit_exceeded",
        "organization_usage_limit_exceeded",
    }:
        return error_code
    if status_code == 429:
        return error_code or "rate_limit"
    if status_code in {500, 502, 503, 504}:
        return f"http_{status_code}"
    return None


def generate_with_openai(prompt):
    """Genera texto con OpenAI mediante Responses API."""
    if not OPENAI_API_KEY or not openai_client:
        raise LLMConfigurationError("OPENAI_API_KEY not configured in environment")

    response = openai_client.responses.create(
        model=OPENAI_MODEL,
        input=prompt,
    )
    text = (response.output_text or "").strip()
    if not text:
        raise LLMResponseError("OpenAI response did not contain output text")

    logger.info("LLM provider: openai; model: %s", OPENAI_MODEL)
    return text


def generate_with_gemini(prompt):
    """Genera texto con la integración Gemini existente."""
    if not GEMINI_API_KEY or not client:
        raise LLMConfigurationError("GEMINI_API_KEY not configured in environment")

    response = client.models.generate_content(
        model=MODEL_ID,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.2,
        ),
    )
    text = (response.text or "").strip()
    if not text:
        raise LLMResponseError("Gemini response did not contain text")

    logger.info("LLM provider: gemini; model: %s", MODEL_ID)
    return text


def generate_llm_text(prompt):
    """Enruta al proveedor principal y aplica fallback selectivo a Gemini."""
    if LLM_PRIMARY_PROVIDER == "gemini":
        return generate_with_gemini(prompt)
    if LLM_PRIMARY_PROVIDER != "openai":
        raise LLMConfigurationError(
            f"Unsupported LLM_PRIMARY_PROVIDER: {LLM_PRIMARY_PROVIDER}"
        )

    if not OPENAI_API_KEY or not openai_client:
        logger.warning("OpenAI no configurado. Utilizando Gemini.")
        return generate_with_gemini(prompt)

    try:
        return generate_with_openai(prompt)
    except Exception as error:
        fallback_reason = openai_fallback_reason(error)
        if not fallback_reason or LLM_FALLBACK_PROVIDER != "gemini":
            logger.exception("OpenAI request failed without eligible fallback")
            raise

        logger.warning(
            "LLM primary provider unavailable; fallback provider: gemini; "
            "fallback reason: %s",
            fallback_reason,
        )
        try:
            return generate_with_gemini(prompt)
        except Exception:
            logger.exception("Fallback provider failed: gemini")
            raise


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
    request_entity = data.get("entity", "")

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

    try:
        text = generate_llm_text(prompt)
        
        # Devolver el formato exacto que Excel espera
        return jsonify(ok=True, entity=ENTITY_NAME, answer=text), 200

    except LLMConfigurationError as e:
        logger.error("LLM configuration error: %s", e)
        return jsonify(error=str(e)), 500

    except Exception as e:
        print("========== ERROR DE LLM ==========", flush=True)
        traceback.print_exc()
        print("==================================", flush=True)
        return jsonify(error="Gemini request failed", details=str(e)), 502
                       
# -------------------- Run local -------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
