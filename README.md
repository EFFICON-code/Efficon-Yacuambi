# EFFICON Yacuambi

Servicio Flask que expone el endpoint `/chatgpt` y utiliza Gemini para generar respuestas.

## Configuración

Configure estas variables de entorno en Railway:

```ini
GEMINI_API_KEY=clave-de-gemini
EFFICON_MODEL=gemini-2.5-pro
EFFICON_ENTITY_NAME=Cuerpo de Bomberos de Loja
EFFICON_EXPIRED_ENTITIES=Cuerpo de Bomberos de Loja
EFFICON_TOKENS=token-opcional
```

`EFFICON_ENTITY_NAME` se conserva para las respuestas y el diagnóstico existentes, pero no participa en el bloqueo por contrato. El control usa el campo `entity` de cada solicitud a `/chatgpt`, porque varias entidades comparten el mismo backend.

`EFFICON_EXPIRED_ENTITIES` admite nombres separados por comas o una lista JSON:

```ini
EFFICON_EXPIRED_ENTITIES=["Cuerpo de Bomberos de Loja", "Cuerpo de Bomberos de Catamayo"]
```

La comparación ignora mayúsculas, minúsculas, tildes, espacios duplicados y espacios en los extremos. Una variable vacía o inexistente no bloquea ninguna entidad. Si una solicitud no contiene `entity`, continúa normalmente para conservar compatibilidad. Para renovar una entidad, elimínela de `EFFICON_EXPIRED_ENTITIES` en Railway y guarde nuevamente las variables.

Cuando la entidad de la solicitud está vencida, `/chatgpt` responde con HTTP `403`, incluye `contract_expired: true` y no llama a Gemini:

```json
{
  "entity": "Cuerpo de Bomberos de Loja",
  "contract_expired": true,
  "error": "contract_expired",
  "message": "Se ha culminado el plazo del contrato..."
}
```

## Pruebas

```shell
python -m pip install -r requirements-dev.txt
python -m pytest
```
