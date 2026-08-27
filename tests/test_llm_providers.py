import logging
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

import main


class SimulatedAPIError(Exception):
    def __init__(self, status_code=None, code=None):
        super().__init__(f"simulated provider error {status_code or code}")
        self.status_code = status_code
        self.body = {"code": code} if code else {}


@pytest.fixture
def provider_setup(monkeypatch):
    openai_create = MagicMock(
        return_value=SimpleNamespace(output_text="Respuesta de OpenAI")
    )
    gemini_generate = MagicMock(
        return_value=SimpleNamespace(text="Respuesta de Gemini")
    )

    monkeypatch.setattr(
        main,
        "openai_client",
        SimpleNamespace(responses=SimpleNamespace(create=openai_create)),
    )
    monkeypatch.setattr(
        main,
        "client",
        SimpleNamespace(models=SimpleNamespace(generate_content=gemini_generate)),
    )
    monkeypatch.setattr(main, "OPENAI_API_KEY", "test-openai-key")
    monkeypatch.setattr(main, "GEMINI_API_KEY", "test-gemini-key")
    monkeypatch.setattr(main, "OPENAI_MODEL", "gpt-5.6-luna")
    monkeypatch.setattr(main, "MODEL_ID", "gemini-2.5-pro")
    monkeypatch.setattr(main, "LLM_PRIMARY_PROVIDER", "openai")
    monkeypatch.setattr(main, "LLM_FALLBACK_PROVIDER", "gemini")

    return openai_create, gemini_generate


def test_openai_is_used_when_configured(provider_setup):
    openai_create, gemini_generate = provider_setup

    result = main.generate_llm_text("Prompt completo")

    assert result == "Respuesta de OpenAI"
    openai_create.assert_called_once()
    gemini_generate.assert_not_called()


def test_openai_luna_model_is_sent(provider_setup):
    openai_create, _ = provider_setup

    main.generate_llm_text("Prompt completo")

    assert openai_create.call_args.kwargs == {
        "model": "gpt-5.6-luna",
        "input": "Prompt completo",
    }


def test_missing_openai_key_keeps_gemini_running(provider_setup, monkeypatch):
    openai_create, gemini_generate = provider_setup
    monkeypatch.setattr(main, "OPENAI_API_KEY", "")
    monkeypatch.setattr(main, "openai_client", None)

    result = main.generate_llm_text("Prompt completo")

    assert result == "Respuesta de Gemini"
    openai_create.assert_not_called()
    gemini_generate.assert_called_once()


def test_insufficient_quota_falls_back_to_gemini(provider_setup):
    openai_create, gemini_generate = provider_setup
    openai_create.side_effect = SimulatedAPIError(
        status_code=429,
        code="insufficient_quota",
    )

    result = main.generate_llm_text("Prompt completo")

    assert result == "Respuesta de Gemini"
    gemini_generate.assert_called_once()


def test_rate_limit_falls_back_to_gemini(provider_setup):
    openai_create, gemini_generate = provider_setup
    openai_create.side_effect = SimulatedAPIError(status_code=429)

    result = main.generate_llm_text("Prompt completo")

    assert result == "Respuesta de Gemini"
    gemini_generate.assert_called_once()


def test_timeout_falls_back_to_gemini(provider_setup, monkeypatch):
    openai_create, gemini_generate = provider_setup

    class SimulatedTimeout(Exception):
        pass

    monkeypatch.setattr(main, "APITimeoutError", SimulatedTimeout)
    openai_create.side_effect = SimulatedTimeout("timeout")

    result = main.generate_llm_text("Prompt completo")

    assert result == "Respuesta de Gemini"
    gemini_generate.assert_called_once()


def test_connection_error_falls_back_to_gemini(provider_setup, monkeypatch):
    openai_create, gemini_generate = provider_setup

    class SimulatedConnectionError(Exception):
        pass

    monkeypatch.setattr(main, "APIConnectionError", SimulatedConnectionError)
    openai_create.side_effect = SimulatedConnectionError("connection")

    result = main.generate_llm_text("Prompt completo")

    assert result == "Respuesta de Gemini"
    gemini_generate.assert_called_once()


@pytest.mark.parametrize("status_code", [500, 502, 503, 504])
def test_openai_5xx_falls_back_to_gemini(provider_setup, status_code):
    openai_create, gemini_generate = provider_setup
    openai_create.side_effect = SimulatedAPIError(status_code=status_code)

    result = main.generate_llm_text("Prompt completo")

    assert result == "Respuesta de Gemini"
    gemini_generate.assert_called_once()


def test_gemini_failure_after_openai_failure_is_controlled_and_logged(
    provider_setup,
    caplog,
):
    openai_create, gemini_generate = provider_setup
    openai_create.side_effect = SimulatedAPIError(status_code=503)
    gemini_generate.side_effect = RuntimeError("Gemini unavailable")

    with caplog.at_level(logging.ERROR), pytest.raises(
        RuntimeError,
        match="Gemini unavailable",
    ):
        main.generate_llm_text("Prompt completo")

    assert "Fallback provider failed: gemini" in caplog.text


def test_endpoint_returns_controlled_error_when_both_providers_fail(
    provider_setup,
    monkeypatch,
    caplog,
):
    openai_create, gemini_generate = provider_setup
    openai_create.side_effect = SimulatedAPIError(status_code=503)
    gemini_generate.side_effect = RuntimeError("Gemini unavailable")
    monkeypatch.setattr(main, "ENTITY_NAME", "Cuerpo de Bomberos de Loja")
    monkeypatch.setattr(main, "EXPIRED_ENTITIES", [])
    monkeypatch.setattr(main, "VALID_TOKENS", [])
    main.app.config.update(TESTING=True)

    with caplog.at_level(logging.ERROR), main.app.test_client() as test_client:
        response = test_client.post("/chatgpt", json={"prompt": "Hola"})

    assert response.status_code == 502
    assert response.get_json() == {
        "error": "Gemini request failed",
        "details": "Gemini unavailable",
    }
    assert "Fallback provider failed: gemini" in caplog.text


def test_gemini_can_run_as_primary_independently(provider_setup, monkeypatch):
    openai_create, gemini_generate = provider_setup
    monkeypatch.setattr(main, "LLM_PRIMARY_PROVIDER", "gemini")

    result = main.generate_llm_text("Prompt completo")

    assert result == "Respuesta de Gemini"
    openai_create.assert_not_called()
    gemini_generate.assert_called_once()


def test_authentication_error_does_not_trigger_fallback(provider_setup):
    openai_create, gemini_generate = provider_setup
    openai_create.side_effect = SimulatedAPIError(status_code=401)

    with pytest.raises(SimulatedAPIError):
        main.generate_llm_text("Prompt completo")

    gemini_generate.assert_not_called()


def test_endpoint_preserves_current_success_contract(
    provider_setup,
    monkeypatch,
):
    monkeypatch.setattr(main, "ENTITY_NAME", "Cuerpo de Bomberos de Loja")
    monkeypatch.setattr(main, "EXPIRED_ENTITIES", [])
    monkeypatch.setattr(main, "VALID_TOKENS", [])
    main.app.config.update(TESTING=True)

    with main.app.test_client() as test_client:
        response = test_client.post("/chatgpt", json={"prompt": "Hola"})

    assert response.status_code == 200
    assert response.get_json() == {
        "ok": True,
        "entity": "Cuerpo de Bomberos de Loja",
        "answer": "Respuesta de OpenAI",
    }


def test_endpoint_uses_gemini_fallback_without_changing_contract(
    provider_setup,
    monkeypatch,
):
    openai_create, gemini_generate = provider_setup
    openai_create.side_effect = SimulatedAPIError(
        status_code=429,
        code="insufficient_quota",
    )
    monkeypatch.setattr(main, "ENTITY_NAME", "Cuerpo de Bomberos de Loja")
    monkeypatch.setattr(main, "EXPIRED_ENTITIES", [])
    monkeypatch.setattr(main, "VALID_TOKENS", [])
    main.app.config.update(TESTING=True)

    with main.app.test_client() as test_client:
        response = test_client.post("/chatgpt", json={"prompt": "Hola"})

    assert response.status_code == 200
    assert response.get_json() == {
        "ok": True,
        "entity": "Cuerpo de Bomberos de Loja",
        "answer": "Respuesta de Gemini",
    }
    gemini_generate.assert_called_once()
