from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

import main


@pytest.fixture
def app_client(monkeypatch):
    generate_content = MagicMock(
        return_value=SimpleNamespace(text="Respuesta de prueba")
    )
    fake_client = SimpleNamespace(
        models=SimpleNamespace(generate_content=generate_content)
    )

    monkeypatch.setattr(main, "ENTITY_NAME", "Cuerpo de Bomberos de Loja")
    monkeypatch.setattr(main, "EXPIRED_ENTITIES", [])
    monkeypatch.setattr(main, "GEMINI_API_KEY", "test-key")
    monkeypatch.setattr(main, "VALID_TOKENS", [])
    monkeypatch.setattr(main, "client", fake_client)
    main.app.config.update(TESTING=True)

    with main.app.test_client() as test_client:
        yield test_client, generate_content


def test_active_entity_continues_to_gemini(app_client):
    client, generate_content = app_client

    response = client.post(
        "/chatgpt",
        json={"prompt": "Hola", "entity": "Cuerpo de Bomberos de Catamayo"},
    )

    assert response.status_code == 200
    assert response.get_json() == {
        "ok": True,
        "entity": "Cuerpo de Bomberos de Loja",
        "answer": "Respuesta de prueba",
    }
    generate_content.assert_called_once()


def test_expired_entity_is_blocked_before_gemini(app_client, monkeypatch):
    client, generate_content = app_client
    monkeypatch.setattr(main, "EXPIRED_ENTITIES", ["Cuerpo de Bomberos de Loja"])

    response = client.post(
        "/chatgpt",
        json={"prompt": "Hola", "entity": "Cuerpo de Bomberos de Loja"},
    )
    payload = response.get_json()

    assert response.status_code == 403
    assert payload["contract_expired"] is True
    assert payload["message"] == main.CONTRACT_EXPIRED_MESSAGE
    generate_content.assert_not_called()


def test_match_ignores_case_accents_and_spaces(app_client, monkeypatch):
    client, generate_content = app_client
    monkeypatch.setattr(
        main,
        "EXPIRED_ENTITIES",
        ["cuerpo de bomberos de loja"],
    )

    response = client.post(
        "/chatgpt",
        json={
            "prompt": "Hola",
            "entity": "  CUERPO   de BOMBEROS de LÓJA  ",
        },
    )

    assert response.status_code == 403
    assert response.get_json()["contract_expired"] is True
    generate_content.assert_not_called()


def test_expired_fire_department_does_not_block_other_entities(
    app_client,
    monkeypatch,
):
    client, generate_content = app_client
    monkeypatch.setattr(main, "EXPIRED_ENTITIES", ["Cuerpo de Bomberos de Loja"])

    expired_response = client.post(
        "/chatgpt",
        json={
            "prompt": "Hola",
            "entity": "Cuerpo de Bomberos de Loja",
        },
    )
    active_response = client.post(
        "/chatgpt",
        json={
            "prompt": "Hola",
            "entity": "Cuerpo de Bomberos de Catamayo",
        },
    )

    assert expired_response.status_code == 403
    assert expired_response.get_json()["entity"] == "Cuerpo de Bomberos de Loja"
    assert active_response.status_code == 200
    generate_content.assert_called_once()


def test_missing_request_entity_preserves_existing_behavior(app_client):
    client, generate_content = app_client

    response = client.post("/chatgpt", json={"prompt": "Hola"})

    assert response.status_code == 200
    generate_content.assert_called_once()


@pytest.mark.parametrize("raw_value", [None, "", "   "])
def test_empty_or_missing_expired_entities_does_not_block(raw_value):
    assert main.parse_expired_entities(raw_value) == []
    assert main.normalize_entity_name("Cuerpo de Bomberos de Loja") not in {
        main.normalize_entity_name(entity)
        for entity in main.parse_expired_entities(raw_value)
    }


def test_json_expired_entities_format_is_supported():
    entities = main.parse_expired_entities(
        '["Cuerpo de Bomberos de Loja", "Cuerpo de Bomberos de Catamayo"]'
    )

    assert entities == [
        "Cuerpo de Bomberos de Loja",
        "Cuerpo de Bomberos de Catamayo",
    ]


def test_healthz_preserves_existing_response(app_client):
    client, _ = app_client

    response = client.get("/healthz")

    assert response.status_code == 200
    assert response.get_json() == {
        "status": "ok",
        "entity": "Cuerpo de Bomberos de Loja",
    }
