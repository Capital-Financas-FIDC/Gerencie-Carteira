"""Testes da ponte de input bidirecional via stdin."""
import io
import json

import pytest

from input_bridge import CancelExecucao, request_input


def _set_stdin(monkeypatch, texto):
    monkeypatch.setattr("sys.stdin", io.StringIO(texto))


def test_request_input_retorna_mapping(monkeypatch, capsys):
    _set_stdin(monkeypatch, json.dumps({"mapping": {"123": "Ana"}}) + "\n")
    resp = request_input("input.gerentes.needed", {"orfaos": []})
    assert resp == {"mapping": {"123": "Ana"}}
    # o request foi emitido como 1 evento JSON Lines
    eventos = [json.loads(l) for l in capsys.readouterr().out.splitlines()]
    assert eventos[0]["step"] == "input.gerentes.needed"
    assert eventos[0]["data"] == {"orfaos": []}


def test_request_input_cancel_sentinela(monkeypatch, capsys):
    _set_stdin(monkeypatch, json.dumps({"cancel": True}) + "\n")
    with pytest.raises(CancelExecucao):
        request_input("input.gerentes.needed", {})


def test_request_input_eof_cancela(monkeypatch, capsys):
    _set_stdin(monkeypatch, "")  # EOF imediato
    with pytest.raises(CancelExecucao):
        request_input("input.gerentes.needed", {})


def test_request_input_ignora_linha_invalida(monkeypatch, capsys):
    _set_stdin(monkeypatch, "nao-e-json\n" + json.dumps({"mapping": {}}) + "\n")
    resp = request_input("input.gerentes.needed", {})
    assert resp == {"mapping": {}}
    eventos = [json.loads(l) for l in capsys.readouterr().out.splitlines()]
    assert any(e.get("step") == "input.gerentes.needed.invalid" for e in eventos)
