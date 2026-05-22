"""Testes da ponte de input bidirecional via stdin."""
import io
import json

import pytest

from input_bridge import CancelExecucao, request_input


class _FakeStdin:
    """
    Substituto de sys.stdin para testes. `request_input` le bytes crus via
    `sys.stdin.buffer` (fix v4.1.1 — evita mojibake do cp1252 no frozen),
    entao o falso stdin precisa expor um buffer binario.
    """

    def __init__(self, data: bytes):
        self.buffer = io.BytesIO(data)


def _set_stdin(monkeypatch, texto: str) -> None:
    monkeypatch.setattr("sys.stdin", _FakeStdin(texto.encode("utf-8")))


def _set_stdin_bytes(monkeypatch, data: bytes) -> None:
    monkeypatch.setattr("sys.stdin", _FakeStdin(data))


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


def test_request_input_decodifica_utf8_com_acentos(monkeypatch, capsys):
    """A resposta da UI chega como UTF-8 (fix v4.1.1: stdin lido como bytes)."""
    _set_stdin(monkeypatch, json.dumps({"mapping": {"1": "ADMINISTRAÇÃO"}}) + "\n")
    resp = request_input("input.gerentes.needed", {})
    assert resp == {"mapping": {"1": "ADMINISTRAÇÃO"}}


def test_request_input_ignora_bytes_invalidos(monkeypatch, capsys):
    """Bytes nao-UTF-8 viram warning e a leitura continua (nao derruba o stream)."""
    # 1a linha: 0xFF isolado (invalido em UTF-8); 2a linha: resposta valida.
    _set_stdin_bytes(
        monkeypatch,
        b"\xff\n" + json.dumps({"mapping": {}}).encode("utf-8") + b"\n",
    )
    resp = request_input("input.gerentes.needed", {})
    assert resp == {"mapping": {}}
    eventos = [json.loads(l) for l in capsys.readouterr().out.splitlines()]
    assert any(e.get("step") == "input.gerentes.needed.invalid" for e in eventos)
