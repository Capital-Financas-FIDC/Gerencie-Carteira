"""
Ponte de input bidirecional (UI -> Python) via stdin.

O protocolo de saida (Python -> UI) continua sendo JSON Lines em stdout
(`log_emitter.emit`). Para PEDIR um dado ao usuario em runtime, o Python:

  1. emite UM evento JSON Lines de request via `emit()` (contrato R1);
  2. BLOQUEIA lendo exatamente UMA linha JSON do stdin (resposta da UI).

A UI responde com 1 linha JSON:
    {"mapping": {"<cnpj>": "<gerente>", ...}}   -> confirmacao
    {"cancel": true}                              -> fechamento/cancelamento

Uma linha vazia / EOF (processo sendo encerrado) tambem e tratada como
cancelamento. Linhas nao-parseaveis sao registradas como warning e a leitura
continua, NUNCA derrubando o stream.
"""

import json
import sys
from typing import Any, Optional

from log_emitter import emit


class CancelExecucao(Exception):
    """Usuario cancelou/fechou durante um request de input em runtime."""


def request_input(step: str, payload: dict[str, Any], *, msg: str = "Aguardando entrada do usuario") -> dict:
    """
    Emite um evento de request e bloqueia ate a UI responder 1 linha JSON.

    Retorna o dict de resposta (ex: {"mapping": {...}}). Levanta
    `CancelExecucao` se a UI responder {"cancel": true} ou se o stdin
    encerrar (EOF) — caso em que o processo esta sendo finalizado.
    """
    emit("step", msg, step=step, data=payload)

    # Le bytes crus e decodifica UTF-8 explicitamente. Electron envia UTF-8
    # via child.stdin.write(), mas o sys.stdin em modo texto do PyInstaller
    # frozen pode ignorar PYTHONIOENCODING e cair em cp1252 (Windows),
    # corrompendo caracteres acentuados (ex: "ADMINISTRAÇÃO" -> "ADMINISTRAÃ‡ÃƒO").
    stdin = sys.stdin.buffer
    while True:
        linha_bytes = stdin.readline()
        if not linha_bytes:  # EOF — processo sendo encerrado
            raise CancelExecucao("stdin fechado (EOF) durante request de input")
        try:
            linha = linha_bytes.decode("utf-8").strip()
        except UnicodeDecodeError:
            emit("warning", "Resposta de input com bytes invalidos (ignorada)",
                 step=f"{step}.invalid")
            continue
        if not linha:
            continue
        try:
            resposta = json.loads(linha)
        except (json.JSONDecodeError, ValueError):
            emit("warning", "Resposta de input invalida (ignorada)",
                 step=f"{step}.invalid")
            continue
        if not isinstance(resposta, dict):
            emit("warning", "Resposta de input nao e objeto (ignorada)",
                 step=f"{step}.invalid")
            continue
        if resposta.get("cancel") is True:
            raise CancelExecucao("Usuario cancelou o request de input")
        return resposta


def peek_cancel(timeout: Optional[float] = None) -> bool:
    """
    Best-effort: verifica se ha um sentinela de cancelamento pendente no
    stdin sem bloquear indefinidamente. Usado em pontos seguros do pipeline.
    Em Windows o select() nao funciona em pipes; por isso o cancelamento
    autoritativo e garantido pelo supervisor (Electron) — esta funcao e
    apenas um atalho cooperativo opcional.
    """
    return False


if __name__ == "__main__":
    # Smoke manual: emite request e ecoa a resposta.
    try:
        r = request_input("input.smoke", {"orfaos": [{"cnpj": "00", "razao_social": "X"}]})
        emit("info", f"resposta recebida: {r}", step="input.smoke.ok")
    except CancelExecucao as e:
        emit("warning", f"cancelado: {e}", step="input.smoke.cancel")
