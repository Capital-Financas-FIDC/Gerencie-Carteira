"""
Emissor de eventos JSON Lines para stdout.

Contrato: cada chamada de `emit()` produz EXATAMENTE uma linha JSON valida
em stdout, terminada por \\n e com flush imediato. O processo pai (Electron
main) consome linha a linha via readline.

Formato de evento:
    {
      "level": "info" | "success" | "warning" | "error" | "step",
      "ts":    "2026-04-17T13:58:04.123-03:00",
      "msg":   "Texto humano da mensagem",
      "step":  "outlook.fetch",             # opcional
      "progress": 42,                        # opcional (0..100)
      "data":  { ... }                       # opcional
    }

Timeline em memoria (instrumentacao de performance):
    reset_timer()   — zera o relogio e a lista interna (chamar no boot de main())
    get_timeline()  — retorna copia da lista de eventos com timing
    emit(step=...)  — quando step != None, agrega {step, msg, t_ms, dt_ms} na lista

    A timeline nao altera o contrato JSON Lines: o stdout continua recebendo
    exatamente 1 linha JSON valida por emit(), sem campos extras. O registro de
    timing e puramente em memoria — nenhuma I/O adicional por evento.
"""

import json
import sys
import time
from datetime import datetime, timezone
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Estado interno de timeline (modulo-level; reinicializado por reset_timer)
# ---------------------------------------------------------------------------
_t0: Optional[float] = None       # monotonic no boot (reset_timer)
_t_prev: Optional[float] = None   # monotonic do emit anterior
_TIMELINE: list[dict[str, Any]] = []  # [{step, msg, t_ms, dt_ms}, ...]


def reset_timer() -> None:
    """Zera o relogio monotonic e a timeline em memoria. Chamar no boot de main()."""
    global _t0, _t_prev, _TIMELINE
    _t0 = time.monotonic()
    _t_prev = _t0
    _TIMELINE = []


def get_timeline() -> list[dict[str, Any]]:
    """Retorna copia da timeline acumulada (lista de dicts, ordem de insercao)."""
    return list(_TIMELINE)


def _now_iso() -> str:
    """Timestamp ISO-8601 com timezone local."""
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="milliseconds")


def emit(
    level: str,
    msg: str,
    *,
    step: Optional[str] = None,
    progress: Optional[int] = None,
    data: Optional[dict[str, Any]] = None,
) -> None:
    """Emite 1 evento JSON em stdout (1 linha, flush imediato).

    Quando `step` nao e None, agrega timing em memoria na _TIMELINE:
      t_ms  = ms desde o reset_timer() (ou 0 se timer nao foi inicializado)
      dt_ms = ms desde o emit anterior com step (idem)
    O JSON emitido em stdout NAO e alterado — contrato intacto.
    """
    global _t_prev

    # --- Registro de timing em memoria (sem I/O) ---
    if step is not None:
        now = time.monotonic()
        t0_ref = _t0 if _t0 is not None else now
        t_prev_ref = _t_prev if _t_prev is not None else now
        t_ms = round((now - t0_ref) * 1000)
        dt_ms = round((now - t_prev_ref) * 1000)
        _TIMELINE.append({"step": step, "msg": msg, "t_ms": t_ms, "dt_ms": dt_ms})
        _t_prev = now

    # --- Emissao JSON Lines (contrato intacto) ---
    payload: dict[str, Any] = {"level": level, "ts": _now_iso(), "msg": msg}
    if step is not None:
        payload["step"] = step
    if progress is not None:
        payload["progress"] = progress
    if data is not None:
        payload["data"] = data
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def emit_result(status: str, spreadsheet_path: Optional[str] = None) -> None:
    """Evento terminal indicando resultado da execucao e path da planilha gerada."""
    level = "success" if status == "ok" else ("warning" if status == "warning" else "error")
    emit(
        level,
        "Execucao concluida",
        step="done",
        data={"result": {"status": status, "spreadsheet_path": spreadsheet_path}},
    )


if __name__ == "__main__":
    emit("info", "teste: evento simples")
    emit("step", "Buscando emails", step="outlook.fetch", progress=20)
    emit("warning", "1 pendencia detectada", data={"cnpj_count": 1})
    emit_result("ok", r"C:\Users\comercial05\Documents\Gerencie_Carteira\planilhas\Gerencie Carteira_2026_04_17.xlsm")
