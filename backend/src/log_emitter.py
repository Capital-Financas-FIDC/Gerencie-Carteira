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
"""

import json
import sys
from datetime import datetime, timezone
from typing import Any, Optional


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
    """Emite 1 evento JSON em stdout (1 linha, flush imediato)."""
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
