"""
Testes da instrumentacao de timeline em memoria do log_emitter (v4.2.8).

Verifica:
  - reset_timer + sequencia de emit(step=...) -> get_timeline() ordenada
  - t_ms e dt_ms sao monotonicos (>=0)
  - emit sem step nao polui a timeline
  - contrato JSON Lines intacto (1 linha valida por emit, mesmo com timeline ativa)
"""
import json
import time

import pytest

from log_emitter import emit, get_timeline, reset_timer


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _capturar(capsys) -> list[dict]:
    out = capsys.readouterr().out.splitlines()
    return [json.loads(linha) for linha in out if linha.strip()]


# ---------------------------------------------------------------------------
# Testes da timeline
# ---------------------------------------------------------------------------

def test_reset_limpa_timeline():
    reset_timer()
    emit("info", "msg sem step")
    emit("step", "etapa 1", step="fase.um")
    reset_timer()
    assert get_timeline() == []


def test_emit_sem_step_nao_polui_timeline(capsys):
    reset_timer()
    emit("info", "sem step")
    emit("warning", "tambem sem step")
    assert get_timeline() == []
    # Contrato JSON: 2 linhas validas no stdout
    events = _capturar(capsys)
    assert len(events) == 2


def test_timeline_registra_apenas_emit_com_step(capsys):
    reset_timer()
    emit("info", "sem step A")
    emit("step", "etapa A", step="fase.a")
    emit("info", "sem step B")
    emit("step", "etapa B", step="fase.b")
    tl = get_timeline()
    assert len(tl) == 2
    assert tl[0]["step"] == "fase.a"
    assert tl[1]["step"] == "fase.b"


def test_timeline_campos_presentes():
    reset_timer()
    emit("step", "etapa X", step="fase.x")
    tl = get_timeline()
    assert len(tl) == 1
    entrada = tl[0]
    assert "step" in entrada
    assert "msg" in entrada
    assert "t_ms" in entrada
    assert "dt_ms" in entrada


def test_t_ms_monotonicos():
    reset_timer()
    emit("step", "a", step="passo.1")
    time.sleep(0.005)
    emit("step", "b", step="passo.2")
    time.sleep(0.005)
    emit("step", "c", step="passo.3")
    tl = get_timeline()
    assert len(tl) == 3
    t_valores = [e["t_ms"] for e in tl]
    assert t_valores == sorted(t_valores), "t_ms deve ser monotonicamente nao-decrescente"


def test_dt_ms_nao_negativo():
    reset_timer()
    emit("step", "primeiro", step="p.1")
    emit("step", "segundo", step="p.2")
    tl = get_timeline()
    for entrada in tl:
        assert entrada["dt_ms"] >= 0, f"dt_ms negativo em {entrada}"


def test_t_ms_primeiro_step_comeca_proximo_de_zero():
    reset_timer()
    emit("step", "imediato", step="imediato")
    tl = get_timeline()
    # Deve ter levado menos de 500 ms entre reset_timer e o emit (margem generosa)
    assert tl[0]["t_ms"] < 500


def test_get_timeline_retorna_copia():
    """Modificar o retorno de get_timeline nao altera o estado interno."""
    reset_timer()
    emit("step", "original", step="orig")
    copia = get_timeline()
    copia.append({"step": "falso", "msg": "injetado", "t_ms": 0, "dt_ms": 0})
    tl2 = get_timeline()
    assert len(tl2) == 1, "get_timeline deve retornar copia independente"


def test_contrato_json_lines_intacto_com_timeline_ativa(capsys):
    """A ativacao da timeline nao altera o JSON emitido em stdout."""
    reset_timer()
    emit("step", "com step", step="fase.z", progress=50, data={"x": 1})
    events = _capturar(capsys)
    assert len(events) == 1
    ev = events[0]
    assert ev["step"] == "fase.z"
    assert ev["progress"] == 50
    assert ev["data"] == {"x": 1}
    # Campos de timing NAO devem aparecer no JSON do stdout
    assert "t_ms" not in ev
    assert "dt_ms" not in ev


def test_multiplos_resets_acumulam_apenas_apos_ultimo(capsys):
    reset_timer()
    emit("step", "antes do reset", step="antes")
    reset_timer()
    emit("step", "depois do reset", step="depois")
    tl = get_timeline()
    assert len(tl) == 1
    assert tl[0]["step"] == "depois"
