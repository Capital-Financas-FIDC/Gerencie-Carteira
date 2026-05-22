"""Testes do redirecionamento de pastas em modo dev (isolamento dev/producao)."""
import configparser
import importlib.util
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

_src = Path(__file__).resolve().parent.parent / "src" / "gerencie_carteira.py"
_spec = importlib.util.spec_from_file_location("pipeline_config_dev", _src)
pipeline = importlib.util.module_from_spec(_spec)  # type: ignore
_spec.loader.exec_module(pipeline)  # type: ignore


def _config_rede() -> configparser.ConfigParser:
    """Config como o config.ini de producao: pastas na rede A:\\."""
    c = configparser.ConfigParser(interpolation=None)
    c["Paths"] = {
        "pasta_destino_html": r"A:\PUBLICA\X\Software\data\html",
        "pasta_diario_excel": r"A:\PUBLICA\X\Software\data\planilhas",
        "pasta_copia_excel": r"A:\PUBLICA\X",
        "pasta_logs": r"A:\PUBLICA\X\Software\data\logs",
    }
    return c


def test_modo_dev_redireciona_pastas_para_o_repo(capsys):
    # pytest roda nao-empacotado -> aplicar_pastas_dev deve redirecionar.
    c = _config_rede()
    pipeline.aplicar_pastas_dev(c)
    esperado = {
        "pasta_destino_html": "html",
        "pasta_diario_excel": "planilhas",
        "pasta_logs": "logs",
        "pasta_copia_excel": "publica",
    }
    for chave, sub in esperado.items():
        p = c["Paths"][chave]
        assert not p.upper().startswith("A:"), f"{chave} ainda aponta p/ a rede"
        assert os.path.basename(p) == sub
        assert os.path.basename(os.path.dirname(p)) == "data"


def test_modo_frozen_nao_altera_pastas(monkeypatch):
    # Simula o app empacotado: aplicar_pastas_dev deve ser no-op.
    monkeypatch.setattr(pipeline.sys, "frozen", True, raising=False)
    c = _config_rede()
    original = dict(c["Paths"])
    pipeline.aplicar_pastas_dev(c)
    assert dict(c["Paths"]) == original
