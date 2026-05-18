"""Testes da cascata de descoberta da planilha base."""
import configparser
import os
from unittest.mock import patch

import pytest


def _make_config(local: str, publica: str) -> configparser.ConfigParser:
    c = configparser.ConfigParser(interpolation=None)
    c["Paths"] = {
        "pasta_destino_html": "ignored",
        "pasta_diario_excel": local,
        "pasta_copia_excel": publica,
        "pasta_logs": "ignored",
    }
    c["Excel"] = {"planilha_dados": "E-Mail BD", "coluna_verificacao": "E"}
    c["Email"] = {"assunto_procurado": "ignored"}
    return c


def _touch(path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write("")


@pytest.fixture
def pipeline_module():
    import importlib
    return importlib.import_module("gerencie_carteira")


# Entrypoint tem nome estavel (sem versao): importavel diretamente.
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import importlib.util
_src = Path(__file__).resolve().parent.parent / "src" / "gerencie_carteira.py"
_spec = importlib.util.spec_from_file_location("pipeline_core", _src)
pipeline = importlib.util.module_from_spec(_spec)  # type: ignore
_spec.loader.exec_module(pipeline)  # type: ignore


def test_buscar_mais_recente_pasta_inexistente(tmp_path):
    assert pipeline._buscar_mais_recente_em(str(tmp_path / "nope")) is None


def test_buscar_mais_recente_pasta_vazia(tmp_path):
    assert pipeline._buscar_mais_recente_em(str(tmp_path)) is None


def test_buscar_mais_recente_ignora_nomes_invalidos(tmp_path):
    _touch(str(tmp_path / "template.xlsm"))
    _touch(str(tmp_path / "outro.xlsx"))
    assert pipeline._buscar_mais_recente_em(str(tmp_path)) is None


def test_buscar_mais_recente_retorna_mais_recente(tmp_path):
    _touch(str(tmp_path / "Gerencie Carteira_2026_04_10.xlsm"))
    _touch(str(tmp_path / "Gerencie Carteira_2026_04_15.xlsm"))
    _touch(str(tmp_path / "Gerencie Carteira_2026_04_03.xlsm"))
    result = pipeline._buscar_mais_recente_em(str(tmp_path))
    assert result is not None
    assert result.endswith("Gerencie Carteira_2026_04_15.xlsm")


def test_copiar_base_para_local_cria_destino(tmp_path):
    origem = tmp_path / "src" / "Gerencie Carteira_2026_04_10.xlsm"
    origem.parent.mkdir()
    origem.write_text("conteudo")
    destino_dir = tmp_path / "dest"
    destino = pipeline._copiar_base_para_local(str(origem), str(destino_dir))
    assert os.path.isfile(destino)
    assert destino.endswith("Gerencie Carteira_2026_04_10.xlsm")


def test_copiar_base_para_local_mesmo_arquivo_noop(tmp_path):
    target = tmp_path / "Gerencie Carteira_2026_04_10.xlsm"
    target.write_text("x")
    result = pipeline._copiar_base_para_local(str(target), str(tmp_path))
    assert os.path.abspath(result) == os.path.abspath(str(target))


def test_cascata_prioriza_local(tmp_path, capsys):
    local = tmp_path / "local"; local.mkdir()
    legada = tmp_path / "legada"; legada.mkdir()
    publica = tmp_path / "publica"; publica.mkdir()

    _touch(str(local / "Gerencie Carteira_2026_04_10.xlsm"))
    _touch(str(legada / "Gerencie Carteira_2026_04_15.xlsm"))  # mais recente, mas deve perder
    _touch(str(publica / "Gerencie Carteira_2026_04_20.xlsm"))

    config = _make_config(str(local), str(publica))
    with patch.object(pipeline, "resolve_legacy_planilhas", return_value=legada):
        result = pipeline.encontrar_arquivo_base_excel(config)
    assert result.endswith("2026_04_10.xlsm")  # pegou o local, ignorando os mais recentes
    events = [line for line in capsys.readouterr().out.splitlines() if line.strip()]
    assert any("excel.base.local" in e for e in events)


def test_cascata_fallback_legada_copia_para_local(tmp_path, capsys):
    local = tmp_path / "local"; local.mkdir()
    legada = tmp_path / "legada"; legada.mkdir()
    publica = tmp_path / "publica"  # nao existe

    _touch(str(legada / "Gerencie Carteira_2026_04_15.xlsm"))

    config = _make_config(str(local), str(publica))
    with patch.object(pipeline, "resolve_legacy_planilhas", return_value=legada):
        result = pipeline.encontrar_arquivo_base_excel(config)

    # Resultado aponta para a COPIA no local
    assert str(local) in result
    assert result.endswith("Gerencie Carteira_2026_04_15.xlsm")
    assert os.path.isfile(result)
    assert os.path.isfile(str(legada / "Gerencie Carteira_2026_04_15.xlsm"))  # original preservado

    events = [line for line in capsys.readouterr().out.splitlines() if line.strip()]
    assert any("excel.base.legacy" in e for e in events)


def test_cascata_fallback_publica(tmp_path, capsys):
    local = tmp_path / "local"; local.mkdir()
    publica = tmp_path / "publica"; publica.mkdir()
    _touch(str(publica / "Gerencie Carteira_2026_04_20.xlsm"))

    config = _make_config(str(local), str(publica))
    with patch.object(pipeline, "resolve_legacy_planilhas", return_value=None):
        result = pipeline.encontrar_arquivo_base_excel(config)

    assert str(local) in result
    assert result.endswith("Gerencie Carteira_2026_04_20.xlsm")
    events = [line for line in capsys.readouterr().out.splitlines() if line.strip()]
    assert any("excel.base.public" in e for e in events)


def test_cascata_esgotada_emite_needs_user_e_sai(tmp_path, capsys):
    local = tmp_path / "local"; local.mkdir()
    publica = tmp_path / "publica"  # nao existe

    config = _make_config(str(local), str(publica))
    with patch.object(pipeline, "resolve_legacy_planilhas", return_value=None):
        with pytest.raises(SystemExit) as exc_info:
            pipeline.encontrar_arquivo_base_excel(config)

    assert exc_info.value.code == pipeline.EXIT_BASE_NEEDS_USER
    events = [line for line in capsys.readouterr().out.splitlines() if line.strip()]
    assert any("excel.base.needs_user" in e for e in events)
