"""Testes da retencao de backups (planilhas .xlsm e anexos .html)."""
import importlib.util
import os
import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

# Entrypoint tem nome estavel (sem versao): carregavel via spec.
_src = Path(__file__).resolve().parent.parent / "src" / "gerencie_carteira.py"
_spec = importlib.util.spec_from_file_location("pipeline_retencao", _src)
pipeline = importlib.util.module_from_spec(_spec)  # type: ignore
_spec.loader.exec_module(pipeline)  # type: ignore


def _touch(caminho) -> None:
    with open(caminho, "w", encoding="utf-8") as f:
        f.write("")


def _datas(n: int) -> list[str]:
    """n datas distintas e crescentes no formato YYYY_MM_DD."""
    base = date(2026, 1, 1)
    return [(base + timedelta(days=i)).strftime("%Y_%m_%d") for i in range(n)]


def _criar(pasta, datas, prefixo: str, ext: str) -> None:
    for d in datas:
        _touch(os.path.join(pasta, f"{prefixo}{d}{ext}"))


# --- planilhas (.xlsm) ---

def test_abaixo_do_limite_nao_remove(tmp_path):
    _criar(tmp_path, _datas(10), "Gerencie Carteira_", ".xlsm")
    removidos = pipeline.limpar_backups_antigos(
        str(tmp_path), pipeline.BASE_FILENAME_PATTERN, 30, "planilhas")
    assert removidos == 0
    assert len(os.listdir(tmp_path)) == 10


def test_no_limite_exato_nao_remove(tmp_path):
    _criar(tmp_path, _datas(30), "Gerencie Carteira_", ".xlsm")
    removidos = pipeline.limpar_backups_antigos(
        str(tmp_path), pipeline.BASE_FILENAME_PATTERN, 30, "planilhas")
    assert removidos == 0
    assert len(os.listdir(tmp_path)) == 30


def test_acima_do_limite_remove_os_mais_antigos(tmp_path):
    datas = _datas(33)
    _criar(tmp_path, datas, "Gerencie Carteira_", ".xlsm")
    removidos = pipeline.limpar_backups_antigos(
        str(tmp_path), pipeline.BASE_FILENAME_PATTERN, 30, "planilhas")
    assert removidos == 3
    restantes = set(os.listdir(tmp_path))
    assert len(restantes) == 30
    for d in datas[:3]:  # os 3 mais antigos sumiram
        assert f"Gerencie Carteira_{d}.xlsm" not in restantes
    for d in datas[3:]:  # os 30 mais recentes ficaram
        assert f"Gerencie Carteira_{d}.xlsm" in restantes


def test_copia_31_dispara_remocao_da_mais_antiga(tmp_path):
    """Cenario do pedido: a copia n 31 remove a planilha n 1."""
    datas = _datas(31)
    _criar(tmp_path, datas, "Gerencie Carteira_", ".xlsm")
    removidos = pipeline.limpar_backups_antigos(
        str(tmp_path), pipeline.BASE_FILENAME_PATTERN, 30, "planilhas")
    assert removidos == 1
    assert not os.path.exists(
        os.path.join(tmp_path, f"Gerencie Carteira_{datas[0]}.xlsm"))
    assert os.path.exists(
        os.path.join(tmp_path, f"Gerencie Carteira_{datas[-1]}.xlsm"))


def test_ignora_arquivos_que_nao_casam_o_padrao(tmp_path):
    _criar(tmp_path, _datas(31), "Gerencie Carteira_", ".xlsm")
    # Ruidos que NAO devem ser contados nem removidos:
    _touch(os.path.join(tmp_path, "template.xlsm"))
    _touch(os.path.join(tmp_path, "Gerencie Carteira_2026_01_01.partial.xlsm"))
    _touch(os.path.join(tmp_path, "Gerencie Carteira_2026_01_01.bak.xlsm"))
    removidos = pipeline.limpar_backups_antigos(
        str(tmp_path), pipeline.BASE_FILENAME_PATTERN, 30, "planilhas")
    assert removidos == 1  # apenas 1 planilha valida excedente (31 - 30)
    assert os.path.exists(os.path.join(tmp_path, "template.xlsm"))
    assert os.path.exists(
        os.path.join(tmp_path, "Gerencie Carteira_2026_01_01.partial.xlsm"))
    assert os.path.exists(
        os.path.join(tmp_path, "Gerencie Carteira_2026_01_01.bak.xlsm"))


# --- html ---

def test_html_acima_do_limite_remove_os_mais_antigos(tmp_path):
    datas = _datas(35)
    _criar(tmp_path, datas, "Gerencie_Carteira_", ".html")
    removidos = pipeline.limpar_backups_antigos(
        str(tmp_path), pipeline.HTML_FILENAME_PATTERN, 30, "html")
    assert removidos == 5
    assert len(os.listdir(tmp_path)) == 30


# --- guardas ---

def test_pasta_inexistente_retorna_zero(tmp_path):
    removidos = pipeline.limpar_backups_antigos(
        str(tmp_path / "nao_existe"), pipeline.BASE_FILENAME_PATTERN,
        30, "planilhas")
    assert removidos == 0


def test_limite_invalido_nao_remove_nada(tmp_path):
    _criar(tmp_path, _datas(31), "Gerencie Carteira_", ".xlsm")
    removidos = pipeline.limpar_backups_antigos(
        str(tmp_path), pipeline.BASE_FILENAME_PATTERN, 0, "planilhas")
    assert removidos == 0
    assert len(os.listdir(tmp_path)) == 31
