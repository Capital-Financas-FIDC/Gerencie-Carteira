"""Testes da escrita transacional do .xlsm (escrever_parcial + promover)."""
import os

import pytest

from xlsm_transacional import (
    _eh_orfao,
    escrever_parcial,
    limpar_publico_antigos,
    promover,
    sweep_orfaos,
)


class _FakeWB:
    """wb cujo save(path) escreve um arquivo com conteudo fixo."""

    def __init__(self, conteudo="novo"):
        self.conteudo = conteudo

    def save(self, path):
        with open(path, "w", encoding="utf-8") as f:
            f.write(self.conteudo)


class _FailWB:
    def save(self, path):
        raise IOError("falha simulada no save")


def _listar(p):
    return sorted(os.listdir(p))


def _sem_orfaos(nomes):
    return not any(_eh_orfao(n) for n in nomes)


def test_fluxo_feliz_cria_destino_sem_residuo(tmp_path):
    destino = str(tmp_path / "Gerencie Carteira_2026_05_18.xlsm")
    parcial = escrever_parcial(_FakeWB("novo"), destino)
    assert os.path.isfile(parcial) and not os.path.exists(destino)  # ainda nao promovido
    assert promover(parcial, destino) == destino
    assert os.path.isfile(destino)
    with open(destino, encoding="utf-8") as f:
        assert f.read() == "novo"
    assert _sem_orfaos(_listar(tmp_path))


def test_colisao_substitui_e_limpa_bak(tmp_path):
    destino = str(tmp_path / "base.xlsm")
    with open(destino, "w", encoding="utf-8") as f:
        f.write("antigo")
    parcial = escrever_parcial(_FakeWB("novo"), destino)
    promover(parcial, destino)
    with open(destino, encoding="utf-8") as f:
        assert f.read() == "novo"
    assert _sem_orfaos(_listar(tmp_path))


def test_falha_no_save_mantem_base_intacta(tmp_path):
    destino = str(tmp_path / "base.xlsm")
    with open(destino, "w", encoding="utf-8") as f:
        f.write("intacto")
    with pytest.raises(IOError):
        escrever_parcial(_FailWB(), destino)
    with open(destino, encoding="utf-8") as f:
        assert f.read() == "intacto"
    assert _sem_orfaos(_listar(tmp_path))


def test_promover_sem_parcial_levanta(tmp_path):
    with pytest.raises(FileNotFoundError):
        promover(str(tmp_path / "nao_existe.partial.xlsm"), str(tmp_path / "x.xlsm"))


def test_sweep_orfaos_idempotente(tmp_path):
    (tmp_path / "Gerencie Carteira_2026_05_16.partial.xlsm").write_text("p")
    (tmp_path / "Gerencie Carteira_2026_05_16.bak.xlsm").write_text("b")
    (tmp_path / "Gerencie Carteira_2026_05_16.xlsm").write_text("k")
    rem1 = sweep_orfaos(str(tmp_path))
    assert len(rem1) == 2
    assert _listar(tmp_path) == ["Gerencie Carteira_2026_05_16.xlsm"]
    assert sweep_orfaos(str(tmp_path)) == []  # idempotente


def test_publico_remove_antigos_apos_parcial(tmp_path):
    (tmp_path / "Gerencie Carteira_2026_05_17.xlsm").write_text("antigo")
    destino = str(tmp_path / "Gerencie Carteira_2026_05_18.xlsm")
    parcial = escrever_parcial(_FakeWB("pub"), destino)
    # parcial ja existe -> agora seguro remover os antigos e promover
    falhas = limpar_publico_antigos(str(tmp_path))
    assert falhas == []
    assert os.path.isfile(parcial)  # nao removido por limpar_publico_antigos
    promover(parcial, destino)
    nomes = _listar(tmp_path)
    assert "Gerencie Carteira_2026_05_17.xlsm" not in nomes
    assert "Gerencie Carteira_2026_05_18.xlsm" in nomes
    assert _sem_orfaos(nomes)


def test_limpar_publico_falha_retorna_caminho(tmp_path, monkeypatch):
    """Se os.remove falhar (planilha aberta em outro Excel), retorna a lista
    de falhas para o caller emitir warning — nao engole silenciosamente."""
    alvo = tmp_path / "Gerencie Carteira_2026_05_17.xlsm"
    alvo.write_text("antigo")

    import xlsm_transacional as xt
    chamadas = []

    def fake_remove(path):
        chamadas.append(path)
        raise PermissionError(32, "arquivo em uso")

    monkeypatch.setattr(xt.os, "remove", fake_remove)

    falhas = xt.limpar_publico_antigos(
        str(tmp_path), tentativas=3, intervalo=0
    )
    assert len(falhas) == 1
    caminho_falho, err = falhas[0]
    assert os.path.basename(caminho_falho) == "Gerencie Carteira_2026_05_17.xlsm"
    assert isinstance(err, PermissionError)
    assert len(chamadas) == 3  # tentativas esgotadas
