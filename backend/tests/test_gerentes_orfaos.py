"""Testes da deteccao de gerentes orfaos e helpers do PROCX."""
from datetime import datetime

import pandas as pd

from gerencie_carteira import (
    _col_idx,
    _normalizar_cnpj,
    detectar_orfaos,
    marcar_emails_lidos,
)


def _df(cnpjs):
    return pd.DataFrame(
        {
            "CNPJ": cnpjs,
            "Razão Social": [f"Empresa {i}" for i in range(len(cnpjs))],
            "Alteração": ["x"] * len(cnpjs),
            "Data do recebimento do e-mail": [datetime(2026, 5, 18)] * len(cnpjs),
        }
    )


def test_normalizar_cnpj_remove_pontuacao():
    assert _normalizar_cnpj("12.345.678/0001-90") == "12345678000190"
    assert _normalizar_cnpj(" 00 11 ") == "0011"
    assert _normalizar_cnpj(None) == ""


def test_col_idx():
    assert _col_idx("A") == 1
    assert _col_idx("B") == 2
    assert _col_idx("C") == 3
    assert _col_idx("Z") == 26
    assert _col_idx("AA") == 27


def test_detectar_orfaos_basico(capsys):
    df = _df(["11.111.111/0001-11", "22222222000122"])
    mapa = {"11111111000111": "Ana"}
    orfaos = detectar_orfaos(df, mapa)
    assert [o["cnpj"] for o in orfaos] == ["22222222000122"]
    assert orfaos[0]["razao_social"] == "Empresa 1"


def test_detectar_orfaos_dedupe(capsys):
    df = _df(["33333333000133", "33.333.333/0001-33", "33333333000133"])
    orfaos = detectar_orfaos(df, {})
    assert len(orfaos) == 1


def test_detectar_orfaos_nenhum(capsys):
    df = _df(["44444444000144"])
    assert detectar_orfaos(df, {"44444444000144": "Bia"}) == []


def test_detectar_orfaos_mapa_vazio_todos_orfaos(capsys):
    df = _df(["55555555000155", "66666666000166"])
    assert len(detectar_orfaos(df, {})) == 2


class _FakeRT:
    def __init__(self, d):
        self._d = d

    def date(self):
        return self._d.date()


class _FakeEmail:
    def __init__(self, dt):
        self.ReceivedTime = _FakeRT(dt)
        self.UnRead = True


def test_marcar_emails_lidos(capsys):
    e1, e2 = _FakeEmail(datetime(2026, 5, 18)), _FakeEmail(datetime(2026, 5, 17))
    marcar_emails_lidos([e1, e2])
    assert e1.UnRead is False and e2.UnRead is False
