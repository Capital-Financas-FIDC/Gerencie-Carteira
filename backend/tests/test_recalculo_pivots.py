"""Recalculo do workbook: full rebuild vs Calculate normal.

v4.2.7: o full rebuild antes do RefreshTable evita que a pivot congele #NOME?
apos `ListRows.Add` (mudanca estrutural no PROCX).
v4.2.9: tornou o full rebuild CONDICIONAL (so com orfaos) — REGRESSAO: a coluna
de verificacao (Gerente) usa XLOOKUP (`_xlfn.XLOOKUP`) e e copiada para linhas
novas a CADA run; um `app.calculate()` normal nao religa essa future-function nas
celulas recem-copiadas -> ficam `#NOME?` e o RefreshTable congela o erro na pivot.
v4.2.10: o call-site (`atualizar_planilha_excel`) voltou a chamar SEMPRE
`recalcular(app, full=True)`. A funcao `recalcular` continua suportando ambos os
modos (testados abaixo); a App roda em modo manual e os metodos Calculate* forcam
o calculo mesmo assim.
v4.2.11 (CINTO): `recalcular` passa a ESPERAR `Application.CalculationState ==
xlDone` antes de retornar — o `CalculateFullRebuild` pode devolver o controle ao
COM antes de a propagacao terminar (Excel mais lento que o dev), e o
`RefreshTable` seguinte congelava o `#NOME?` transitorio no cache da pivot. A
espera e limitada por `timeout` (warning + segue ao estourar).
"""
import pytest

from gerencie_carteira import recalcular


class _FakeApi:
    def __init__(self):
        self.calls = []

    def CalculateFullRebuild(self):
        self.calls.append("full_rebuild")

    def CalculateUntilAsyncQueriesDone(self):
        self.calls.append("async_done")


class _FakeApp:
    def __init__(self, api):
        self.api = api
        self.calculate_calls = 0

    def calculate(self):
        self.calculate_calls += 1


def test_full_true_reconstroi_arvore_e_espera_async():
    api = _FakeApi()
    app = _FakeApp(api)

    recalcular(app, full=True)

    # Full rebuild + espera do async; sem Calculate simples (rebuild ja calcula).
    assert api.calls == ["full_rebuild", "async_done"]
    assert app.calculate_calls == 0


def test_full_false_faz_calculate_normal_nao_rebuild():
    api = _FakeApi()
    app = _FakeApp(api)

    recalcular(app, full=False)

    # Caso comum: Calculate normal (NAO full rebuild) + espera do async.
    assert "full_rebuild" not in api.calls
    assert api.calls == ["async_done"]
    assert app.calculate_calls == 1


def test_full_true_cai_no_fallback_quando_rebuild_indisponivel():
    class _ApiSemRebuild:
        def __init__(self):
            self.calls = []

        def CalculateFullRebuild(self):
            raise AttributeError("metodo indisponivel")

        def CalculateUntilAsyncQueriesDone(self):
            self.calls.append("async_done")

    app = _FakeApp(_ApiSemRebuild())
    recalcular(app, full=True)
    # Sem full rebuild, ao menos um Calculate simples deve ter rodado.
    assert app.calculate_calls == 1


def test_espera_calculo_assentar_antes_de_retornar():
    """Com a App reportando xlPending->xlPending->xlDone, recalcular deve
    POLLAR o CalculationState ate xlDone (e nao retornar no primeiro pending)."""
    XL_PENDING, XL_DONE = 2, 0

    class _ApiComEstado:
        def __init__(self):
            self.calls = []
            self.estados = [XL_PENDING, XL_PENDING, XL_DONE]
            self.leituras = 0

        def CalculateFullRebuild(self):
            self.calls.append("full_rebuild")

        def CalculateUntilAsyncQueriesDone(self):
            self.calls.append("async_done")

        @property
        def CalculationState(self):
            estado = self.estados[min(self.leituras, len(self.estados) - 1)]
            self.leituras += 1
            return estado

    api = _ApiComEstado()
    app = _FakeApp(api)

    recalcular(app, full=True, timeout=5.0)

    # Pollou ate o xlDone (3 leituras: pending, pending, done).
    assert api.leituras == 3
    assert api.calls.count("full_rebuild") == 1


def test_timeout_quando_calculo_nunca_assenta_nao_trava():
    """CalculationState eternamente xlPending: recalcular deve respeitar o
    timeout, NAO travar, e nao propagar excecao."""
    XL_PENDING = 2

    class _ApiTravada:
        def CalculateFullRebuild(self):
            pass

        def CalculateUntilAsyncQueriesDone(self):
            pass

        @property
        def CalculationState(self):
            return XL_PENDING

    app = _FakeApp(_ApiTravada())
    # timeout=0 -> sai na primeira verificacao de deadline, sem sleep.
    recalcular(app, full=True, timeout=0.0)


def test_recalcular_nunca_propaga_excecao():
    class _ApiQuebrada:
        def CalculateFullRebuild(self):
            raise RuntimeError("COM morto")

        def CalculateUntilAsyncQueriesDone(self):
            raise RuntimeError("COM morto")

    class _AppQuebrado:
        api = _ApiQuebrada()

        def calculate(self):
            raise RuntimeError("COM morto")

    # Falhas de COM nao podem derrubar o pipeline (full=True e full=False).
    try:
        recalcular(_AppQuebrado(), full=True)
        recalcular(_AppQuebrado(), full=False)
    except Exception:  # pragma: no cover
        pytest.fail("recalcular nao deve propagar excecao")
