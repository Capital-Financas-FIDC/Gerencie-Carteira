"""Regressao v4.2.7: recalculo COMPLETO antes de atualizar as pivots.

Bug: apos `reinjetar_procx` (ListRows.Add na Tabela PROCX), um `app.calculate()`
simples nao reconstroi a arvore de dependencias; as referencias estruturadas da
coluna de gerente em 'E-Mail BD' ficavam em #NOME? e o RefreshTable congelava
esse erro no cache da pivot. `recalcular_completo` forca CalculateFullRebuild +
espera o calculo assincrono, eliminando a janela de erro.
"""
import pytest

from gerencie_carteira import recalcular_completo


class _FakeApi:
    def __init__(self):
        self.calls = []
        self.Calculation = None

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


def test_recalculo_forca_rebuild_completo_e_async():
    api = _FakeApi()
    app = _FakeApp(api)

    recalcular_completo(app)

    # Reconstroi a arvore inteira (nao apenas Calculate) e espera o async.
    assert api.calls == ["full_rebuild", "async_done"]
    # Modo automatico (xlCalculationAutomatic).
    assert api.Calculation == -4105
    # Full rebuild bem-sucedido nao precisa do fallback Calculate simples.
    assert app.calculate_calls == 0


def test_recalculo_cai_no_fallback_quando_rebuild_indisponivel():
    class _ApiSemRebuild:
        def __init__(self):
            self.Calculation = None

        def CalculateFullRebuild(self):
            raise AttributeError("metodo indisponivel")

        def CalculateUntilAsyncQueriesDone(self):
            pass

    app = _FakeApp(_ApiSemRebuild())
    recalcular_completo(app)
    # Sem full rebuild, ao menos um Calculate simples deve ter rodado.
    assert app.calculate_calls == 1


def test_recalculo_nunca_propaga_excecao():
    class _ApiQuebrada:
        @property
        def Calculation(self):  # leitura ok
            return None

        @Calculation.setter
        def Calculation(self, _):
            raise RuntimeError("COM morto")

        def CalculateFullRebuild(self):
            raise RuntimeError("COM morto")

        def CalculateUntilAsyncQueriesDone(self):
            raise RuntimeError("COM morto")

    class _AppQuebrado:
        api = _ApiQuebrada()

        def calculate(self):
            raise RuntimeError("COM morto")

    # Falhas de COM nao podem derrubar o pipeline.
    try:
        recalcular_completo(_AppQuebrado())
    except Exception:  # pragma: no cover
        pytest.fail("recalcular_completo nao deve propagar excecao")
