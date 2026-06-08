"""Regressao v4.2.7: reinjecao no PROCX replica as formulas da linha anterior.

`ListRows.Add` so autopreenche colunas REGISTRADAS como "coluna calculada" do
ListObject. Na Tabela PROCX, A (XLOOKUP de matriz), D (CNPJ Numeros) e E (Raiz
CNPJ) NAO sao colunas calculadas e ficavam vazias nas linhas reinjetadas. O fix
copia a formula da linha anterior para toda coluna de formula, exceto as de
input (gerente/CNPJ).
"""
import configparser

from gerencie_carteira import reinjetar_procx


class _FakeCell:
    """Celula com formula/valor; `.copy(dest)` transfere a formula (como o
    Range.Copy do Excel, simplificado para o teste)."""

    def __init__(self, store, row, col):
        self._store = store
        self._key = (row, col)

    @property
    def formula(self):
        return self._store.get(self._key, {}).get("formula", "")

    @property
    def value(self):
        return self._store.get(self._key, {}).get("value")

    @value.setter
    def value(self, v):
        self._store[self._key] = {"value": v, "formula": ""}

    def copy(self, dest):
        dest._store[dest._key] = {"value": None, "formula": self.formula}


class _FakeListColumns:
    def __init__(self, count):
        self.Count = count


class _FakeRange:
    def __init__(self, row):
        self.Row = row


class _FakeListRows:
    def __init__(self, lo):
        self._lo = lo

    def Add(self):
        self._lo._last_row += 1
        return type("Nova", (), {"Range": _FakeRange(self._lo._last_row)})()


class _FakeListObject:
    def __init__(self, n_cols, last_row):
        self.ListColumns = _FakeListColumns(n_cols)
        self._last_row = last_row
        self.ListRows = _FakeListRows(self)
        self.Name = "Tabela2_1"


class _FakeListObjects:
    def __init__(self, lo):
        self._lo = lo
        self.Count = 1

    def Item(self, i):
        return self._lo


class _FakeApi:
    def __init__(self, lo):
        self.ListObjects = _FakeListObjects(lo)


class _FakeSheet:
    def __init__(self, store, lo):
        self._store = store
        self.api = _FakeApi(lo)

    def range(self, rc):
        row, col = rc
        return _FakeCell(self._store, row, col)


class _FakeWB:
    def __init__(self, sheet):
        self._sheet = sheet

    @property
    def sheets(self):
        return {"PROCX GERENTES": self._sheet}


def _config():
    cfg = configparser.ConfigParser()
    cfg["Excel"] = {
        "sheet_procx": "PROCX GERENTES",
        "col_procx_gerente": "B",
        "col_procx_cnpj": "C",
        "tabela_procx": "",
    }
    return cfg


def test_reinjecao_replica_formulas_das_colunas_nao_input(capsys):
    # Tabela A..G (7 cols), ultima linha de dados = 10, com formulas em
    # A, D, E, F, G e inputs em B, C.
    store = {
        (10, 1): {"formula": "=XLOOKUP(C10)", "value": None},
        (10, 2): {"formula": "", "value": "GERENTE ANTIGO"},
        (10, 3): {"formula": "", "value": "000.000/0001-00"},
        (10, 4): {"formula": "=LIMPA(C10)", "value": None},
        (10, 5): {"formula": "=LEFT(D10,8)", "value": None},
        (10, 6): {"formula": "=MINIFS(...C10...)", "value": None},
        (10, 7): {"formula": "=MAXIFS(...C10...)", "value": None},
    }
    lo = _FakeListObject(n_cols=7, last_row=10)
    wb = _FakeWB(_FakeSheet(store, lo))

    n = reinjetar_procx(wb, _config(), {"12345678000199": "NOVO GERENTE"})

    assert n == 1
    nova = 11
    # Colunas de formula (A, D, E, F, G) foram replicadas da linha anterior.
    for col in (1, 4, 5, 6, 7):
        assert store[(nova, col)]["formula"] == store[(10, col)]["formula"]
        assert store[(nova, col)]["formula"].startswith("=")
    # Colunas de input recebem o valor, nao a formula.
    assert store[(nova, 2)]["value"] == "NOVO GERENTE"
    assert store[(nova, 3)]["value"] == "12345678000199"


def test_reinjecao_nao_copia_colunas_input_da_linha_anterior(capsys):
    # Garante que B/C nao herdam o valor antigo da linha anterior.
    store = {
        (5, 1): {"formula": "=F(C5)", "value": None},
        (5, 2): {"formula": "", "value": "GERENTE ANTIGO"},
        (5, 3): {"formula": "", "value": "CNPJ ANTIGO"},
    }
    lo = _FakeListObject(n_cols=3, last_row=5)
    wb = _FakeWB(_FakeSheet(store, lo))

    reinjetar_procx(wb, _config(), {"99": "GERENTE NOVO"})

    assert store[(6, 2)]["value"] == "GERENTE NOVO"
    assert store[(6, 3)]["value"] == "99"
