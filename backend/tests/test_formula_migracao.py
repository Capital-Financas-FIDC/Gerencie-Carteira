"""Migracao da coluna de verificacao: XLOOKUP -> INDEX/MATCH (v4.2.15).

Bug #NOME?: o template grava a coluna de gerente como `_xlfn.XLOOKUP` (future-
function). Nas maquinas do cadastro esse token nao re-linka no
`CalculateFullRebuild` em processo (xlwings) -> a coluna inteira vira `#NOME?` e o
`RefreshTable` congela o erro no cache da pivot (a coluna se cura sozinha so quando
alguem abre o arquivo; a pivot, com "atualizar ao abrir" desmarcado, nao). Prova:
o arquivo salvo pela maquina do cadastro tinha as 8815 celulas da coluna E em
cache como `#NAME?`. INDEX/MATCH e nativo (sem `_xlfn`), resolve em qualquer Excel
e tambem no rebuild em processo. VLOOKUP nao serve: o Gerente (col B do PROCX) fica
A ESQUERDA do CNPJ (col C).
"""
import pytest

from gerencie_carteira import (
    xlookup_para_index_match,
    _split_top_level,
    classificar_erros,
)


# --- _split_top_level ---

def test_split_respeita_colchetes_de_ref_estruturada():
    # As virgulas internas de Tabela2_1[...] nao devem dividir os argumentos.
    args = _split_top_level("$A2,Tabela2_1[CNPJ],Tabela2_1[Gerente]")
    assert args == ["$A2", "Tabela2_1[CNPJ]", "Tabela2_1[Gerente]"]


def test_split_respeita_parenteses_e_aspas():
    args = _split_top_level('SE(A1>0,"sim, ok","nao"),B1')
    assert args == ['SE(A1>0,"sim, ok","nao")', "B1"]


# --- xlookup_para_index_match ---

def test_converte_xlookup_estruturado_para_index_match():
    f = "=_xlfn.XLOOKUP($A8815,Tabela2_1[CNPJ],Tabela2_1[Gerente])"
    out = xlookup_para_index_match(f, linha_alvo=2)
    assert out == "=INDEX(Tabela2_1[Gerente],MATCH($A2,Tabela2_1[CNPJ],0))"


def test_converte_sem_prefixo_xlfn():
    f = "=XLOOKUP($A50,Tabela2_1[CNPJ],Tabela2_1[Gerente])"
    out = xlookup_para_index_match(f, linha_alvo=7)
    assert out == "=INDEX(Tabela2_1[Gerente],MATCH($A7,Tabela2_1[CNPJ],0))"


def test_linha_alvo_reescreve_a_linha_da_busca():
    f = "=XLOOKUP($A999,Tab[CNPJ],Tab[Ger])"
    assert "MATCH($A3," in xlookup_para_index_match(f, linha_alvo=3)


def test_index_match_existente_retorna_none():
    # Steady state: ja migrada -> None (caller mantem o comportamento de copia).
    f = "=INDEX(Tabela2_1[Gerente],MATCH($A2,Tabela2_1[CNPJ],0))"
    assert xlookup_para_index_match(f) is None


def test_formula_custom_ou_vazia_retorna_none():
    assert xlookup_para_index_match("=VLOOKUP(A2,B:C,2,0)") is None
    assert xlookup_para_index_match("") is None
    assert xlookup_para_index_match(None) is None
    assert xlookup_para_index_match("texto solto") is None


def test_xlookup_com_args_extras_usa_os_tres_primeiros():
    # XLOOKUP(busca, array_busca, array_ret, [se_nao_achar], ...) -> ignora extras.
    f = '=XLOOKUP($A2,Tab[CNPJ],Tab[Ger],"sem cadastro")'
    out = xlookup_para_index_match(f)
    assert out == "=INDEX(Tab[Ger],MATCH($A2,Tab[CNPJ],0))"


# --- classificar_erros ---

def test_classifica_nome_e_nd_separadamente():
    vals = ["Joao", "#NOME?", "#N/D", "#NAME?", "#N/A", "Maria", 123, None]
    cont = classificar_erros(vals)
    assert cont == {"NOME": 2, "NA": 2}


def test_ignora_valores_nao_erro():
    assert classificar_erros(["a", "b", 1, 2.5, None]) == {}


def test_erro_desconhecido_cai_em_outro():
    assert classificar_erros(["#ALGO!"]) == {"OUTRO": 1}
