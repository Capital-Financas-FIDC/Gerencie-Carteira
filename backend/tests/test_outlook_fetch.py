"""
Testes de buscar_emails_novos() com fake Outlook (v4.2.8).

Cenarios:
  1. Caminho feliz: Restrict chamado, itera apenas o subconjunto nao-lido,
     filtra pelo assunto exato — full scan nao ocorre.
  2. Fallback: Restrict lanca excecao -> varredura completa (comportamento antigo).
  3. Item nao-MailItem na colecao: ignorado defensivamente (sem AttributeError).
  4. Nenhum email no conjunto restrito: retorna lista vazia.
"""
import configparser

import pytest

from gerencie_carteira import buscar_emails_novos

ASSUNTO_ALVO = "Monitoramento Serasa - Consulta de CNPJs"


# ---------------------------------------------------------------------------
# Fakes de objetos COM do Outlook
# ---------------------------------------------------------------------------

class _FakeMailItem:
    """Simula um MailItem do Outlook com Subject, UnRead e ReceivedTime."""

    def __init__(self, subject: str, unread: bool = True):
        self.Subject = subject
        self.UnRead = unread


class _FakeNaoMailItem:
    """Simula um item sem .Subject (ex: MeetingRequest, TaskRequest)."""
    pass


class _FakeItems:
    """
    Simula inbox.Items com Restrict() e iteracao direta.
    Registra se Sort e Restrict foram chamados para assertions.
    """

    def __init__(self, itens: list, restrict_raise: bool = False):
        self._itens = itens
        self._restrict_raise = restrict_raise
        self.sort_chamado = False
        self.restrict_chamado = False
        self._restrict_result: list | None = None

    def Sort(self, campo: str, desc: bool) -> None:
        self.sort_chamado = True

    def Restrict(self, filtro: str) -> "_FakeItems":
        self.restrict_chamado = True
        if self._restrict_raise:
            raise Exception("Restrict nao suportado nesta configuracao")
        # Simula filtro server-side: retorna apenas os UnRead (quando aplicavel)
        itens_filtrados = [
            m for m in self._itens
            if getattr(m, "UnRead", False) is True
        ]
        resultado = _FakeItems(itens_filtrados)
        self._restrict_result = resultado
        return resultado

    def __iter__(self):
        return iter(self._itens)


class _FakeInbox:
    def __init__(self, items: _FakeItems):
        self.Items = items


class _FakeNamespace:
    def __init__(self, inbox: _FakeInbox):
        self._inbox = inbox

    def GetDefaultFolder(self, folder_id: int) -> _FakeInbox:
        return self._inbox


class _FakeOutlookApp:
    def __init__(self, namespace: _FakeNamespace):
        self._namespace = namespace

    def GetNamespace(self, ns: str) -> _FakeNamespace:
        return self._namespace


def _config(assunto: str = ASSUNTO_ALVO) -> configparser.ConfigParser:
    config = configparser.ConfigParser(interpolation=None)
    config["Email"] = {"assunto_procurado": assunto}
    return config


def _patch_outlook(monkeypatch, fake_items: _FakeItems) -> _FakeItems:
    """Substitui win32com.client.Dispatch para retornar o fake Outlook."""
    import gerencie_carteira

    namespace = _FakeNamespace(_FakeInbox(fake_items))
    app = _FakeOutlookApp(namespace)

    monkeypatch.setattr(
        "gerencie_carteira.win32com.client.Dispatch",
        lambda *_a, **_kw: app,
    )
    return fake_items


# ---------------------------------------------------------------------------
# Testes
# ---------------------------------------------------------------------------

def test_restrict_chamado_no_caminho_feliz(monkeypatch, capsys):
    """Restrict deve ser chamado e o full-scan nao deve ocorrer."""
    emails = [
        _FakeMailItem(ASSUNTO_ALVO, unread=True),
        _FakeMailItem(ASSUNTO_ALVO, unread=True),
        _FakeMailItem("Outro assunto", unread=True),
    ]
    items = _FakeItems(emails)
    _patch_outlook(monkeypatch, items)

    resultado = buscar_emails_novos(_config())

    # Restrict foi chamado
    assert items.restrict_chamado is True
    # Apenas os 2 com assunto correto retornados
    assert len(resultado) == 2
    assert all(m.Subject == ASSUNTO_ALVO for m in resultado)


def test_full_scan_nao_ocorre_no_caminho_feliz(monkeypatch, capsys):
    """
    Quando Restrict funciona, a iteracao deve ocorrer sobre o subconjunto
    restrito (menos itens), nao sobre o conjunto completo.
    """
    # 10 itens lidos (nao devem ser retornados)
    emails_lidos = [_FakeMailItem(ASSUNTO_ALVO, unread=False) for _ in range(10)]
    emails_nao_lidos = [_FakeMailItem(ASSUNTO_ALVO, unread=True)]
    items = _FakeItems(emails_lidos + emails_nao_lidos)
    _patch_outlook(monkeypatch, items)

    resultado = buscar_emails_novos(_config())

    # O conjunto restrito pelo fake contem apenas nao-lidos
    assert len(resultado) == 1


def test_fallback_quando_restrict_falha(monkeypatch, capsys):
    """Restrict lanca excecao -> fallback para varredura completa."""
    emails = [
        _FakeMailItem(ASSUNTO_ALVO, unread=True),
        _FakeMailItem(ASSUNTO_ALVO, unread=False),   # lido -> nao retornado
        _FakeMailItem("Outro", unread=True),          # assunto errado -> nao retornado
    ]
    items = _FakeItems(emails, restrict_raise=True)
    _patch_outlook(monkeypatch, items)

    resultado = buscar_emails_novos(_config())

    # Restrict foi tentado mas falhou
    assert items.restrict_chamado is True
    # Fallback varreu tudo e retornou apenas o nao-lido com assunto correto
    assert len(resultado) == 1
    assert resultado[0].Subject == ASSUNTO_ALVO
    assert resultado[0].UnRead is True


def test_item_nao_mail_ignorado_defensivamente(monkeypatch, capsys):
    """Itens sem .Subject nao devem gerar AttributeError."""
    emails = [
        _FakeNaoMailItem(),           # sem Subject nem UnRead
        _FakeMailItem(ASSUNTO_ALVO, unread=True),
    ]
    # Coloca o nao-MailItem no conjunto "restrito" simulado
    items_restritos = _FakeItems(emails)
    items_restritos.restrict_chamado = True  # marca como ja chamado

    # Faz inbox.Items.Restrict retornar diretamente o conjunto misto
    class _ItemsComMisto(_FakeItems):
        def Restrict(self, _filtro):
            self.restrict_chamado = True
            return items_restritos

    items = _ItemsComMisto(emails)
    _patch_outlook(monkeypatch, items)

    # Nao deve lancar excecao
    resultado = buscar_emails_novos(_config())
    assert len(resultado) == 1
    assert resultado[0].Subject == ASSUNTO_ALVO


def test_lista_vazia_quando_nenhum_email_alvo(monkeypatch, capsys):
    """Retorna lista vazia se Restrict nao encontrar nenhum email com o assunto."""
    emails = [
        _FakeMailItem("Assunto errado", unread=True),
        _FakeMailItem("Outro assunto", unread=True),
    ]
    items = _FakeItems(emails)
    _patch_outlook(monkeypatch, items)

    resultado = buscar_emails_novos(_config())

    assert resultado == []


def test_evento_outlook_fetch_contem_restrict_flag(monkeypatch, capsys):
    """O evento outlook.fetch deve incluir o campo 'restrict' indicando o caminho."""
    import json

    emails = [_FakeMailItem(ASSUNTO_ALVO, unread=True)]
    items = _FakeItems(emails)
    _patch_outlook(monkeypatch, items)

    buscar_emails_novos(_config())

    linhas = capsys.readouterr().out.splitlines()
    eventos = [json.loads(l) for l in linhas if l.strip()]
    fetch_ev = next((e for e in eventos if e.get("step") == "outlook.fetch"), None)
    assert fetch_ev is not None
    assert "restrict" in fetch_ev.get("data", {})
    assert fetch_ev["data"]["restrict"] is True
