# ==============================================================================
# CORE DE AUTOMACAO GERENCIE CARTEIRA
#
# Entrypoint do pipeline Python (Outlook COM -> parse HTML -> Excel .xlsm).
# Emite eventos JSON Lines em stdout, consumidos pelo shell Electron.
#
# VERSIONAMENTO: SemVer. A fonte unica da versao e o campo "version" de
# app/package.json. Este arquivo NAO carrega a versao no nome nem hardcoda
# a versao: ela chega via env APP_VERSION (injetada pelo Electron) com
# fallback para app/package.json (dev) ou "dev". Ver AGENTS.md > Versionamento.
# ==============================================================================

import configparser
import json
import logging
import os
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import win32com.client
import xlwings as xw
from bs4 import BeautifulSoup

from log_emitter import emit, emit_result, reset_timer, get_timeline
from directory_bootstrap import (
    ensure_workspace,
    verify_public_path,
    detect_legacy_workspace,
    resolve_legacy_planilhas,
)
from input_bridge import request_input, CancelExecucao
from xlsm_transacional import escrever_parcial, promover, limpar_publico_antigos

# Exit codes
EXIT_CONFIG_ERROR = 2
EXIT_BASE_MISSING = 3
EXIT_BASE_NEEDS_USER = 4  # Cascata falhou — UI deve pedir planilha ao usuario
EXIT_PROCX_MISSING = 5    # Aba PROCX inexistente — nao ha como reinjetar

BASE_FILENAME_PATTERN = re.compile(r"Gerencie Carteira_(\d{4}_\d{2}_\d{2})\.xlsm$")
HTML_FILENAME_PATTERN = re.compile(r"Gerencie_Carteira_(\d{4}_\d{2}_\d{2})\.html$")

# Retencao: backups mantidos quando o config nao define [Retencao] max_arquivos.
LIMITE_BACKUPS_PADRAO = 30


def aplicar_pastas_dev(config: configparser.ConfigParser) -> None:
    """
    Em modo dev (script NAO empacotado), redireciona as pastas de [Paths] para
    uma pasta `data` local do repo — assim execucoes de teste nao gravam na
    pasta de producao da rede. No app empacotado (frozen) e no-op: ele usa os
    caminhos do config.ini.
    """
    if getattr(sys, "frozen", False):
        return
    dev_data = Path(__file__).resolve().parent.parent.parent / "data"
    config["Paths"]["pasta_destino_html"] = str(dev_data / "html")
    config["Paths"]["pasta_diario_excel"] = str(dev_data / "planilhas")
    config["Paths"]["pasta_logs"] = str(dev_data / "logs")
    config["Paths"]["pasta_copia_excel"] = str(dev_data / "publica")
    emit("info", f"Modo dev: pastas de trabalho redirecionadas para {dev_data}",
         step="config.dev")


def carregar_configuracoes(caminho_config_file: str) -> configparser.ConfigParser:
    """Le config.ini com interpolation desabilitada (para preservar %USERPROFILE%)."""
    config = configparser.ConfigParser(interpolation=None)
    if not os.path.exists(caminho_config_file):
        emit("error", f"Arquivo de configuracao nao encontrado: {caminho_config_file}",
             step="config.missing")
        sys.exit(EXIT_CONFIG_ERROR)

    config.read(caminho_config_file, encoding="utf-8")

    required = [
        ("Paths", "pasta_destino_html"),
        ("Paths", "pasta_diario_excel"),
        ("Paths", "pasta_copia_excel"),
        ("Paths", "pasta_logs"),
        ("Excel", "planilha_dados"),
        ("Excel", "coluna_verificacao"),
        ("Excel", "sheet_procx"),
        ("Excel", "col_procx_gerente"),
        ("Excel", "col_procx_cnpj"),
        ("Email", "assunto_procurado"),
    ]
    for section, key in required:
        if not config.has_option(section, key):
            emit("error", f"Chave de configuracao ausente: [{section}] {key}",
                 step="config.invalid")
            sys.exit(EXIT_CONFIG_ERROR)

    # Expande %USERPROFILE% em-place para todos os paths
    for key in ("pasta_destino_html", "pasta_diario_excel", "pasta_copia_excel",
                "pasta_logs"):
        if config.has_option("Paths", key):
            config["Paths"][key] = os.path.expandvars(config["Paths"][key])

    # Dev: redireciona as pastas de trabalho para o repo (isola dev de producao).
    aplicar_pastas_dev(config)

    emit("info", "Configuracao carregada", step="config.loaded",
         data={"paths": dict(config["Paths"])})
    return config


def configurar_logging(config: configparser.ConfigParser) -> None:
    """Configura logging para arquivo em pasta_logs."""
    pasta_logs = config["Paths"]["pasta_logs"]
    os.makedirs(pasta_logs, exist_ok=True)
    log_file = os.path.join(pasta_logs, "automacao_gerencie_carteira.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="[%Y_%m_%d %H:%M:%S]",
        handlers=[logging.FileHandler(log_file, encoding="utf-8")],
    )


def _buscar_mais_recente_em(pasta: str) -> Optional[str]:
    """Retorna o caminho da planilha mais recente em `pasta` que bate o padrao, ou None."""
    if not os.path.isdir(pasta):
        return None
    datas_arquivos: list[tuple[str, str]] = []
    for nome in os.listdir(pasta):
        m = BASE_FILENAME_PATTERN.search(nome)
        if m:
            datas_arquivos.append((m.group(1), nome))
    if not datas_arquivos:
        return None
    _, nome = max(datas_arquivos, key=lambda x: tuple(map(int, x[0].split("_"))))
    return os.path.join(pasta, nome)


def _copiar_base_para_local(origem: str, pasta_local: str) -> str:
    """Copia a base encontrada para a pasta local (idempotente)."""
    os.makedirs(pasta_local, exist_ok=True)
    destino = os.path.join(pasta_local, os.path.basename(origem))
    if os.path.abspath(origem) != os.path.abspath(destino):
        shutil.copy2(origem, destino)
    return destino


def encontrar_arquivo_base_excel(config: configparser.ConfigParser) -> str:
    """
    Localiza a planilha base em cascata:
      1. pasta_diario_excel (local)
      2. pasta legada v2.14.1 (Documents\\Gerencie Carteira\\Diário)
      3. pasta publica (A:\\PUBLICA\\...)
    Se encontrada fora do local, copia para pasta_diario_excel.
    Se nada for encontrado, emite excel.base.needs_user e sai com codigo 4.
    """
    pasta_local = config["Paths"]["pasta_diario_excel"]
    pasta_publica = config["Paths"]["pasta_copia_excel"]
    pasta_legada = resolve_legacy_planilhas()

    # 1. Local
    local = _buscar_mais_recente_em(pasta_local)
    if local:
        emit("info", f"Planilha base local encontrada: {os.path.basename(local)}",
             step="excel.base.local", data={"path": local})
        return local

    # 2. Legada v2.14.1
    if pasta_legada is not None:
        legada = _buscar_mais_recente_em(str(pasta_legada))
        if legada:
            destino = _copiar_base_para_local(legada, pasta_local)
            emit("success",
                 f"Planilha base copiada da pasta legada v2.14.1: {os.path.basename(legada)}",
                 step="excel.base.legacy",
                 data={"origem": legada, "destino": destino})
            return destino

    # 3. Publica (se acessivel)
    pub_check = verify_public_path(pasta_publica)
    if pub_check["accessible"]:
        publica = _buscar_mais_recente_em(pasta_publica)
        if publica:
            destino = _copiar_base_para_local(publica, pasta_local)
            emit("success",
                 f"Planilha base copiada da pasta publica: {os.path.basename(publica)}",
                 step="excel.base.public",
                 data={"origem": publica, "destino": destino})
            return destino

    # 4. Cascata falhou
    searched = {
        "local": pasta_local,
        "legada": str(pasta_legada) if pasta_legada else None,
        "publica": pasta_publica if pub_check["accessible"] else f"{pasta_publica} (offline)",
    }
    emit("error",
         "Nenhuma planilha base encontrada. Selecione uma planilha .xlsm manualmente.",
         step="excel.base.needs_user",
         data={"searched": searched})
    sys.exit(EXIT_BASE_NEEDS_USER)


def limpar_backups_antigos(pasta: str, padrao: re.Pattern, limite: int,
                           rotulo: str) -> int:
    """
    Mantem no maximo `limite` arquivos que casam `padrao` em `pasta`, removendo
    os mais antigos. A ordem e definida pela data YYYY_MM_DD capturada no nome
    do arquivo (grupo 1 do padrao) — independe do mtime do sistema de arquivos.

    Idempotente. Arquivos que NAO casam o padrao (templates, artefatos
    transacionais .partial/.bak) sao ignorados: nunca contados, nunca
    removidos. Retorna a quantidade de arquivos removidos.

    `rotulo` ("planilhas", "html", ...) identifica a pasta nos eventos emitidos.
    """
    if limite < 1 or not os.path.isdir(pasta):
        return 0

    datados: list[tuple[tuple[int, ...], str]] = []
    for nome in os.listdir(pasta):
        m = padrao.search(nome)
        if m:
            chave = tuple(int(p) for p in m.group(1).split("_"))
            datados.append((chave, nome))

    if len(datados) <= limite:
        return 0

    datados.sort(key=lambda item: item[0])   # mais antigo primeiro
    excedentes = datados[:-limite]            # tudo, exceto os `limite` recentes

    removidos = 0
    for _, nome in excedentes:
        try:
            os.remove(os.path.join(pasta, nome))
            removidos += 1
        except OSError as e:
            emit("warning", f"Falha ao remover backup antigo '{nome}': {e}",
                 step="retencao.fail", data={"rotulo": rotulo})

    if removidos:
        emit("step",
             f"Retencao ({rotulo}): {removidos} backup(s) antigo(s) removido(s); "
             f"mantidos os {limite} mais recentes",
             step="retencao.aplicada",
             data={"rotulo": rotulo, "removidos": removidos, "limite": limite})
    return removidos


def _eh_email_alvo(item: object, assunto: str) -> bool:
    """
    Retorna True se `item` e um MailItem nao lido com o assunto exato.
    Guarda defensiva: itens nao-MailItem (MeetingRequest, etc.) nao tem
    `.Subject` garantido — qualquer AttributeError e suprimido.
    """
    try:
        return bool(item.Subject == assunto and item.UnRead)
    except AttributeError:
        return False


def buscar_emails_novos(config: configparser.ConfigParser) -> list:
    """
    Retorna a lista de MailItems nao lidos com o assunto exato configurado.

    Otimizacao (v4.2.8): usa `Items.Restrict("[Unread] = true")` para filtrar
    server-side antes de iterar — o conjunto resultante e pequeno (apenas nao
    lidos), eliminando a varredura O(N) sobre toda a caixa.

    Fallback: se `Restrict` falhar (caixa delegada, versao antiga do Outlook,
    etc.), cai silenciosamente na varredura completa original — semantica identica,
    so muda a velocidade.
    """
    assunto = config["Email"]["assunto_procurado"]
    emit("step", "Conectando ao Outlook", step="outlook.connect")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", False)

    # Tenta Restrict server-side (filtra so nao-lidos; conjunto bem menor)
    usando_restrict = False
    try:
        conjunto = items.Restrict("[Unread] = true")
        usando_restrict = True
    except Exception:
        conjunto = items   # fallback: varredura completa original

    filtrados = [m for m in conjunto if _eh_email_alvo(m, assunto)]

    emit(
        "info",
        f"{len(filtrados)} email(s) nao lidos encontrados",
        step="outlook.fetch",
        data={"count": len(filtrados), "restrict": usando_restrict},
    )
    return filtrados


def extrair_dados_dos_anexos(emails, config: configparser.ConfigParser) -> tuple[pd.DataFrame, list]:
    """
    Parseia os anexos HTML e retorna (df, emails_processados).

    IMPORTANTE (R5): este passo NAO marca mais os e-mails como lidos. A
    marcacao `UnRead=False` foi movida para `marcar_emails_lidos()`, chamada
    apenas APOS o `wb.save()` bem-sucedido — reduzindo a janela de perda de
    dados (divida conhecida) e permitindo o input de gerentes orfaos antes de
    consumir os e-mails.
    """
    pasta_html = config["Paths"]["pasta_destino_html"]
    dados: list[list] = []
    processados: list = []
    for email in emails:
        data_email = email.ReceivedTime.date()
        try:
            for anexo in email.Attachments:
                if not anexo.FileName.endswith(".html"):
                    continue
                nome_html = f"Gerencie_Carteira_{email.ReceivedTime.strftime('%Y_%m_%d')}.html"
                caminho_html = os.path.join(pasta_html, nome_html)
                anexo.SaveAsFile(caminho_html)
                emit("info", f"Anexo salvo: {nome_html}", step="html.saved")

                with open(caminho_html, "r", encoding="utf-8") as f:
                    soup = BeautifulSoup(f, "html.parser")
                tabela = soup.find(
                    lambda t: t.name == "table" and "CNPJ" in t.get_text() and "Razão Social" in t.get_text()
                )
                if not tabela:
                    emit("error", f"Tabela nao encontrada em {nome_html}",
                         step="html.parse.fail")
                    continue
                for linha in tabela.find_all("tr")[1:]:
                    cols = [td.get_text(strip=True) for td in linha.find_all("td")]
                    if len(cols) >= 3:
                        dados.append([cols[0], cols[1], cols[2], data_email])
                processados.append(email)
                emit("info", f"Email de {data_email.strftime('%d/%m/%Y')} parseado",
                     step="html.parsed")
                break
        except Exception as e:
            emit("error", f"Falha ao processar email de {data_email}: {e}",
                 step="html.parse.error")
            continue

    if not dados:
        return pd.DataFrame(), processados

    df = pd.DataFrame(dados, columns=["CNPJ", "Razão Social", "Alteração", "Data do recebimento do e-mail"])
    df["Data do recebimento do e-mail"] = pd.to_datetime(df["Data do recebimento do e-mail"], errors="coerce")
    return df.sort_values(by="Data do recebimento do e-mail").reset_index(drop=True), processados


def marcar_emails_lidos(emails: list) -> None:
    """
    Marca os e-mails processados como lidos (`UnRead=False`).

    Invariante (R5): chamado SOMENTE apos `wb.save()` bem-sucedido. Se o save
    falhar ou a execucao for cancelada/fechada, os e-mails permanecem nao
    lidos e o dia pode ser reprocessado.
    """
    for email in emails:
        try:
            data_email = email.ReceivedTime.date()
            email.UnRead = False
            emit("info", f"Email de {data_email.strftime('%d/%m/%Y')} marcado como lido",
                 step="outlook.marked")
        except Exception as e:
            emit("warning", f"Falha ao marcar e-mail como lido: {e}",
                 step="outlook.mark.fail")


class ProcxSheetMissing(Exception):
    """A aba PROCX configurada nao existe no workbook (nao apenas vazia)."""


def _normalizar_cnpj(valor) -> str:
    """Normaliza CNPJ para comparacao (apenas digitos)."""
    if valor is None:
        return ""
    return re.sub(r"\D", "", str(valor))


def ler_mapa_procx(wb, config: configparser.ConfigParser) -> dict[str, str]:
    """
    Le a aba PROCX (config [Excel] sheet_procx) e monta o mapa
    {cnpj_normalizado: gerente}. Aba inexistente -> ProcxSheetMissing.
    Aba existente porem vazia -> mapa vazio (todos os CNPJs serao orfaos).
    """
    nome_sheet = config["Excel"]["sheet_procx"]
    col_ger = config["Excel"]["col_procx_gerente"]
    col_cnpj = config["Excel"]["col_procx_cnpj"]

    nomes = [s.name for s in wb.sheets]
    if nome_sheet not in nomes:
        raise ProcxSheetMissing(nome_sheet)

    ws = wb.sheets[nome_sheet]
    last_row = ws.range(f"{col_cnpj}{ws.cells.rows.count}").end("up").row

    mapa: dict[str, str] = {}
    if last_row < 1:
        return mapa

    gerentes = ws.range(f"{col_ger}1:{col_ger}{last_row}").value
    cnpjs = ws.range(f"{col_cnpj}1:{col_cnpj}{last_row}").value
    if not isinstance(gerentes, list):
        gerentes = [gerentes]
    if not isinstance(cnpjs, list):
        cnpjs = [cnpjs]

    for ger, cnpj in zip(gerentes, cnpjs):
        chave = _normalizar_cnpj(cnpj)
        if chave and ger not in (None, ""):
            mapa[chave] = str(ger).strip()

    emit("info", f"PROCX '{nome_sheet}': {len(mapa)} cadastro(s) CNPJ->gerente",
         step="excel.procx.loaded", data={"count": len(mapa)})
    return mapa


def detectar_orfaos(df: pd.DataFrame, mapa: dict[str, str]) -> list[dict]:
    """
    Cruza os CNPJs do df com o mapa PROCX. Um CNPJ e orfao se nao existir no
    mapa. CNPJ duplicado entre e-mails e resolvido uma unica vez (dedupe).
    """
    orfaos: list[dict] = []
    vistos: set[str] = set()
    for _, row in df.iterrows():
        cnpj_raw = row["CNPJ"]
        chave = _normalizar_cnpj(cnpj_raw)
        if not chave or chave in vistos or chave in mapa:
            continue
        vistos.add(chave)
        orfaos.append({
            "cnpj": str(cnpj_raw).strip(),
            "razao_social": str(row.get("Razão Social", "")).strip(),
        })
    emit("info", f"{len(orfaos)} CNPJ(s) orfao(s) sem gerente no PROCX",
         step="excel.orfaos.detected", data={"count": len(orfaos)})
    return orfaos


def _col_idx(letra: str) -> int:
    """Converte letra de coluna Excel (ex: 'B', 'AA') em indice 1-based."""
    idx = 0
    for ch in letra.strip().upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _obter_tabela_procx(ws, config: configparser.ConfigParser):
    """
    Retorna o ListObject (Tabela formatada) da aba PROCX, ou None.

    Se `config [Excel] tabela_procx` estiver definido, busca por nome; senao
    usa a primeira Tabela da aba (caso tipico: 1 Tabela por aba).
    """
    try:
        los = ws.api.ListObjects
        total = int(los.Count)
    except Exception:
        return None
    if total < 1:
        return None
    nome = config.get("Excel", "tabela_procx", fallback="").strip()
    if nome:
        for i in range(1, total + 1):
            lo = los.Item(i)
            if str(lo.Name) == nome:
                return lo
        return None
    return los.Item(1)


def reinjetar_procx(wb, config: configparser.ConfigParser,
                    mapping: dict[str, str]) -> int:
    """
    Acrescenta os gerentes resolvidos na aba PROCX (gerente/CNPJ nas colunas
    do config). A aba e uma Tabela formatada: a linha e inserida via API do
    ListObject (`ListRows.Add`), o que EXPANDE a Tabela e faz as colunas
    calculadas se autopreencherem — assim a referencia estruturada do PROCX
    em 'E-Mail BD' passa a cobrir a nova linha (evita o #N/D de antes).

    Fallback (aba sem Tabela): escreve celulas e replica a linha anterior.
    Retorna a quantidade de linhas inseridas.
    """
    if not mapping:
        return 0

    nome_sheet = config["Excel"]["sheet_procx"]
    col_ger = config["Excel"]["col_procx_gerente"]
    col_cnpj = config["Excel"]["col_procx_cnpj"]
    idx_ger = _col_idx(col_ger)
    idx_cnpj = _col_idx(col_cnpj)
    ws = wb.sheets[nome_sheet]

    inseridos = 0
    lo = _obter_tabela_procx(ws, config)

    if lo is not None:
        try:
            n_cols = int(lo.ListColumns.Count)
        except Exception:
            n_cols = max(idx_ger, idx_cnpj)
        for cnpj, gerente in mapping.items():
            nova = lo.ListRows.Add()  # ao final; expande a Tabela
            linha = int(nova.Range.Row)
            # `ListRows.Add` so autopreenche colunas REGISTRADAS como "coluna
            # calculada" do ListObject. Colunas com formula que perderam esse
            # status (ex.: A=XLOOKUP de matriz, D=CNPJ Numeros, E=Raiz CNPJ)
            # ficam vazias. Replicamos a formula da linha anterior via
            # Range.Copy (ajusta refs relativas e preserva formula de matriz)
            # em todas as colunas com formula, exceto as de input (gerente/CNPJ).
            prev = linha - 1
            if prev >= 1:
                for c in range(1, n_cols + 1):
                    if c in (idx_ger, idx_cnpj):
                        continue
                    src = ws.range((prev, c))
                    try:
                        tem_formula = str(src.formula).startswith("=")
                    except Exception:
                        tem_formula = False
                    if tem_formula:
                        src.copy(ws.range((linha, c)))
            ws.range((linha, idx_ger)).value = gerente
            ws.range((linha, idx_cnpj)).value = cnpj
            inseridos += 1
        emit("step",
             f"{inseridos} gerente(s) reinjetado(s) na Tabela de '{nome_sheet}'",
             step="excel.procx.reinjetado",
             data={"count": inseridos, "modo": "tabela"})
        return inseridos

    # Fallback: aba sem Tabela formatada
    emit("warning",
         f"Aba '{nome_sheet}' sem Tabela formatada — usando fallback de celulas",
         step="excel.procx.fallback")
    last_row = ws.range(f"{col_cnpj}{ws.cells.rows.count}").end("up").row
    max_col = max(ws.used_range.last_cell.column, idx_ger, idx_cnpj)
    linha = last_row + 1 if last_row >= 1 else 1
    for cnpj, gerente in mapping.items():
        prev = linha - 1
        if prev >= 1:
            for c in range(1, max_col + 1):
                if c in (idx_ger, idx_cnpj):
                    continue
                ws.range((prev, c)).copy(ws.range((linha, c)))
        ws.range((linha, idx_ger)).value = gerente
        ws.range((linha, idx_cnpj)).value = cnpj
        linha += 1
        inseridos += 1

    emit("step", f"{inseridos} gerente(s) reinjetado(s) em '{nome_sheet}'",
         step="excel.procx.reinjetado", data={"count": inseridos, "modo": "celulas"})
    return inseridos


def recalcular(app, *, full: bool) -> None:
    """
    Recalcula o workbook e espera o calculo assincrono terminar.

    `full=True` -> `CalculateFullRebuild`: reconstroi a arvore de dependencias
    INTEIRA. Necessario apos mudanca ESTRUTURAL (ex.: `ListRows.Add` no PROCX em
    reinjetar_procx). Sem isso, as referencias estruturadas da coluna de gerente
    em 'E-Mail BD' ficam transitoriamente em `#NOME?` e o `RefreshTable` CONGELA
    esse erro no cache da pivot (a fonte se auto-cura no proximo recalculo, mas a
    pivot fica quebrada).

    `full=False` -> `Calculate` normal: suficiente quando so houve colagem de
    dados (sem mudanca estrutural) e MUITO mais barato que o full rebuild — usado
    no caso comum (dia sem orfaos).

    Ambos forcam o calculo mesmo com a App em modo MANUAL (os metodos Calculate*
    ignoram o modo de calculo). `CalculateUntilAsyncQueriesDone` garante que
    nenhum calculo assincrono fique pendente. Tudo guardado: falhas de COM nao
    derrubam o pipeline.
    """
    try:
        if full:
            app.api.CalculateFullRebuild()
        else:
            app.calculate()
    except Exception:
        # Fallback: ao menos um calculo normal
        try:
            app.calculate()
        except Exception:
            pass
    try:
        app.api.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass


def atualizar_pivots(wb) -> int:
    """
    Atualiza (RefreshTable) TODAS as PivotTables de todas as abas do workbook.
    Falha em uma pivot vira warning e nao interrompe o pipeline. Retorna a
    contagem de pivots atualizadas com sucesso.
    """
    total = 0
    for sheet in wb.sheets:
        try:
            pivots = sheet.api.PivotTables()
            count = int(pivots.Count)
        except Exception:
            continue
        for i in range(1, count + 1):
            try:
                pt = pivots.Item(i)
                try:
                    pt.PivotCache().BackgroundQuery = False
                except Exception:
                    pass
                pt.RefreshTable()
                total += 1
            except Exception as e:
                emit("warning",
                     f"Falha ao atualizar pivot #{i} em '{sheet.name}': {e}",
                     step="excel.pivot.warning")
    if total:
        emit("step", f"{total} tabela(s) dinamica(s) atualizada(s)",
             step="excel.pivot.refreshed", data={"count": total})
    else:
        emit("info", "Nenhuma tabela dinamica encontrada para atualizar",
             step="excel.pivot.none")
    return total


def atualizar_planilha_excel(df: pd.DataFrame, config: configparser.ConfigParser,
                             caminho_base_excel: str) -> tuple[str, bool]:
    """
    Retorna (caminho_planilha_salva, tem_pendencias).

    Sessao xlwings unica: le PROCX -> detecta orfaos -> (se houver) pede os
    gerentes a UI via stdin -> reinjeta no PROCX -> cola em E-Mail BD ->
    salva atomicamente (.partial -> rename). A base nunca e sobrescrita
    in-place (zero-copy). DIRECIONA foi removido: #N/D residual vira apenas
    warning, nada e aberto.
    """
    nome_planilha = config["Excel"]["planilha_dados"]
    col_verificacao = config["Excel"]["coluna_verificacao"]
    pasta_diario = config["Paths"]["pasta_diario_excel"]
    pasta_copia = config["Paths"]["pasta_copia_excel"]

    data_nome = df["Data do recebimento do e-mail"].max().strftime("%Y_%m_%d")
    novo_nome = f"Gerencie Carteira_{data_nome}.xlsm"
    caminho_novo = os.path.join(pasta_diario, novo_nome)

    emit("step", f"Abrindo base Excel: {os.path.basename(caminho_base_excel)}",
         step="excel.open")

    app = None
    wb = None
    pendencias = False
    try:
        app = xw.App(visible=False, add_book=False)
        # Modo MANUAL durante a automacao: evita o recalculo automatico pesado no
        # open e a cada edicao (XLOOKUP de matriz + MINIFS/MAXIFS sobre a base que
        # cresce). Recalculamos explicitamente via recalcular() quando precisamos
        # de valores; restauramos AUTOMATICO antes de salvar (a copia publica deve
        # abrir recalculando p/ o gestor).
        try:
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass
        emit("info", "Excel iniciado", step="excel.app")
        wb = app.books.open(caminho_base_excel)
        try:
            app.api.Calculation = -4135  # reassegura apos o open (o wb pode repor)
        except Exception:
            pass
        emit("info", f"Workbook aberto: {os.path.basename(caminho_base_excel)}",
             step="excel.workbook.opened")
        time.sleep(1)

        # --- PROCX: mapa, orfaos e captura em runtime (ANTES de qualquer save
        #     e ANTES de marcar e-mails lidos — R5) ---
        mapa = ler_mapa_procx(wb, config)        # ProcxSheetMissing se ausente
        orfaos = detectar_orfaos(df, mapa)
        houve_reinjecao = False
        if orfaos:
            resposta = request_input(
                "input.gerentes.needed",
                {"orfaos": orfaos},
                msg=f"Informe o gerente de {len(orfaos)} CNPJ(s) sem cadastro",
            )
            mapping = resposta.get("mapping") or {}
            # `ListRows.Add` muda a Tabela PROCX estruturalmente -> exige full
            # rebuild adiante (senao #NOME? na pivot). O recalculo NAO e feito
            # aqui: o recalculo unico antes das pivots (condicional) ja cobre,
            # depois da colagem — evita um rebuild a mais.
            houve_reinjecao = reinjetar_procx(wb, config, mapping) > 0

        # --- Colagem em E-Mail BD (logica preservada) ---
        ws = wb.sheets[nome_planilha]
        last_row = ws.range("A" + str(ws.cells.rows.count)).end("up").row
        if last_row > 1 and not ws.range(f"{col_verificacao}{last_row}").formula.startswith("="):
            raise ValueError(f"Coluna '{col_verificacao}' nao contem formula")

        primeira_linha_vazia = ws.range("A" + str(ws.cells.rows.count)).end("up").row + 1
        ws.range(f"A{primeira_linha_vazia}").options(pd.DataFrame, index=False, header=False).value = df

        ult_linha_nova = primeira_linha_vazia + len(df) - 1

        # Copia a formula VLOOKUP para as novas linhas (evita #REF)
        linha_origem = primeira_linha_vazia - 1
        origem = ws.range(f"{col_verificacao}{linha_origem}")
        destino = ws.range(f"{col_verificacao}{primeira_linha_vazia}:{col_verificacao}{ult_linha_nova}")
        origem.copy(destino)

        emit("step", f"{len(df)} linhas inseridas", step="excel.insert",
             data={"count": int(len(df))})

        # Recalculo unico antes de ler a coluna de verificacao E das pivots.
        # Full rebuild SO quando houve reinjecao de orfaos (mudanca estrutural na
        # Tabela PROCX -> evita #NOME? no cache da pivot). No caso comum (sem
        # orfaos) um Calculate normal basta e e bem mais barato.
        recalcular(app, full=houve_reinjecao)

        verificados = ws.range(
            f"{col_verificacao}{primeira_linha_vazia}:{col_verificacao}{ult_linha_nova}"
        ).options(err_to_str=True).value
        if not isinstance(verificados, list):
            verificados = [verificados]
        if any(isinstance(v, str) and v.startswith("#") for v in verificados):
            pendencias = True

        # Atualiza tabelas dinamicas com os dados ja recalculados.
        # Roda ANTES do save local p/ que a copia publica saia ja atualizada.
        atualizar_pivots(wb)

        # Restaura AUTOMATICO antes de salvar: a copia publica DEVE abrir
        # recalculando p/ o gestor. Os valores ja estao computados (recalcular
        # acima), entao isto e um no-op barato — so persiste o modo no arquivo.
        try:
            app.api.Calculation = -4105  # xlCalculationAutomatic
        except Exception:
            pass

        # --- Fase 1: escreve os parciais com o Excel AINDA ABERTO ---
        # (a promocao/rename so e possivel apos o Excel liberar o lock)
        parcial_local = escrever_parcial(wb, caminho_novo)

        public_check = verify_public_path(pasta_copia)
        public_acessivel = public_check["accessible"]
        caminho_publico = os.path.join(pasta_copia, novo_nome)
        parcial_pub = None
        if public_acessivel:
            # Sweep antecipado: melhor esforco. Se o usuario fechar a planilha
            # antiga entre agora e a promocao, o sweep final ja nao precisa
            # remover nada. Falhas aqui sao silenciosas — o sweep final emite.
            limpar_publico_antigos(pasta_copia)
            try:
                parcial_pub = escrever_parcial(wb, caminho_publico)
            except Exception as e:
                emit("error", f"Falha ao preparar copia publica: {e}",
                     step="publico.fail")
                parcial_pub = None

    except CancelExecucao as e:
        # Fechamento/cancelamento durante o input: nada e salvo, base intacta,
        # e-mails permanecem nao lidos (marcacao so ocorre apos save no main).
        emit("warning", f"Execucao cancelada pelo usuario — nada foi gravado ({e})",
             step="excel.cancelled")
        raise
    except Exception as e:
        emit("error", f"Erro na automacao do Excel: {e}", step="excel.error")
        logging.critical(f"Erro na automacao do Excel: {e}", exc_info=True)
        raise
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()

    # --- Fase 2: Excel ja fechado -> promocao atomica (rename) ---
    promover(parcial_local, caminho_novo)
    emit("step", f"Planilha diaria salva: {caminho_novo}", step="excel.save",
         data={"path": caminho_novo})

    if pendencias:
        emit("warning",
             "Planilha salva com pendencias (#N/D na coluna de verificacao)",
             step="excel.warnings")
        logging.warning(f"Arquivo principal salvo com pendencias em: {caminho_novo}")

    if not public_acessivel:
        emit("warning", f"Pasta publica indisponivel: {pasta_copia} — etapa pulada",
             step="publico.skipped")
    elif parcial_pub is not None:
        emit("step", f"Atualizando pasta publica: {pasta_copia}", step="publico.copy")
        try:
            falhas_limpeza = limpar_publico_antigos(pasta_copia)
            for caminho_falho, err in falhas_limpeza:
                emit("warning",
                     f"Nao consegui remover planilha antiga "
                     f"'{os.path.basename(caminho_falho)}' (provavelmente aberta "
                     f"em outro Excel): {err}",
                     step="publico.cleanup.fail",
                     data={"path": caminho_falho})
            promover(parcial_pub, caminho_publico)
            emit("info", f"Copia publica salva: {caminho_publico}", step="publico.saved",
                 data={"path": caminho_publico})
        except Exception as e:
            emit("error", f"Falha ao publicar copia: {e}", step="publico.fail")

    return caminho_novo, pendencias


def _resolver_versao(base_path: Path) -> str:
    """Versao via env APP_VERSION (Electron); fallback app/package.json; senao 'dev'."""
    env_v = os.environ.get("APP_VERSION")
    if env_v:
        return env_v
    try:
        pkg = base_path.parent.parent / "app" / "package.json"
        if pkg.exists():
            return json.loads(pkg.read_text(encoding="utf-8")).get("version", "dev")
    except Exception:
        pass
    return "dev"


def _dump_timings(config: configparser.ConfigParser, versao: str) -> None:
    """
    Escreve timings_<run>.json em pasta_logs e emite perf.summary.
    Best-effort: qualquer falha (share offline, permissao) e suprimida — nunca
    derruba o pipeline. Chamada no finally de main().
    """
    try:
        timeline = get_timeline()
        if not timeline:
            return
        total_ms = timeline[-1]["t_ms"] if timeline else 0
        pasta_logs = config["Paths"]["pasta_logs"]
        ts_run = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho = os.path.join(pasta_logs, f"timings_{ts_run}.json")
        payload = {
            "version": versao,
            "total_ms": total_ms,
            "steps": timeline,
        }
        try:
            os.makedirs(pasta_logs, exist_ok=True)
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # share offline ou sem permissao — silencioso

        # top-3 etapas por dt_ms (delta individual)
        top = sorted(timeline, key=lambda e: e["dt_ms"], reverse=True)[:3]
        emit(
            "info",
            f"Execucao concluida em {total_ms} ms",
            step="perf.summary",
            data={"total_ms": total_ms, "top_steps": top},
        )
    except Exception:
        pass  # nunca propaga


def main() -> None:
    base_path = Path(__file__).resolve().parent if not getattr(sys, "frozen", False) else Path(sys.executable).resolve().parent

    # Inicia o relogio de instrumentacao (antes do primeiro emit de etapa)
    reset_timer()

    versao = _resolver_versao(base_path)
    emit("info", f"Iniciando Gerencie Carteira v{versao}", step="boot",
         data={"version": versao})

    # Config
    config_file = (base_path.parent.parent / "config" / "config.ini").resolve()
    if not config_file.exists():
        # Fallback para runtime empacotado (PyInstaller): config ao lado do .exe
        config_file = base_path / "config" / "config.ini"
    config = carregar_configuracoes(str(config_file))

    # Bootstrap workspace — a raiz de trabalho deriva do config.ini ([Paths]):
    # planilhas/html/logs vivem todas sob uma pasta `data` comum.
    data_root = os.path.dirname(config["Paths"]["pasta_logs"])
    ws = ensure_workspace(data_root)
    emit(
        "info" if ws["already_existed"] else "success",
        "Estrutura de trabalho pronta" if ws["already_existed"]
            else f"{len(ws['created'])} pasta(s) criadas",
        step="workspace.bootstrap",
        data=ws,
    )

    # Verifica pasta legada (somente informativo)
    legacy = detect_legacy_workspace()
    if legacy:
        emit("warning",
             "Pasta legada detectada — migracao de arquivos antigos e manual",
             step="workspace.legacy", data={"path": legacy})

    # Logging pra arquivo
    configurar_logging(config)

    try:
        # Verifica rota publica
        pub = verify_public_path(config["Paths"]["pasta_copia_excel"])
        if not pub["accessible"]:
            emit("warning", f"Pasta publica nao acessivel: {pub['path']}",
                 step="publico.offline")

        # Pipeline
        emails = buscar_emails_novos(config)
        if not emails:
            emit("warning", "Nenhum email nao lido com o assunto especificado", step="outlook.empty")
            emit_result("warning", None)
            return

        # IMPORTANTE: resolver a planilha base ANTES de extrair os anexos.
        # extrair_dados_dos_anexos() marca os e-mails como lidos; se a cascata
        # esgotasse depois disso (sys.exit EXIT_BASE_NEEDS_USER), o auto-rerun
        # apos o usuario escolher a base nao acharia mais e-mails nao lidos e o
        # dia de dados seria perdido silenciosamente. Resolvendo a base primeiro,
        # um exit 4 ocorre sem consumir nenhum e-mail.
        caminho_base = encontrar_arquivo_base_excel(config)

        df, emails_processados = extrair_dados_dos_anexos(emails, config)
        if df.empty:
            emit("warning", "Nenhum dado valido extraido dos anexos", step="html.empty")
            emit_result("warning", None)
            return

        try:
            caminho_salvo, pendencias = atualizar_planilha_excel(df, config, caminho_base)
        except CancelExecucao:
            # Fechamento/cancelamento durante a captura de gerentes orfaos.
            # Nada foi gravado; e-mails continuam nao lidos -> reprocessaveis.
            emit_result("warning", None)
            return
        except ProcxSheetMissing as e:
            emit("error",
                 f"Aba PROCX '{e}' inexistente — impossivel reinjetar gerentes",
                 step="excel.procx.missing")
            emit_result("error", None)
            sys.exit(EXIT_PROCX_MISSING)

        # R5: e-mails so sao marcados como lidos APOS o save bem-sucedido.
        marcar_emails_lidos(emails_processados)

        # Retencao: limita planilhas e html a ~1 mes de backup (config [Retencao]).
        # Roda apos o save bem-sucedido — a "copia n+1" dispara a remocao da mais
        # antiga. So conta arquivos que casam o padrao de data (ignora .partial/.bak).
        limite_backups = config.getint("Retencao", "max_arquivos",
                                       fallback=LIMITE_BACKUPS_PADRAO)
        limpar_backups_antigos(config["Paths"]["pasta_diario_excel"],
                               BASE_FILENAME_PATTERN, limite_backups, "planilhas")
        limpar_backups_antigos(config["Paths"]["pasta_destino_html"],
                               HTML_FILENAME_PATTERN, limite_backups, "html")

        status = "warning" if pendencias else "ok"
        emit_result(status, caminho_salvo)

    finally:
        # Best-effort: dump de timings ao fim do run (share offline nao derruba).
        _dump_timings(config, versao)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        emit("error", f"Erro global inesperado: {e}", step="fatal")
        logging.critical("Erro global inesperado", exc_info=True)
        emit_result("error", None)
        sys.exit(1)
