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
import glob
import json
import logging
import os
import re
import shutil
import sys
import time
from pathlib import Path
from typing import Optional

import pandas as pd
import win32com.client
import xlwings as xw
from bs4 import BeautifulSoup

from log_emitter import emit, emit_result
from directory_bootstrap import (
    ensure_workspace,
    verify_public_path,
    detect_legacy_workspace,
    resolve_legacy_planilhas,
)

# Exit codes
EXIT_CONFIG_ERROR = 2
EXIT_BASE_MISSING = 3
EXIT_BASE_NEEDS_USER = 4  # Cascata falhou — UI deve pedir planilha ao usuario

BASE_FILENAME_PATTERN = re.compile(r"Gerencie Carteira_(\d{4}_\d{2}_\d{2})\.xlsm$")


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
        ("Email", "assunto_procurado"),
    ]
    for section, key in required:
        if not config.has_option(section, key):
            emit("error", f"Chave de configuracao ausente: [{section}] {key}",
                 step="config.invalid")
            sys.exit(EXIT_CONFIG_ERROR)

    # Expande %USERPROFILE% em-place para todos os paths
    for key in ("pasta_destino_html", "pasta_diario_excel", "pasta_copia_excel",
                "pasta_logs", "executavel_direciona"):
        if config.has_option("Paths", key):
            config["Paths"][key] = os.path.expandvars(config["Paths"][key])

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


def buscar_emails_novos(config: configparser.ConfigParser) -> list:
    assunto = config["Email"]["assunto_procurado"]
    emit("step", "Conectando ao Outlook", step="outlook.connect")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", False)
    filtrados = [m for m in messages if m.Subject == assunto and m.UnRead]
    emit("info", f"{len(filtrados)} email(s) nao lidos encontrados",
         step="outlook.fetch", data={"count": len(filtrados)})
    return filtrados


def extrair_dados_dos_anexos(emails, config: configparser.ConfigParser) -> pd.DataFrame:
    pasta_html = config["Paths"]["pasta_destino_html"]
    dados: list[list] = []
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
                email.UnRead = False
                emit("info", f"Email de {data_email.strftime('%d/%m/%Y')} marcado como lido",
                     step="outlook.marked")
                break
        except Exception as e:
            emit("error", f"Falha ao processar email de {data_email}: {e}",
                 step="html.parse.error")
            continue

    if not dados:
        return pd.DataFrame()

    df = pd.DataFrame(dados, columns=["CNPJ", "Razão Social", "Alteração", "Data do recebimento do e-mail"])
    df["Data do recebimento do e-mail"] = pd.to_datetime(df["Data do recebimento do e-mail"], errors="coerce")
    return df.sort_values(by="Data do recebimento do e-mail").reset_index(drop=True)


def atualizar_planilha_excel(df: pd.DataFrame, config: configparser.ConfigParser,
                             caminho_base_excel: str) -> tuple[str, bool]:
    """Retorna (caminho_planilha_salva, tem_pendencias)."""
    nome_planilha = config["Excel"]["planilha_dados"]
    col_verificacao = config["Excel"]["coluna_verificacao"]
    executavel = config.get("Paths", "executavel_direciona", fallback=None)
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
        wb = app.books.open(caminho_base_excel)
        time.sleep(1)

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

        # Forca recalculo antes de ler a coluna de verificacao.
        # As formulas VLOOKUP recem-coladas podem nao ter recalculado ainda
        # (calc assincrono / modo manual), causando falso negativo/positivo
        # na deteccao de pendencias #N/D.
        try:
            app.calculate()
        except Exception:
            pass
        time.sleep(1)

        verificados = ws.range(
            f"{col_verificacao}{primeira_linha_vazia}:{col_verificacao}{ult_linha_nova}"
        ).options(err_to_str=True).value
        if not isinstance(verificados, list):
            verificados = [verificados]
        if any(isinstance(v, str) and v.startswith("#") for v in verificados):
            pendencias = True

        wb.save(caminho_novo)
        emit("step", f"Planilha diaria salva: {caminho_novo}", step="excel.save",
             data={"path": caminho_novo})

        if pendencias:
            emit("warning", "Planilha salva com pendencias (#N/D na coluna de verificacao)",
                 step="excel.warnings")
            logging.warning(f"Arquivo principal salvo com pendencias em: {caminho_novo}")
            if executavel and os.path.exists(executavel):
                os.startfile(executavel)

        # Copia publica
        public_check = verify_public_path(pasta_copia)
        if not public_check["accessible"]:
            emit("warning", f"Pasta publica indisponivel: {pasta_copia} — etapa pulada",
                 step="publico.skipped")
        else:
            emit("step", f"Atualizando pasta publica: {pasta_copia}", step="publico.copy")
            try:
                for f in glob.glob(os.path.join(pasta_copia, "Gerencie*.xls*")):
                    try:
                        os.remove(f)
                        emit("info", f"Removido: {os.path.basename(f)}", step="publico.cleanup")
                    except OSError as e:
                        emit("error", f"Falha ao remover {f}: {e}", step="publico.cleanup.fail")

                caminho_publico = os.path.join(pasta_copia, novo_nome)
                wb.save(caminho_publico)
                emit("info", f"Copia publica salva: {caminho_publico}", step="publico.saved",
                     data={"path": caminho_publico})
            except Exception as e:
                emit("error", f"Falha ao publicar copia: {e}", step="publico.fail")

    except Exception as e:
        emit("error", f"Erro na automacao do Excel: {e}", step="excel.error")
        logging.critical(f"Erro na automacao do Excel: {e}", exc_info=True)
        raise
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()

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


def main() -> None:
    base_path = Path(__file__).resolve().parent if not getattr(sys, "frozen", False) else Path(sys.executable).resolve().parent

    versao = _resolver_versao(base_path)
    emit("info", f"Iniciando Gerencie Carteira v{versao}", step="boot",
         data={"version": versao})

    # Config
    config_file = (base_path.parent.parent / "config" / "config.ini").resolve()
    if not config_file.exists():
        # Fallback para runtime empacotado (PyInstaller): config ao lado do .exe
        config_file = base_path / "config" / "config.ini"
    config = carregar_configuracoes(str(config_file))

    # Bootstrap workspace
    ws = ensure_workspace()
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

    df = extrair_dados_dos_anexos(emails, config)
    if df.empty:
        emit("warning", "Nenhum dado valido extraido dos anexos", step="html.empty")
        emit_result("warning", None)
        return

    caminho_salvo, pendencias = atualizar_planilha_excel(df, config, caminho_base)

    status = "warning" if pendencias else "ok"
    emit_result(status, caminho_salvo)


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
