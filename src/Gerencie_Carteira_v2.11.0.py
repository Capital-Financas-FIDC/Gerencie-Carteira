# ==============================================================================
# SCRIPT DE AUTOMAÇÃO GERENCIE CARTEIRA v2.11.0
#
# NOTAS DA VERSÃO:
# - Adicionada funcionalidade para criar uma cópia da planilha gerada em uma
#   pasta separada, removendo macros e salvando como .xlsx.
# - O script agora deleta a cópia do dia anterior na pasta de destino.
# ==============================================================================

import win32com.client
import os
import pandas as pd
import re
import xlwings as xw
import sys
import configparser
import time
import logging
import glob
from bs4 import BeautifulSoup
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.traceback import Traceback

# Inicializa o console globalmente para ser usado por qualquer função
console = Console()

def configurar_logging(config):
    """Configura o sistema de logging para salvar em arquivo e exibir no console."""
    FORMATO_LOG = "%(asctime)s - %(levelname)s - %(message)s"
    try:
        caminho_pasta_logs = config.get('Paths', 'pasta_logs')
    except (configparser.NoSectionError, configparser.NoOptionError):
        caminho_pasta_logs = None
    if caminho_pasta_logs:
        os.makedirs(caminho_pasta_logs, exist_ok=True)
        log_file = os.path.join(caminho_pasta_logs, 'automacao_gerencie_carteira.log')
    else:
        script_path = os.path.dirname(os.path.realpath(__file__))
        log_file = os.path.join(script_path, 'automacao_gerencie_carteira.log')
    logging.basicConfig(
        level="INFO",
        format=FORMATO_LOG,
        datefmt="[%Y_%m_%d %H:%M:%S]",
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8')
        ]
    )

def carregar_configuracoes(caminho_config_file):
    """
    Lê o arquivo config.ini e retorna um objeto de configuração.
    Encerra o script se o arquivo ou uma chave essencial não for encontrado.
    """
    config = configparser.ConfigParser()
    if not os.path.exists(caminho_config_file):
        console.print(f"[bold red]ERRO CRÍTICO: Arquivo '{caminho_config_file}' não encontrado.[/bold red]")
        logging.critical(f"ERRO CRÍTICO: Arquivo '{caminho_config_file}' não encontrado.")
        sys.exit("Pressione ENTER para sair")
    config.read(caminho_config_file, encoding='utf-8')
    try:
        _ = config['Paths']['pasta_destino_html']
        _ = config['Paths']['pasta_diario_excel']
        _ = config['Paths']['pasta_copia_excel'] # <-- VERIFICA A NOVA CHAVE
        _ = config['Excel']['planilha_dados']
        _ = config['Excel']['coluna_verificacao']
        _ = config['Email']['assunto_procurado']
    except KeyError as e:
        trace = Traceback(show_locals=True, word_wrap=True, width=120)
        console.print(f"\n[on red]ERRO CRÍTICO: Chave de configuração '{e.args[0]}' não encontrada no 'config.ini'[/on red]")
        console.print(trace)
        logging.critical(f"ERRO CRÍTICO: Chave de configuração não encontrada no arquivo config.ini")
        sys.exit("Pressione ENTER para sair")
    return config

def encontrar_arquivo_base_excel(config):
    """
    Analisa a pasta de diários e retorna o caminho completo para a planilha
    Excel mais recente que servirá de base.
    """
    pasta_diario = config['Paths']['pasta_diario_excel']
    padrao_nome = re.compile(r"Gerencie Carteira_(\d{4}_\d{2}_\d{2})")
    arquivos_diario = os.listdir(pasta_diario)
    datas_encontradas = [match.group(1) for arquivo in arquivos_diario if (match := padrao_nome.search(arquivo))]
    if not datas_encontradas:
        console.print(f"[bold yellow]Nenhum arquivo base encontrado na pasta:[/bold yellow] [cyan]{pasta_diario}[/cyan]")
        logging.critical(f"Nenhum arquivo base encontrado na pasta: {pasta_diario}")
        os.startfile(pasta_diario)
        sys.exit("Pressione ENTER para sair")
    ultima_data_str = max(datas_encontradas, key=lambda d: tuple(map(int, d.split('_'))))
    nome_arquivo_base = f"Gerencie Carteira_{ultima_data_str}.xlsm"
    return os.path.join(pasta_diario, nome_arquivo_base)

def buscar_emails_novos(config):
    """
    Conecta-se ao Outlook e retorna uma lista de e-mails não lidos
    que correspondem ao assunto procurado.
    """
    assunto = config['Email']['assunto_procurado']
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", False)
    emails_filtrados = [msg for msg in messages if msg.Subject == assunto and msg.UnRead]
    return emails_filtrados

def extrair_dados_dos_anexos(emails, config):
    """
    Itera sobre os e-mails, salva seus anexos HTML, extrai os dados,
    marca o e-mail como lido e retorna um DataFrame pandas consolidado.
    """
    pasta_destino_html = config['Paths']['pasta_destino_html']
    dados_extraidos = []
    for email in emails:
        data_email = email.ReceivedTime.date()
        try:
            for anexo in email.Attachments:
                if anexo.FileName.endswith(".html"):
                    data_nome_arquivo = email.ReceivedTime.strftime("%Y_%m_%d")
                    nome_arquivo_html = f"Gerencie_Carteira_{data_nome_arquivo}.html"
                    caminho_arquivo_html = os.path.join(pasta_destino_html, nome_arquivo_html)
                    anexo.SaveAsFile(caminho_arquivo_html)
                    console.print(f"Anexo salvo: [green]'{nome_arquivo_html}'[/green]")
                    with open(caminho_arquivo_html, "r", encoding="utf-8") as f:
                        soup = BeautifulSoup(f, "html.parser")
                    tabela_dados = soup.find(lambda tag: tag.name == 'table' and 'CNPJ' in tag.get_text() and 'Razão Social' in tag.get_text())
                    if not tabela_dados:
                        console.print(f"[bold red]ERRO: Nenhuma tabela com 'CNPJ' e 'Razão Social' encontrada em '{nome_arquivo_html}'.[/bold red]")
                        logging.critical(f"ERRO: Nenhuma tabela com 'CNPJ' e 'Razão Social' encontrada em '{nome_arquivo_html}'.")
                        os.startfile(caminho_arquivo_html)
                        continue
                    for linha in tabela_dados.find_all("tr")[1:]:
                        colunas = [td.get_text(strip=True) for td in linha.find_all("td")]
                        if len(colunas) >= 3:
                            dados_extraidos.append([colunas[0], colunas[1], colunas[2], data_email])
                    email.UnRead = False
                    console.print(f"E-mail de [cyan]{data_email.strftime('%d/%m/%Y')}[/cyan] marcado como lido.")
                    break
        except Exception:
            trace = Traceback(show_locals=True, word_wrap=True, width=120)
            console.print(f"\n[on red]ERRO CRÍTICO: Erro ao processar o email de {data_email} [/on red]")
            console.print(trace)
            logging.critical(f"ERRO CRÍTICO: Erro ao processar o email de {data_email}")
            continue
    if not dados_extraidos: return pd.DataFrame()
    df = pd.DataFrame(dados_extraidos, columns=["CNPJ", "Razão Social", "Alteração", "Data do recebimento do e-mail"])
    df["Data do recebimento do e-mail"] = pd.to_datetime(df["Data do recebimento do e-mail"], errors='coerce')
    return df.sort_values(by="Data do recebimento do e-mail").reset_index(drop=True)

def apresentar_dados_no_console(df):
    """
    Recebe um DataFrame e exibe uma tabela formatada no console.
    """
    tabela_rich = Table(title="Empresas Monitoradas para Atualização", show_lines=True)
    tabela_rich.add_column("CNPJ", justify="center")
    tabela_rich.add_column("Razão Social", justify="left", min_width=30)
    tabela_rich.add_column("Alteração", justify="left")
    tabela_rich.add_column("Data E-mail", justify="center")
    for _, row in df.iterrows():
        alteracao_norm = str(row["Alteração"]).strip().upper()
        if "INCLUSAO" in alteracao_norm:
            cor = "bold red"
        elif "EXCLUSAO" in alteracao_norm:
            cor = "bold #1cb900"
        else:
            cor = "bold yellow"
        tabela_rich.add_row(
            str(row["CNPJ"]),
            str(row["Razão Social"]),
            f"[{cor}]{row['Alteração']}[/]",
            row["Data do recebimento do e-mail"].strftime('%d/%m/%Y')
        )
    console.print("\n")
    console.print(tabela_rich)

def atualizar_planilha_excel(df, config, caminho_base_excel):
    """
    Executa a manipulação do Excel: abre a base, insere dados, salva o novo arquivo,
    e cria uma cópia .xlsx sem macros em outra pasta, substituindo a antiga.
    """
    nome_planilha = config['Excel']['planilha_dados']
    col_verificacao = config['Excel']['coluna_verificacao']
    executavel = config.get('Paths', 'executavel_direciona', fallback=None) # Usar get com fallback
    pasta_diario = config['Paths']['pasta_diario_excel']
    pasta_copia = config['Paths']['pasta_copia_excel']

    data_nome = df['Data do recebimento do e-mail'].max().strftime('%Y_%m_%d')
    novo_nome_arquivo = f"Gerencie Carteira_{data_nome}.xlsm"
    caminho_novo_excel = os.path.join(pasta_diario, novo_nome_arquivo)

    console.print(f"\nIniciando automação do Excel. Base: [cyan]{os.path.basename(caminho_base_excel)}[/cyan]")
    
    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(caminho_base_excel)
        time.sleep(1)

        ws = wb.sheets[nome_planilha]
        last_row = ws.range('A' + str(ws.cells.rows.count)).end('up').row
        if last_row > 1 and not ws.range(f'{col_verificacao}{last_row}').formula.startswith('='):
            raise ValueError(f"Coluna '{col_verificacao}' não contém fórmula.")
        
        primeira_linha_vazia = ws.range('A' + str(ws.cells.rows.count)).end('up').row + 1
        ws.range(f'A{primeira_linha_vazia}').options(pd.DataFrame, index=False, header=False).value = df
        
        erro_encontrado = False
        ult_linha_nova = primeira_linha_vazia + len(df) - 1
        dados_verificados = ws.range(f'{col_verificacao}{primeira_linha_vazia}:{col_verificacao}{ult_linha_nova}').options(err_to_str=True).value
        if not isinstance(dados_verificados, list): dados_verificados = [dados_verificados]
        if any(isinstance(v, str) and v.startswith('#') for v in dados_verificados):
            erro_encontrado = True
        
        # Salva o arquivo principal com macros (.xlsm)
        wb.save(caminho_novo_excel)
        
        if erro_encontrado:
            console.print(Panel.fit(f"Arquivo principal salvo com pendências em:\n[yellow]{caminho_novo_excel}[/yellow]", title="Ação Necessária", border_style="yellow"))
            logging.warning(f"Arquivo principal salvo com pendências em: {caminho_novo_excel}")
            if executavel and os.path.exists(executavel): os.startfile(executavel)
        else:
            console.print(Panel.fit(f"Arquivo principal salvo com sucesso em:\n[green]{caminho_novo_excel}[/green]", title="Sucesso", border_style="green"))

        # --- INÍCIO DA FUNCIONALIDADE TEMP ---
        
        # 1. Deletar cópia antiga (.xlsx) na pasta de destino
        console.print(f"Verificando arquivos antigos na pasta de cópia: [cyan]{pasta_copia}[/cyan]")
        arquivos_antigos = glob.glob(os.path.join(pasta_copia, "*.xlsx"))
        for f in arquivos_antigos:
            try:
                os.remove(f)
                console.print(f"Arquivo antigo removido: [yellow]{os.path.basename(f)}[/yellow]")
            except OSError:
                traceos = Traceback(show_locals=True, word_wrap=True, width=120)
                console.print(f"[bold red]ERRO ao remover o arquivo antigo '{f}': {e}[/bold red]")
                console.print(traceos)
                logging.error(f"ERRO ao remover o arquivo antigo '{f}': {e}")


        # 2. Criar o nome e caminho para a nova cópia .xlsx
        nome_copia_xlsx = f"Gerencie Carteira_{data_nome}.xlsx"
        caminho_copia_xlsx = os.path.join(pasta_copia, nome_copia_xlsx)

        # 3. Salvar a cópia sem macros (.xlsx)
        # O xlwings remove as macros automaticamente ao salvar de .xlsm para .xlsx
        wb.save(caminho_copia_xlsx)
        console.print(Panel.fit(f"Cópia sem macros salva com sucesso em:\n[green]{caminho_copia_xlsx}[/green]", title="Cópia Gerada", border_style="green"))
        
        # --- FIM DA FUNCIONALIDADE TEMP ---

    except Exception:
        trace = Traceback(show_locals=True, word_wrap=True, width=120)
        console.print("\n[on red]ERRO CRÍTICO: Ocorreu um erro inesperado na automação do Excel![/on red]")
        console.print(trace)
        logging.critical(f"ERRO CRÍTICO: Ocorreu um erro inesperado na automação do Excel!")
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()
        console.print("[dim]Processo do Excel finalizado com segurança.[/dim]")


def main():
    """
    Função principal que orquestra a execução do script.
    """
    console.print(Panel.fit("[bold cyan]Iniciando Automação 'Gerencie Carteira' v2.11.0[/bold cyan]"))
    
    # 1. Carregar Configurações
    script_dir = os.path.dirname(os.path.realpath(__file__))
    root_dir = os.path.dirname(script_dir)
    config_file = os.path.join(root_dir, 'config', 'config.ini')
    config = carregar_configuracoes(config_file)

    # 2. Configurar Logging
    configurar_logging(config) 

    # 3. Buscar E-mails
    emails = buscar_emails_novos(config)
    if not emails:
        console.print("[bold yellow]Nenhum e-mail não lido encontrado com o assunto especificado.[/bold yellow]")
        logging.info("Nenhum e-mail não lido encontrado.")
        return 

    console.print(f"Encontrados [bold green]{len(emails)}[/bold green] e-mail(s) para processar.")
    
    # 4. Extrair Dados dos Anexos
    df_dados = extrair_dados_dos_anexos(emails, config)
    if df_dados.empty:
        console.print("[bold yellow]Nenhum dado válido foi extraído dos anexos.[/bold yellow]")
        logging.warning("Nenhum dado válido foi extraído dos anexos.")
        return
        
    # 5. Apresentar Dados no Console
    apresentar_dados_no_console(df_dados)
    
    # 6. Encontrar Base e Atualizar Excel
    caminho_base = encontrar_arquivo_base_excel(config)
    atualizar_planilha_excel(df_dados, config, caminho_base)

    console.print("\n[bold cyan]Automação finalizada.[/bold cyan]")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        # Captura saídas limpas (como "Pressione ENTER para sair") para evitar Traceback
        if e.code is not None:
             console.print(f"\n[yellow]Script encerrado pelo usuário ou por erro previsto.[/yellow]")
    except Exception:
        # Captura qualquer outra exceção inesperada
        console.print("\n[on red]UM ERRO GLOBAL INESPERADO OCORREU:[/on red]")
        console.print(Traceback(show_locals=True, word_wrap=True, width=120))
        logging.critical("Um erro global inesperado ocorreu", exc_info=True)
    finally:
        input("\nPressione ENTER para sair")