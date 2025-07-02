# ==============================================================================
# SCRIPT DE AUTOMAÇÃO GERENCIE CARTEIRA v2.9.0 (Refatorado)
#
# NOTAS DA VERSÃO 2.9.0:
# - Código totalmente refatorado em funções modulares para maior
#   legibilidade, manutenibilidade e testabilidade.
# - Corrigido o erro intermitente 'NoneType' object has no attribute 'Cells'.
# - Mantém todas as funcionalidades das versões anteriores
# ==============================================================================

import win32com.client
import os
import pandas as pd
import re
import xlwings as xw
import sys
import configparser
from bs4 import BeautifulSoup
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
import time

# Inicializa o console globalmente para ser usado por qualquer função
console = Console()

def carregar_configuracoes(caminho_config_file):
    """
    Lê o arquivo config.ini e retorna um objeto de configuração.
    Encerra o script se o arquivo ou uma chave essencial não for encontrado.
    """
    config = configparser.ConfigParser()
    if not os.path.exists(caminho_config_file):
        console.print(f"[bold red]ERRO CRÍTICO: Arquivo '{caminho_config_file}' não encontrado.[/bold red]")
        sys.exit("Pressione ENTER para sair")

    config.read(caminho_config_file, encoding='utf-8')

    try:
        # Valida se as seções e chaves necessárias existem
        _ = config['Paths']['pasta_destino_html']
        _ = config['Paths']['pasta_diario_excel']
        _ = config['Excel']['planilha_dados']
        _ = config['Excel']['coluna_verificacao']
        _ = config['Email']['assunto_procurado']
    except KeyError as e:
        console.print(f"[bold red]ERRO CRÍTICO: Chave de configuração não encontrada no 'config.ini': {e}[/bold red]")
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
    inbox = outlook.GetDefaultFolder(6)  # 6 = Caixa de Entrada
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
                    
                    # --- LINHA CORRIGIDA ---
                    # Esta busca é mais robusta. Ela procura uma tag <table> que em seu
                    # conteúdo de texto (incluindo todas as sub-tags) tenha as palavras-chave.
                    tabela_dados = soup.find(lambda tag: tag.name == 'table' and 
                                                         'CNPJ' in tag.get_text() and 
                                                         'Razão Social' in tag.get_text())
                    
                    if not tabela_dados:
                        console.print(f"[bold red]ERRO: Nenhuma tabela com 'CNPJ' e 'Razão Social' encontrada em '{nome_arquivo_html}'.[/bold red]")
                        # Abre o arquivo para inspeção manual
                        os.startfile(caminho_arquivo_html)
                        continue

                    for linha in tabela_dados.find_all("tr")[1:]:
                        colunas = [td.get_text(strip=True) for td in linha.find_all("td")]
                        if len(colunas) >= 3:
                            dados_extraidos.append([colunas[0], colunas[1], colunas[2], data_email])
                    
                    email.UnRead = False
                    console.print(f"E-mail de [cyan]{data_email.strftime('%d/%m/%Y')}[/cyan] marcado como lido.")
                    break
        except Exception as e:
            console.print(f"[bold red]Erro ao processar e-mail de {data_email}: {e}[/bold red]")
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
    Executa toda a lógica de manipulação do Excel: abre a base, verifica a
    integridade, insere os dados, verifica erros e salva o novo arquivo.
    VERSÃO ROBUSTA: Garante uma nova instância do Excel e adiciona pausas.
    """
    nome_planilha = config['Excel']['planilha_dados']
    col_verificacao = config['Excel']['coluna_verificacao']
    executavel = config['Paths']['executavel_direciona']
    pasta_diario = config['Paths']['pasta_diario_excel']
    data_nome = df['Data do recebimento do e-mail'].max().strftime('%Y_%m_%d')
    novo_nome_arquivo = f"Gerencie Carteira_{data_nome}.xlsm"
    caminho_novo_excel = os.path.join(pasta_diario, novo_nome_arquivo)

    console.print(f"\nIniciando automação do Excel. Base: [cyan]{os.path.basename(caminho_base_excel)}[/cyan]")
    
    app = None
    wb = None
    try:
        # --- MUDANÇA 1: GARANTIR UMA INSTÂNCIA NOVA E ISOLADA ---
        # add_book=False inicia o Excel sem nenhuma pasta de trabalho aberta.
        app = xw.App(visible=False, add_book=False)
        
        # Agora abrimos nossa pasta de trabalho nesta instância limpa.
        wb = app.books.open(caminho_base_excel)

        # --- MUDANÇA 2: PAUSA DE SEGURANÇA ---
        # Dá ao Excel 1 segundo para processar completamente a abertura do arquivo.
        time.sleep(1) 

        # 1. Verificação de Integridade (continua igual)
        try:
            ws = wb.sheets[nome_planilha]
            last_row = ws.range('A' + str(ws.cells.rows.count)).end('up').row
            if last_row > 1 and not ws.range(f'{col_verificacao}{last_row}').formula.startswith('='):
                raise ValueError(f"Coluna '{col_verificacao}' não contém fórmula.")
        except Exception as e:
            console.print(f"[bold red]ERRO: Planilha base '{os.path.basename(caminho_base_excel)}' inválida. {e}[/bold red]")
            os.startfile(caminho_base_excel)
            return

        # O resto da lógica permanece o mesmo...
        primeira_linha_vazia = ws.range('A' + str(ws.cells.rows.count)).end('up').row + 1
        ws.range(f'A{primeira_linha_vazia}').options(pd.DataFrame, index=False, header=False).value = df

        erro_encontrado = False
        ult_linha_nova = primeira_linha_vazia + len(df) - 1
        dados_verificados = ws.range(f'{col_verificacao}{primeira_linha_vazia}:{col_verificacao}{ult_linha_nova}').options(err_to_str=True).value
        if not isinstance(dados_verificados, list): dados_verificados = [dados_verificados]
        if any(isinstance(v, str) and v.startswith('#') for v in dados_verificados):
            erro_encontrado = True
        
        wb.save(caminho_novo_excel)
        if erro_encontrado:
            console.print(Panel.fit(f"Arquivo salvo com pendências em:\n[yellow]{caminho_novo_excel}[/yellow]", title="Ação Necessária", border_style="yellow"))
            if os.path.exists(executavel): os.startfile(executavel)
        else:
            console.print(Panel.fit(f"Arquivo salvo e atualizado com sucesso em:\n[green]{caminho_novo_excel}[/green]", title="Sucesso", border_style="green"))

    except Exception as e:
        console.print(f"[bold red]Erro inesperado na automação do Excel:[/bold red]\n{e}")
    finally:
        # --- MUDANÇA 3: LIMPEZA ROBUSTA ---
        # Garante que a pasta de trabalho seja fechada antes de fechar o app.
        if wb:
            wb.close()
        if app:
            app.quit()
        console.print("[dim]Processo do Excel finalizado com segurança.[/dim]")

def main():
    """
    Função principal que orquestra a execução do script.
    """
    console.print(Panel.fit("[bold cyan]Iniciando Automação 'Gerencie Carteira'[/bold cyan]"))
    
    # 1. Carregar Configurações
    script_path = os.path.dirname(os.path.realpath(__file__))
    config_file = os.path.join(script_path, 'config.ini')
    config = carregar_configuracoes(config_file)

    # 2. Buscar E-mails
    emails = buscar_emails_novos(config)
    if not emails:
        console.print("[bold yellow]Nenhum e-mail não lido encontrado com o assunto especificado.[/bold yellow]")
        return # Encerra a função principal

    console.print(f"Encontrados [bold green]{len(emails)}[/bold green] e-mail(s) para processar.")
    
    # 3. Extrair Dados dos Anexos
    df_dados = extrair_dados_dos_anexos(emails, config)
    if df_dados.empty:
        console.print("[bold yellow]Nenhum dado válido foi extraído dos anexos.[/bold yellow]")
        return
        
    # 4. Apresentar Dados no Console
    apresentar_dados_no_console(df_dados)
    
    # 5. Encontrar Base e Atualizar Excel
    caminho_base = encontrar_arquivo_base_excel(config)
    atualizar_planilha_excel(df_dados, config, caminho_base)

    console.print("\n[bold cyan]Automação finalizada.[/bold cyan]")


if __name__ == "__main__":
    main()
    input("\nPressione ENTER para sair")