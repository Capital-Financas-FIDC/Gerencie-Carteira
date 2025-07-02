#Mantém todas as funcionalidades da versão anterior

#Notas da v2.8.1:
# -Adiciona verificações para melhoria de robustez

import win32com.client
import os
from bs4 import BeautifulSoup
import pandas as pd
import re
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
import xlwings as xw
import sys
import configparser

# Inicializa o objeto Console da biblioteca rich para impressões coloridas e formatadas no terminal.
console = Console()

# --- INÍCIO DO BLOCO DE CONFIG---

# Pega o caminho absoluto da pasta onde este script Python está localizado.
# __file__ é uma variável especial que contém o caminho do próprio script.
pasta_script = os.path.dirname(os.path.realpath(__file__))

# Cria o caminho completo e exato para o arquivo config.ini.
caminho_config = os.path.join(pasta_script, 'config.ini')

config = configparser.ConfigParser()
# Verifica se o arquivo config.ini existe antes de tentar ler
if not os.path.exists(caminho_config):
    console.print("[bold red]ERRO CRÍTICO: Arquivo 'config.ini' não encontrado.[/bold red]")
    console.print("Certifique-se de que o arquivo de configuração está na mesma pasta que o script.")
    input("\nPressione ENTER para sair")
    sys.exit()

config.read(caminho_config, encoding='utf-8')

# Acessa as configurações usando as seções e chaves do arquivo
try:
    # Seção [Paths]
    pasta_destino_html = config['Paths']['pasta_destino_html']
    pasta_diario_excel = config['Paths']['pasta_diario_excel']
    executavel_direciona = config['Paths']['executavel_direciona']
    # Seção [Excel]
    nome_planilha_dados = config['Excel']['planilha_dados']
    coluna_verificacao = config['Excel']['coluna_verificacao']
    # Seção [Email]
    assunto_procurado = config['Email']['assunto_procurado']
except KeyError as e:
    console.print(f"[bold red]ERRO CRÍTICO: Chave de configuração não encontrada no 'config.ini': {e}[/bold red]")
    console.print("Verifique se o seu 'config.ini' contém todas as chaves necessárias.")
    input("\nPressione ENTER para sair")
    sys.exit()
# --- FIM DO BLOCO DE CONFIG ---

# Conectar ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Acessar a caixa de entrada (pasta Inbox)
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

# Ordenar e-mails por data, do mais antigo ao mais recente
messages.Sort("[ReceivedTime]", False)

# Filtrar apenas os e-mails não lidos com o assunto especificado
emails_nao_lidos = [msg for msg in messages if msg.Subject == assunto_procurado and msg.UnRead]

# Verifica se há e-mails não lidos com o assunto procurado
if emails_nao_lidos:
    # Definir o caminho para a pasta onde os arquivos HTML serão salvos
    pasta_destino = pasta_destino_html
    dados = []

    # Extrair a data do e-mail mais recente (embora não seja utilizada posteriormente para nomear o arquivo HTML individualmente,
    # a lógica para a data do Excel é mais robusta e independente da data do e-mail)
    data_mais_recente = max(email.ReceivedTime for email in emails_nao_lidos)

    # Definir o caminho para a pasta das planilhas diárias do Excel
    pasta_diario = pasta_diario_excel
    
    # Lista todos os arquivos na pasta 'Diário'
    arquivos_diario = os.listdir(pasta_diario)
    
    # Compila um padrão de expressão regular para filtrar arquivos Excel com o nome desejado e extrair a data
    # O padrão foi ajustado para AAAA_MM_DD
    padrao_nome_excel = re.compile(r"Gerencie Carteira_(\d{4}_\d{2}_\d{2})")
    
    datas_encontradas_lista = []

    # Itera sobre os arquivos na pasta 'Diário' para encontrar a data da última planilha existente
    for arquivo_diario in arquivos_diario:
        match = padrao_nome_excel.search(arquivo_diario)
        if match:
            datas_encontradas_lista.append(match.group(1))
    if not datas_encontradas_lista:
        console.print("[bold yellow]Nenhum arquivo com data foi encontrado na pasta! Verifique manualmente[/bold yellow]")
        console.print(f"Pasta verificada: [yellow]{pasta_diario_excel}[/yellow]")
        os.startfile(pasta_diario)
        input("\nPressione ENTER para sair")
        sys.exit() # Sair se nenhum arquivo base for encontrado

    # Se datas de arquivos Excel foram encontradas, determina a data da planilha mais recente
    if datas_encontradas_lista:
        # A chave de ordenação foi ajustada para lidar com o formato AAAA_MM_DD
        ultima_data_diario = max(datas_encontradas_lista, key=lambda x: tuple(map(int, x.split('_'))))
        # Constrói o nome completo do arquivo Excel mais recente
        nome_arquivo_excel = f"Gerencie Carteira_{ultima_data_diario}.xlsm"
        # Constrói o caminho completo para o arquivo Excel mais recente
        caminho_excel = os.path.join(pasta_diario, nome_arquivo_excel)
    else:
        # Informa se nenhuma planilha foi encontrada e abre a pasta para verificação manual
        console.print("[bold yellow]Nenhuma planilha encontrada na pasta! Verifique manualmente[/bold yellow]")
        os.startfile(pasta_diario)

    # Processa cada e-mail não lido encontrado
    for email in emails_nao_lidos:
        # Extrai a data do e-mail
        data_email = email.ReceivedTime.date() # Formata para a coluna "Data da Operação" como DD/MM/AAAA
        data_email_nome = email.ReceivedTime.strftime("%Y_%m_%d") # Formata para o nome do arquivo HTML/Excel como AAAA_MM_DD
        caminho_arquivo = None
        
        # Itera sobre os anexos do e-mail para encontrar o arquivo HTML
        for anexo in email.Attachments:
            if anexo.FileName.endswith(".html"):
                # Define o nome do arquivo HTML a ser salvo com a data formatada como AAAA_MM_DD
                nome_arquivo = f"Gerencie_Carteira_{data_email_nome}.html"
                # Constrói o caminho completo para salvar o anexo
                caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
                # Salva o anexo
                anexo.SaveAsFile(caminho_arquivo)
                console.print(f"Anexo salvo como: [green]'{nome_arquivo}'[/green] na pasta HTML's do Gerencie Carteira")
                print("\n")
                break

        # Se o arquivo HTML foi salvo, abre-o e extrai os dados da tabela
        if caminho_arquivo:
            with open(caminho_arquivo, "r", encoding="utf-8") as arquivo:
                soup = BeautifulSoup(arquivo, "html.parser")

            # --- BUSCA INTELIGENTE PELA TABELA DE DADOS ---
            # Em vez de pegar a "segunda tabela", procuramos pela tabela que contenha
            # os textos 'CNPJ' e 'Razão Social', que são únicos da tabela de dados.
            tabela_dados = soup.find(lambda tag: tag.name == 'table' and 
                                                 'CNPJ' in tag.get_text() and 
                                                 'Razão Social' in tag.get_text())
            # Após a busca, verificamos se a tabela foi realmente encontrada.
            if tabela_dados:
                # Se encontrou a tabela, o código prossegue normalmente.
                tbody = tabela_dados.find("tbody")
                linhas = tbody.find_all("tr") if tbody else tabela_dados.find_all("tr")

                # Itera sobre as linhas da tabela para extrair os dados
                for linha in linhas:
                    colunas = linha.find_all("td")
                    # Verifica se há pelo menos 3 colunas (CNPJ, Razão Social, Alteração)
                    if len(colunas) >= 3:
                        cnpj = colunas[0].get_text(strip=True)
                        razao_social = colunas[1].get_text(strip=True)
                        alteracao = colunas[2].get_text(strip=True)

                        # Adiciona os dados à lista se não forem os cabeçalhos da tabela
                        if "CNPJ" not in cnpj and "Razão Social" not in razao_social and "Alteração" not in alteracao:
                            dados.append([cnpj, razao_social, alteracao, data_email])
            
            else:
                # Se a tabela de dados não foi encontrada no arquivo, informa o erro.
                console.print(f"[bold red]ERRO no arquivo '{nome_arquivo}':[/bold red]")
                console.print("[yellow]Nenhuma tabela com os cabeçalhos 'CNPJ' e 'Razão Social' foi encontrada.[/yellow]")
                os.startfile(caminho_arquivo) # Abre o arquivo HTML para inspeção manual.
            
            # --- FIM DA BUSCA INTELIGENTE ---

    # Se houver dados extraídos, processa-os e insere-os no Excel
    if dados:
        # Cria um DataFrame do pandas com os dados e a nova coluna "Data da Operação"
        df = pd.DataFrame(dados, columns=["CNPJ", "Razão Social", "Alteração", "Data do recebimento do e-mail"])
        
         # Converte a coluna "Data da Operação" para uma data propriamente
        df["Data do recebimento do e-mail"]=pd.to_datetime(df["Data do recebimento do e-mail"], errors='coerce')       
        
        # Converte as colunas "CNPJ", "Razão Social" e "Alteração"
        colunas_para_string = ["CNPJ", "Razão Social", "Alteração"]
        for coluna in colunas_para_string:
            df[coluna] = df[coluna].astype(str)
        

        
        # Ordena os dados pela "Data da Operação" do mais antigo para o mais novo
        df = df.sort_values(by="Data do recebimento do e-mail", ascending=True)

        # Configura a tabela para exibição no terminal usando a biblioteca 'rich'
        tabela_rich = Table(title="Empresas Monitoradas", show_lines=True)
        tabela_rich.add_column("CNPJ", no_wrap=True, justify="center")
        tabela_rich.add_column("Razão Social", no_wrap=True, justify="center")
        tabela_rich.add_column("Alteração", justify="center")
        tabela_rich.add_column("Data do recebimento do e-mail", justify="center")

        # Popula a tabela 'rich' com os dados do DataFrame, aplicando formatação condicional à coluna "Alteração"
        for _, row in df.iterrows():
            cnpj_rich = str(row["CNPJ"]).strip()
            razao_social_rich = str(row["Razão Social"]).strip()
            alteracao_rich = str(row["Alteração"]).strip()
            data_rich = str(row["Data do recebimento do e-mail"]).strip()

            alteracao_normalizada = alteracao_rich.upper().replace('  ', ' ')

            # Aplica cores diferentes para as alterações de inclusão, exclusão ou outras
            if alteracao_normalizada == "INCLUSAO ANOT.INADIMPLENCIA":
                alteracao_rich_formatada = f"[bold red]{alteracao_rich}[/bold red]"
            elif alteracao_normalizada == "EXCLUSAO ANOT.INADIMPLENCIA":
                alteracao_rich_formatada = f"[bold #1cb900]{alteracao_rich}[/bold #1cb900]"
            else:
                alteracao_rich_formatada = f"[bold yellow]{alteracao_rich}[/bold yellow]"
            
            tabela_rich.add_row(cnpj_rich, razao_social_rich, alteracao_rich_formatada, data_rich)

        # Imprime a tabela formatada no terminal
        console.print(tabela_rich)

        try:
            # --- PARTE 1: PREPARAÇÃO DOS NOMES E CAMINHOS ---
            # Essa parte define o nome do NOVO arquivo com base na data mais recente dos dados.
            data_email_nome = df['Data do recebimento do e-mail'].max().strftime('%Y_%m_%d')
            novo_nome_arquivo = f"Gerencie Carteira_{data_email_nome}.xlsm"
            caminho_novo_excel = os.path.join(pasta_diario, novo_nome_arquivo)

            # A sua mensagem de log também está no lugar certo.
            console.print(f"\nIniciando automação do Excel para criar o arquivo: [green]{novo_nome_arquivo}[/green]")
            console.print(f"Abrindo o arquivo base: [green]{nome_arquivo_excel}[/green]") # Adicionei um log extra para clareza

            # --- PARTE 2: MANIPULAÇÃO DO EXCEL (Lógica com xlwings) ---
            with xw.App(visible=False) as app:
                # Abre o arquivo MAIS RECENTE JÁ EXISTENTE para usar como base.
                # A variável 'caminho_excel' já foi definida no config.ini
                wb = app.books.open(caminho_excel)

                # --- INÍCIO DA VERIFICAÇÃO DE INTEGRIDADE ---
                try:
                    # Verifica se a aba especificada existe
                    ws_dados = wb.sheets[nome_planilha_dados] 

                    ultima_linha_com_dados = ws_dados.range('A' + str(ws_dados.cells.rows.count)).end('up').row

                    # Só checa se houver mais do que apenas o cabeçalho
                    if ultima_linha_com_dados > 1:
                        celula_para_checar = f'{coluna_verificacao}{ultima_linha_com_dados}'
                        formula_da_celula = ws_dados.range(celula_para_checar).formula

                        if not formula_da_celula.startswith('='):
                            # Se não começar com '=', não é uma fórmula. É um erro crítico.
                            raise ValueError(f"A célula de verificação '{celula_para_checar}' não contém uma fórmula do Excel.")
                except (KeyError, ValueError) as e:
                    # Se qualquer uma das verificações falhar, o erro será capturado aqui
                    console.print(f"\n[bold red]ERRO CRÍTICO: O arquivo base '{nome_arquivo_excel}' falhou na verificação de integridade.[/bold red]")
                    console.print(f"[yellow]Detalhe do erro: {e}[/yellow]")
                    console.print("O script não pode continuar. Por favor, corrija o arquivo base.")
                    wb.close()
                    os.startfile(caminho_excel) # Abre o arquivo problemático para o usuário
                    sys.exit() # Interrompe a execução
                # --- FIM DA VERIFICAÇÃO DE INEGRIDADE ---
                    
                # Se passou pela verificação, o script continua normalmente
                console.print("\n[bold green]Arquivo base verificado com sucesso. Prosseguindo com a atualização...[/bold green]")
                    
                # Encontra a primeira linha vazia e insere os dados
                primeira_linha_vazia = ws_dados.range('A' + str(ws_dados.cells.rows.count)).end('up').row + 1
                num_novas_linhas = len(df)

                if num_novas_linhas > 0:
                    ultima_linha_nova = primeira_linha_vazia + num_novas_linhas - 1
                    intervalo_texto = f'A{primeira_linha_vazia}:C{ultima_linha_nova}'
                    ws_dados.range(intervalo_texto).number_format = '@'
                    ws_dados.range(f'A{primeira_linha_vazia}').options(pd.DataFrame, index=False, header=False).value = df
                    # Verificação do erro #N/D
                    intervalo_verif= f'{coluna_verificacao}{primeira_linha_vazia}:{coluna_verificacao}{ultima_linha_nova}'
                    dados_a_verificar = ws_dados.range(intervalo_verif).options(err_to_str=True).value
                        
                    erro_encontrado = False
                    # Garante que dados_coluna_d seja sempre uma lista para o loop
                    if not isinstance(dados_a_verificar, list):
                        dados_a_verificar = [dados_a_verificar]

                    for i, valor_celula in enumerate(dados_a_verificar):
                        if isinstance(valor_celula, str) and valor_celula.startswith('#'):
                            linha_do_erro = primeira_linha_vazia + i
                            endereco_celula = f'{coluna_verificacao}{linha_do_erro}'
                            console.print(f"[bold red]O erro #N/D foi encontrado na célula {endereco_celula} da planilha '{ws_dados.name}'.[/bold red]")
                            erro_encontrado = True
                    # Lógica final de salvar e avisar o usuário
                    if erro_encontrado:
                        wb.save(caminho_novo_excel)
                        console.print(Panel.fit(f"Arquivo salvo em :\n[yellow]{caminho_novo_excel}[/yellow]", title="Cedente sem gerente", border_style="yellow"))
                        os.startfile(executavel_direciona)
                    else:
                        wb.save(caminho_novo_excel)
                        console.print(Panel.fit(f"Arquivo salvo e atualizado com sucesso em:\n[green]{caminho_novo_excel}[/green]", title="Sucesso", border_style="green"))
                    
                else: # Caso o DataFrame venha vazio
                    console.print(Panel.fit(f"[red]Gerencie Carteira veio vazio hoje![/red]", title="AVISO", border_style="yellow"))
                    # Decida se quer sair ou apenas não criar um arquivo novo
                    sys.exit()
                
                wb.close()

        except Exception as e:
            console.print(f"[bold red]Ocorreu um erro inesperado durante a automação do Excel:[/bold red]")
            console.print(e)
    else:
        # Informa se nenhuma tabela válida foi encontrada nos arquivos HTML
        console.print("[bold yellow]Nenhuma tabela válida encontrada nos arquivos HTML.[/bold yellow]")
        os.startfile(caminho_arquivo)
else:
    # Informa se nenhum e-mail não lido com o assunto especificado foi encontrado
    console.print("[bold yellow]Nenhum e-mail não lido encontrado com o assunto especificado.[/bold yellow]")

# Aguarda a entrada do usuário antes de sair para que as mensagens no terminal possam ser lidas
input("\nPressione ENTER para sair")