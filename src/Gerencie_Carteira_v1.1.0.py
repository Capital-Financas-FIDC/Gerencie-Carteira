import win32com.client
import os
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook

# Conectando ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Acessando a caixa de entrada
inbox = outlook.GetDefaultFolder(6)  # 6 corresponde à pasta Inbox
messages = inbox.Items

# Definir o assunto do e-mail para pesquisa
assunto_procurado = "Gerencie Carteira - Consulte as Empresas Monitoradas"

# Ordenar e-mails por data, do mais recente ao mais antigo
messages.Sort("[ReceivedTime]", True)

# Buscar o e-mail mais recente com o assunto especificado
email_encontrado = None
for message in messages:
    if message.Subject == assunto_procurado:
        email_encontrado = message
        break

if email_encontrado:
    # Extrair a data do e-mail e formatá-la
    data_email = email_encontrado.ReceivedTime.strftime("%Y_m_%d")
    
    # Pasta de destino
    pasta_destino = r"C:\Users\comercial05\Documents\Gerencia Carteira\HTML"
    
    # Baixar o anexo .html
    caminho_arquivo = None
    for anexo in email_encontrado.Attachments:
        if anexo.FileName.endswith(".html"):
            nome_arquivo = f"Gerencie_Carteira_{data_email}.html"
            caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
            anexo.SaveAsFile(caminho_arquivo)
            print(f"Anexo salvo em: {caminho_arquivo}")
            break

    if caminho_arquivo:
        # Abrir o arquivo HTML e extrair os dados da tabela
        with open(caminho_arquivo, "r", encoding="utf-8") as arquivo:
            soup = BeautifulSoup(arquivo, "html.parser")

        tabelas = soup.find_all("table")
        # Encontrar a tabela no HTML que contém as informações necessárias
        tabela = tabelas[1]

        # Encontrar o corpo da tabela (tbody)
        tbody = tabela.find("tbody")
        linhas = tbody.find_all("tr") if tbody else tabela.find_all("tr")

        # Extração de dados, ignorando cabeçalhos
        dados = []
        for linha in linhas:
            colunas = linha.find_all("td")
            if len(colunas) >= 3:
                cnpj = colunas[0].get_text(strip=True)
                razao_social = colunas[1].get_text(strip=True)
                alteracao = colunas[2].get_text(strip=True)

                # Ignorar a linha dos títulos
                if "CNPJ" not in cnpj and "Razão Social" not in razao_social and "Alteração" not in alteracao:
                    dados.append([cnpj, razao_social, alteracao])

        # Criar um DataFrame do pandas
        df = pd.DataFrame(dados, columns=["CNPJ", "Razão Social", "Alteração"])
        print(df)

        # Definir caminho para o arquivo Excel
        caminho_excel = r"C:\Users\comercial05\Desktop\PROGRAMAS PYTHON\TESTE PYTHON\TESTE.xlsx"

        # Escrever os dados no Excel
        with pd.ExcelWriter(caminho_excel, mode="w", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Dados", index=False, header=True)

        # Carregar o arquivo Excel e excluir as células A1, B1 e C1
        wb = load_workbook(caminho_excel)
        ws = wb["Dados"]

        # Excluir a primeira linha
        ws.delete_rows(1)

        # Salvar as alterações no arquivo Excel
        wb.save(caminho_excel)

        print(f"Dados copiados para {caminho_excel} e a primeira linha foi removida.")
    else:
        print("Nenhuma tabela encontrada no arquivo HTML.")
else:
    print("Nenhum e-mail encontrado com o assunto especificado.")