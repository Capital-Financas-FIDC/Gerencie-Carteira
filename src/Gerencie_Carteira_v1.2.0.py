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

# Ordenar e-mails por data, do mais antigo ao mais recente
messages.Sort("[ReceivedTime]", False)

# Filtrar apenas os e-mails não lidos com o assunto especificado
emails_nao_lidos = [msg for msg in messages if msg.Subject == assunto_procurado and msg.UnRead]

if emails_nao_lidos:
    pasta_destino = r"C:\Users\comercial05\Documents\Gerencia Carteira\HTML"
    caminho_excel = r"C:\Users\comercial05\Desktop\PROGRAMAS PYTHON\TESTE PYTHON\TESTE.xlsx"
    dados = []

    for email in emails_nao_lidos:
        # Extrair a data do e-mail e formatá-la
        data_email = email.ReceivedTime.strftime("%Y/%m/%d")

        # Baixar o anexo .html
        caminho_arquivo = None
        for anexo in email.Attachments:
            if anexo.FileName.endswith(".html"):
                nome_arquivo = f"Gerencie_Carteira_{data_email.replace('/', '_')}.html"
                caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
                anexo.SaveAsFile(caminho_arquivo)
                print(f"Anexo salvo em: {caminho_arquivo}")
                break

        if caminho_arquivo:
            # Abrir o arquivo HTML e extrair os dados da tabela
            with open(caminho_arquivo, "r", encoding="utf-8") as arquivo:
                soup = BeautifulSoup(arquivo, "html.parser")

            tabelas = soup.find_all("table")
            if len(tabelas) > 1:
                tabela = tabelas[1]
                tbody = tabela.find("tbody")
                linhas = tbody.find_all("tr") if tbody else tabela.find_all("tr")

                for linha in linhas:
                    colunas = linha.find_all("td")
                    if len(colunas) >= 3:
                        cnpj = colunas[0].get_text(strip=True)
                        razao_social = colunas[1].get_text(strip=True)
                        alteracao = colunas[2].get_text(strip=True)

                        if "CNPJ" not in cnpj and "Razão Social" not in razao_social and "Alteração" not in alteracao:
                            dados.append([cnpj, razao_social, alteracao, data_email])

    if dados:
        # Criar um DataFrame do pandas com a nova coluna "Data da Operação"
        df = pd.DataFrame(dados, columns=["CNPJ", "Razão Social", "Alteração", "Data da Operação"])

        # Ordenar os dados do mais antigo para o mais novo
        df = df.sort_values(by="Data da Operação", ascending=True)

        print(df)

        # Escrever os dados no Excel
        with pd.ExcelWriter(caminho_excel, mode="w", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Dados", index=False, header=True)

        # Carregar o arquivo Excel e excluir a primeira linha
        wb = load_workbook(caminho_excel)
        ws = wb["Dados"]
        ws.delete_rows(1)  # Exclui a primeira linha (cabeçalhos)

        # Salvar as alterações no arquivo Excel
        wb.save(caminho_excel)
        print(f"Dados copiados para {caminho_excel}, com os mais antigos primeiro e a primeira linha removida.")
    else:
        print("Nenhuma tabela válida encontrada nos arquivos HTML.")
else:
    print("Nenhum e-mail não lido encontrado com o assunto especificado.")