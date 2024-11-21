import pandas as pd
import numpy as np
# Define o número máximo de colunas e linhas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
import warnings      # Importa o módulo warnings para controlar avisos
warnings.filterwarnings("ignore")  # Configura para ignorar avisos
from pyspark.sql.functions import col
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from io import BytesIO
import xlrd
import datetime
from datetime import datetime, timedelta
from locale import setlocale, LC_TIME
from requests_ntlm import HttpNtlmAuth
import pandas as pd
import locale
from datetime import datetime
import re
import os
import unidecode
import camelot
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential
import camelot
import re
import time
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
import fitz  # PyMuPDF
import time
from dotenv import load_dotenv

load_dotenv()
password = os.getenv("password")
username = os.getenv("username")
site_url = os.getenv("site_url")
now= datetime.now()
dia =now.day
ano = now.strftime('%Y')
mes = now.strftime('%m')
numero_medicao="10"
ultima_medicao = f"{numero_medicao}{mes}{ano}"


# Função para criar o contexto do cliente com reconexão
def criar_contexto():
    user_credentials = UserCredential(username, password)
    return ClientContext(site_url).with_credentials(user_credentials)

# Inicializar o contexto do cliente
ctx = criar_contexto()

# Lista para armazenar os dados extraídos dos PDFs
dados_extraidos = []

# Função para buscar arquivos em pastas e subpastas
def buscar_arquivos_em_pastas(folder_url):
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    ctx.load(folder)
    ctx.execute_query()
    
    # Carregar e iterar sobre os arquivos na pasta atual
    files = folder.files
    ctx.load(files)
    ctx.execute_query()

    for file in files:
        if re.search(r'620|RF', file.properties['Name'], re.IGNORECASE) and file.properties['Name'].endswith('.pdf'):
            yield file

    # Carregar e iterar sobre subpastas
    subfolders = folder.folders
    ctx.load(subfolders)
    ctx.execute_query()
    for subfolder in subfolders:
        time.sleep(1)
        yield from buscar_arquivos_em_pastas(subfolder.properties['ServerRelativeUrl'])

# Função para extrair texto da primeira página do PDF
def extrair_texto_pdf(file_url):
    global ctx
    tentativas = 0
    max_tentativas = 3
    
    while tentativas < max_tentativas:
        try:
            # Baixa o PDF do SharePoint
            response = File.open_binary(ctx, file_url)
            
            # Salva o PDF temporariamente
            with open("temp.pdf", "wb") as temp_pdf:
                temp_pdf.write(response.content)
            
            # Extrair texto da primeira página usando PyMuPDF
            try:
                pdf_document = fitz.open("temp.pdf")
                
                print(f"Processando o arquivo: {file_url.split('/')[-1]}")
                page = pdf_document[0]  # Apenas a primeira página
                text = page.get_text("text")  # Extrai o texto em formato simples
                print(f"Texto da primeira página:\n{text}")
                
                # Armazenar os dados extraídos
                dados_extraidos.append({
                    'Arquivo': file_url.split('/')[-1],
                    'Texto': text
                })
                
                pdf_document.close()

            except Exception as e:
                print(f"Erro ao processar texto no arquivo {file_url.split('/')[-1]}: {e}")
            break
        except Exception as e:
            print(f"Erro ao acessar o arquivo {file_url.split('/')[-1]}: {e}")
            tentativas += 1
            if tentativas < max_tentativas:
                print(f"Tentando reconectar... (tentativa {tentativas} de {max_tentativas})")
                ctx = criar_contexto()
                time.sleep(2)
            else:
                print(f"Falha após {max_tentativas} tentativas.")

# URL da pasta inicial no SharePoint
initial_folder_url = '/sites/GerenciamentodaConstruo2/Documentos Compartilhados/General/11 - FATURAMENTO'

# Função para processar arquivos usando ThreadPoolExecutor
def processar_arquivos_concurrently():
    with ThreadPoolExecutor(max_workers=5) as executor:  # Ajuste `max_workers` conforme necessário
        futures = []
        for arquivo in buscar_arquivos_em_pastas(initial_folder_url):
            print(f"Enfileirando arquivo: {arquivo.properties['Name']}")
            futures.append(executor.submit(extrair_texto_pdf, arquivo.properties["ServerRelativeUrl"]))
        
        # Aguardar a conclusão de todos os arquivos
        for future in futures:
            future.result()

# Função para exportar dados para CSV
def exportar_para_csv():
    # Converter os dados extraídos para um DataFrame
    df = pd.DataFrame(dados_extraidos)
    
    # Exportar para CSV
    df.to_csv('dados_extraidos.csv', index=False, encoding='utf-8')
    print("Arquivo CSV exportado com sucesso!")

# Iniciar a execução
if __name__ == "__main__":
    processar_arquivos_concurrently()
    exportar_para_csv()