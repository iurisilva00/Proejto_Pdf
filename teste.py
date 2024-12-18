

import os

import time

import re

import tempfile

from datetime import datetime

from office365.sharepoint.client_context import ClientContext

from office365.sharepoint.files.file import File

from office365.runtime.auth.user_credential import UserCredential

from dotenv import load_dotenv

import fitz  # PyMuPDF

import pandas as pd

import pytesseract

from PIL import Image

import cv2





pytesseract.pytesseract.tesseract_cmd = r'C:\tess\tesseract.exe'

# Carregar variáveis de ambiente do arquivo .env

load_dotenv()

password = os.getenv("password")

username = "00027336@progen.com.br"

site_url = os.getenv("site_url")



# Obter data e hora atuais

now = datetime.now()

ano = now.strftime('%Y')

mes = now.strftime('%m')

numero_medicao = "10"

ultima_medicao = f"{numero_medicao}{mes}{ano}"

erros = []



# Função para criar o contexto do cliente com reconexão

def criar_contexto():

    user_credentials = UserCredential(username, password)

    return ClientContext(site_url).with_credentials(user_credentials)



# Inicializar o contexto do cliente

ctx = criar_contexto()



# Função para buscar arquivos em pastas e subpastas no SharePoint

from requests.exceptions import HTTPError



def buscar_arquivos_em_pastas(folder_url):

    tentativas = 0

    max_tentativas = 3



    while tentativas < max_tentativas:

        try:

            folder = ctx.web.get_folder_by_server_relative_url(folder_url)

            ctx.load(folder)

            ctx.execute_query()

           

            files = folder.files

            ctx.load(files)

            ctx.execute_query()



            for file in files:

                if re.search(r'NFS', file.properties['Name'], re.IGNORECASE) and file.properties['Name'].endswith('.pdf'):

                    print(f"Encontrado arquivo: {file.properties['Name']}")

                    yield file



            subfolders = folder.folders

            ctx.load(subfolders)

            ctx.execute_query()



            for subfolder in subfolders:

                time.sleep(2)  # Intervalo entre subpastas

                yield from buscar_arquivos_em_pastas(subfolder.properties['ServerRelativeUrl'])



            break  # Sai do loop se tudo der certo

        except HTTPError as e:

            if e.response.status_code == 429:

                tentativas += 1

                print(f"Erro 429: Tentando novamente ({tentativas}/{max_tentativas})...")

                time.sleep(5)  # Espera antes de tentar novamente

            else:

                raise  # Re-levanta o erro se não for 429



# Função para extrair texto do PDF usando PyMuPDF

def extrair_texto_pdf(file_url):
    global ctx
    tentativas = 0
    max_tentativas = 3

    while tentativas < max_tentativas:
        try:
            # Baixa o PDF do SharePoint
            response = File.open_binary(ctx, file_url)

            # Cria um arquivo temporário para armazenar o PDF baixado
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(response.content)
                temp_pdf_path = temp_pdf.name

            # Verificar se o arquivo foi escrito corretamente
            if not os.path.exists(temp_pdf_path):
                raise FileNotFoundError(f"Arquivo temporário não encontrado: {temp_pdf_path}")

            texto_extraido = ""
            texto_ocr = ""

            # Extrair texto diretamente do PDF usando PyMuPDF
            with fitz.open(temp_pdf_path) as pdf:
                for page in pdf:
                    texto_pagina = page.get_text()
                    texto_extraido += texto_pagina  # Adiciona o texto extraído diretamente

            # Se o texto extraído for vazio ou menor que 100 caracteres, usar OCR
            if not texto_extraido.strip() or len(texto_extraido) < 100:
                with fitz.open(temp_pdf_path) as pdf:
                    for page in pdf:
                        # Converte a página em imagem para usar OCR
                        pix = page.get_pixmap()
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        texto_ocr += pytesseract.image_to_string(img, lang="por")  # OCR com idioma definido
                texto_extraido += texto_ocr  # Adiciona o texto OCR ao resultado final

            # Remover arquivos temporários
            os.remove(temp_pdf_path)

            print(f"Texto extraído do arquivo {file_url.split('/')[-1]}:\n{texto_extraido[:200]}...")  # Exibe os primeiros 200 caracteres

            return (file_url.split('/')[-1], texto_extraido)

        except Exception as e:
            print(f"Erro ao processar texto no arquivo {file_url.split('/')[-1]}: {e}")
            tentativas += 1
            if tentativas < max_tentativas:
                print(f"Tentando reconectar... (tentativa {tentativas} de {max_tentativas})")
                ctx = criar_contexto()
                time.sleep(2)
            else:
                print(f"Falha após {max_tentativas} tentativas.")
                erros.append(file_url.split('/')[-1])
    return (file_url.split('/')[-1], "")


# URL da pasta inicial no SharePoint onde os PDFs estão armazenados

initial_folder_url = '/sites/GerenciamentodaConstruo2/Documentos Compartilhados/General/11 - FATURAMENTO'



# Buscar arquivos no SharePoint e processá-los

print("Buscando arquivos no SharePoint...")

arquivos = list(buscar_arquivos_em_pastas(initial_folder_url))

print(f"Total de arquivos encontrados: {len(arquivos)}")



# Extrair texto de cada arquivo PDF encontrado

resultados = []

for arquivo in arquivos:

    file_url = arquivo.properties["ServerRelativeUrl"]

    nome_arquivo, texto = extrair_texto_pdf(file_url)

    resultados.append({"Arquivo": nome_arquivo, "Texto": texto})



def sanitize_column_names(df):

    # Converta os nomes das colunas para strings

    df.columns = df.columns.astype(str)

    # Substitua caracteres inválidos nos nomes das colunas

    df.columns = df.columns.str.replace(r'[^A-Za-z0-9_]', '', regex=True)

    return df





# Exportar os dados extraídos para um arquivo Excel

print("Exportando resultados...")

df_resultados = pd.DataFrame(resultados)

df_resultados = sanitize_column_names(df_resultados)

df_resultados.to_csv("dados_extraidos.csv", index=False)

print("Arquivo Excel exportado com sucesso!")



# Exportar lista de arquivos com erro

print("Exportando arquivos com erro...")

df_erros = pd.DataFrame(erros)

df_erros = sanitize_column_names(df_erros)

df_erros.to_csv("arquivos_com_erro.csv", index=False)

print("Arquivos com erro exportados com sucesso!")



