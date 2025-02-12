import os
import camelot
import pandas as pd
import logging
logging.basicConfig(level=logging.INFO)
from  configs.regras import rules_dict
from configs.conect_sharepoint import executa_arquivo
from configs.criar_contexto import executa_conexao
import os
import time
import re
import tempfile
import tabula
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential
import io
from dotenv import load_dotenv
from requests.exceptions import HTTPError
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)
import requests
from configs.buscandoarquivos import arquivosFim
from configs.criar_contextoGraph import criar_contextoGraph
from configs.conecta_list import processadata
class PDFExtract:
    def __init__(self, pdf_bytes, configs):

        self.configs = configs
        self.pdf_bytes = pdf_bytes
        
    

    def start(self):
        """Processa o PDF automaticamente e remove o arquivo temporário após uso."""
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(self.pdf_bytes.getvalue())  
            temp_pdf_path = temp_pdf.name  # Obtém o caminho do arquivo temporário
        
        

        try:
            df = self.get_table_data(temp_pdf_path, self.configs["table_area_0"], self.configs["columns_0"])
        except Exception as e:
            
            df = None
        finally:
            os.remove(temp_pdf_path)  # 🔥 Remove o arquivo TEMPORÁRIO automaticamente
            

        return df  


    def get_table_data(self,pdf_path,t_area,t_columns):
       
        df_list = []
        tables = camelot.read_pdf(
            pdf_path,
            flavor = self.configs['flavor'],
            table_areas = t_area,
            columns = t_columns,
            row_tol=40,
            pages=self.configs["page"]
        )
        logging.info(f"Transforma em DataFrame")
        df = tables[0].df
        # 🔹 Concatena todas as colunas em uma única string para busca
        text = "".join(df[1].astype(str).tolist())
    
        # 🔹 Usa Regex para extrair os valores corretos
        data = {
            "N_RF": re.search(r'Nº RF:\s*(\d+)', text),
            "Data_Geracao": re.search(r'relatório:\s*(\d{2}\.\d{2}\.\d{4})', text),
            "Numero_contrato": re.search(r'Nº Contrato:\s*(\d+)', text),
            "Numero_Pedido": re.search(r'Nº Pedido/item:\s*([\d/]+)', text),
            "Numero_FRS": re.search(r'Nº FRS:\s*(\d+)', text),
            "Medicao": re.search(r'(Período.*?\d{2}\.\d{2}\.\d{4}\s*a\s*\d{2}\.\d{2}\.\d{4})', text, re.DOTALL),
            "Numero_fornecedor": re.search(r'Cod. Fornecedor:\s*(\d+)', text),
            "Codigo_servico": re.search(r'Valor R\$.*?\n\s*(\d+)', text, re.DOTALL),
            "Descricao_servico":re.search(r'DESCRIÇÃO SERVIÇO.*?\n(.*?)(?=\n*Valor do Serviço\(s\) \(Bruto\))', text, re.DOTALL),
            "LC116": re.search(r'Valor R\$.*?(\d+\.\d+)', text, re.DOTALL),
            "Valor_R": re.search(r'Valor R\$.*?\d{1,3}(?:\.\d{3})*,\d{2}.*?(\d{1,3}(?:\.\d{3})*,\d{2})', text, re.DOTALL),
            "Valor_Bruto": re.search(r'Valor do Serviço\(s\) \(Bruto\)\s*\n\s*(\d{1,3}(?:\.\d{3})*,\d{2})', text),
            "IRRF_Valor": re.search(r'IRRF:\s*SIM\s*(\d+,\d+)', text),
            "ISS_valor": re.search(r'ISS:\s*SIM\s*(\d+,\d+)', text),
            "PIS_valor": re.search(r'PIS:\s*SIM\s*(\d+,\d+)', text),
            "INSS_valor": re.search(r'INSS:\s*(SIM|NÃO)\s*(\d*,\d+)?', text),
            "COFINS_valor": re.search(r'COFINS:\s*SIM\s*(\d+,\d+)', text),
            "INSS_ad_Sat_Valor": re.search(r'INSS Ad\(SAT\):\s*(SIM|NÃO)\s*(\d*,\d+)?', text),
            "CSLL_valor": re.search(r'CSLL:\s*SIM\s*(\d+,\d+)', text)
        }

        # 🔹 Converte `MatchObject` para string ou usa `"**"` se não encontrar
        data = {key: match.group(1) if match else "**" for key, match in data.items()}
   

        # 🔹 Processa cada linha, quebrando células que têm '\n' e mantendo estrutura
        # new_rows = []
        # for index, row in df.iterrows():
        #     split_cells = [str(cell).split("\n") for cell in row]  # Divide células que contêm "\n"
            
        #     # 🔹 Preenche espaços vazios com "**"
        #     split_cells = [[x if x.strip() else "**" for x in cell] for cell in split_cells]

        #     # 🔹 Mantém estrutura uniforme (mesmo número de elementos por linha)
        #     max_len = max(len(cell) for cell in split_cells)  # Maior número de elementos
        #     split_cells = [cell + ["**"] * (max_len - len(cell)) for cell in split_cells]  # Preenche células vazias
        #     new_rows.extend(zip(*split_cells))  # Reorganiza como tabela
        
        # # 🔹 Cria novo DataFrame mantendo as colunas corretamente alinhadas

        # df_cleaned = df.copy()
        # print(df_cleaned[1])
        # # 🔹 Garante que **todos os espaços vazios** estejam preenchidos
        # df_cleaned.fillna("**", inplace=True)
        # df_cleaned.replace("", "**", inplace=True)  # Para strings vazias
        # df_cleaned.to_excel('teste.xlsx', index=False)
        # df_cleaned["N_RF"] = str(row[1]).split("\n")[0]  
        # df_cleaned["Data_Geracao"] = row[1].split("\n")[-1]
        # df_cleaned["Numero_contrato"] = row[1].split("\n")[0]
        # df_cleaned["Numero_Pedido"] = row[1].split("\n")[0]
        # df_cleaned["Numero_FRS"] = row[1].split("\n")[0]
        # df_cleaned["Medicao"] = row[1]
        # df_cleaned.to_excel('teste.xlsx', index=False)
        # # df_cleaned['N_RF'] = df_cleaned.iloc[1, 1]#
        # # df_cleaned['Data_Geracao'] = df_cleaned.iloc[3, 1]#
        # # df_cleaned['Numero_contrato'] = df_cleaned.iloc[6, 1]#
        # # df_cleaned['Numero_Pedido'] = df_cleaned.iloc[9, 1]#
        # # df_cleaned['Numero_FRS'] = df_cleaned.iloc[12, 1]#
        # # df_cleaned['Medicao'] = df_cleaned.iloc[14, 1]#
        # # df_cleaned['Numero_fornecedor'] = df_cleaned.iloc[19, 1]#
        # # df_cleaned['Codigo_servico'] = df_cleaned.iloc[34, 1]
        # # df_cleaned['Valor_'] = df_cleaned.iloc[37, 1]
        # # df_cleaned['Valor_bruto'] = df_cleaned.iloc[40, 1]
        # # df_cleaned['Lei'] = df_cleaned.iloc[36, 1]
        # # df_cleaned['Descricao_servico'] = df_cleaned.iloc[35, 1] + " " + df_cleaned.iloc[38, 1]
        # # df_cleaned['IRRF_Valor'] = df_cleaned.iloc[49, 1]
        # # df_cleaned['ISS_valor'] = df_cleaned.iloc[51, 1]
        # # df_cleaned['PIS_valor'] = df_cleaned.iloc[53, 1]
        # # df_cleaned['INSS_valor'] = df_cleaned.iloc[55, 1]
        # # df_cleaned['COFINS_valor'] = df_cleaned.iloc[57, 1]
        # # df_cleaned['INSS_ad_Sat_Valor'] = df_cleaned.iloc[59, 1]
        # # df_cleaned['CSLL_valor'] = df_cleaned.iloc[61, 1]
      
        # 🔹 Conteúdo Bruto da Descrição do Serviço
        conteudo_bruto = data["Descricao_servico"].strip()
        print("🔹 Conteúdo Bruto:")
       

        # 🔹 Filtra apenas texto em maiúsculas
        descricao_servico = "\n".join(re.findall(r'[A-ZÀ-Ú\s]+', conteudo_bruto))
        descricao_servico = re.sub(r'\s+', ' ', descricao_servico).strip()  # Remove múltiplos espaços

        # 🔹 Atualiza o dicionário
        data["Descricao_servico"] = descricao_servico

   
        # 🔹 Converte para DataFrame do Pandas
        df_cleaned = pd.DataFrame([data])



        # df_cleaned.drop([0,1,2], axis=1, inplace=True)
        df_cleaned.drop_duplicates(inplace=True)

        return df_cleaned 
    
    def save_csv(self,df,file_name):
        if not os.path.exists(self.csv_path):
            os.makedirs(self.csv_path, exist_ok=True)
        path = os.path.join(self.csv_path, f"{file_name}.csv")
        df.to_csv(path, sep=";", index=False)

   
    def sanitize_colun_names(self,df):
        df.columns = df.columns.str.replace(" ", "_")
        df.columns = df.columns.str.replace(r'\W','', regex = True)

def ler_pdf_sem_salvar(drive_id, file_id, nome_arquivo):
    """ Lê o conteúdo do PDF diretamente da memória (sem salvar) """
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
   
    response = requests.get(url, headers=criar_contextoGraph(), stream=True)
    response.raise_for_status()  

    pdf_bytes = io.BytesIO(response.content)  # 🔹 Lê direto para a memória
    return pdf_bytes
# completo: 6202558656
# incompleto: 6202558469


if __name__ == "__main__":
    logging.info("📌 Iniciando a extração do SharePoint...")
    arquivo = "progen"
    
    logging.info("🔍 Buscando arquivos no SharePoint...")
    files = arquivosFim()

    df_final = pd.DataFrame()  # 🔹 DataFrame que irá armazenar todos os dados

    logging.info("📂 Iniciando a leitura dos arquivos...")
    
    for file in files:
        drive_id = file["drive_id"]
        file_id = file["item_id"]
        nome_arquivo = file["nome_do_item"]

        pdf_bytes = ler_pdf_sem_salvar(drive_id, file_id, nome_arquivo)

        if pdf_bytes:
            
            
            extractor = PDFExtract(pdf_bytes, configs=rules_dict[arquivo])  # 🔹 Passa o BytesIO
            df = extractor.start()

            if not df.empty:
                df["Nome_Arquivo"] = nome_arquivo  
                df["Data_Entrada"] = pd.Timestamp.now()  
                
                df_final = pd.concat([df_final, df], ignore_index=True)  

    # 🔹 Salva o DataFrame consolidado em CSV
    if not df_final.empty:
        processadata(df_final)
    else:
        print("⚠ Nenhum dado foi extraído. Verifique os arquivos PDF.")