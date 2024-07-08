import os
import glob
import time
import pandas as pd
from openpyxl import load_workbook
import win32com.client as win32

# Função para encontrar o arquivo mais recente em uma pasta, ignorando arquivos temporários
def encontrar_arquivo_mais_recente(pasta, extensao="*.xls"):
    arquivos = glob.glob(os.path.join(pasta, extensao))
    arquivos_validos = [arquivo for arquivo in arquivos if not os.path.basename(arquivo).startswith('~$')]
    if not arquivos_validos:
        return None
    arquivo_mais_recente = max(arquivos_validos, key=os.path.getctime)
    return arquivo_mais_recente

# Função para converter arquivo .xls para .xlsx
def converter_xls_para_xlsx(caminho_arquivo_xls):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(caminho_arquivo_xls)
    caminho_arquivo_xlsx = caminho_arquivo_xls + "x"
    
    # Verificar se o arquivo .xlsx já existe e removê-lo
    if os.path.exists(caminho_arquivo_xlsx):
        os.remove(caminho_arquivo_xlsx)
    
    wb.SaveAs(caminho_arquivo_xlsx, FileFormat=51)  # 51 é o código para .xlsx
    wb.Close()
    excel.Application.Quit()
    return caminho_arquivo_xlsx

# Função para copiar dados de uma planilha para outra e adicionar colunas extras
def copiar_dados_entre_planilhas(origem, destino, base):
    try:
        # Carregar a planilha de origem
        df_origem = pd.read_excel(origem)

        # Obter o dia 29 do mês e ano atual
        data_atual = time.strftime("29/%m/%Y")

        # Adicionar colunas "DATA" e "EMPRESA" antes dos dados existentes
        df_origem.insert(0, "EMPRESA", base)
        df_origem.insert(0, "DATA", data_atual)

        # Carregar a planilha de destino
        if os.path.exists(destino):
            df_destino = pd.read_excel(destino)
            df_destino = pd.concat([df_destino, df_origem], ignore_index=True)
        else:
            df_destino = df_origem

        # Salvar a planilha de destino
        df_destino.to_excel(destino, index=False)
        print(f"Dados copiados de {origem} para {destino}")
    except Exception as e:
        print(f"Erro ao copiar dados entre planilhas: {e}")

# Caminhos dos arquivos de teste
pasta_downloads = "C:\\Users\\AGREGAR\\Downloads"
caminho_arquivo_destino = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\FRAN MAKES\\3. Finanças\\3 - Relatórios Financeiros\\BASE DE DADOS - TINY\\VENDAS POR PRODUTOS X MARCAS - GRUPO - automacao.xlsx"

# Encontrar o arquivo mais recente na pasta de downloads
caminho_arquivo_origem = encontrar_arquivo_mais_recente(pasta_downloads)

if caminho_arquivo_origem:
    print(f"Arquivo mais recente encontrado: {caminho_arquivo_origem}")
    
    # Converter o arquivo .xls para .xlsx
    caminho_arquivo_origem_xlsx = converter_xls_para_xlsx(caminho_arquivo_origem)
    
    # Copiar dados para a planilha de destino
    base = "FRAN"  # Ou "FPML", dependendo da origem dos dados
    copiar_dados_entre_planilhas(caminho_arquivo_origem_xlsx, caminho_arquivo_destino, base)
else:
    print("Nenhum arquivo .xls válido encontrado na pasta de downloads.")