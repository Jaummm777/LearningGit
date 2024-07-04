# Bibliotecas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import shutil
import glob
import pandas as pd
import pyxlsb
import win32com.client as win32
from openpyxl import load_workbook

# Função para verificar e converter o arquivo XLS para XLSX usando pyxlsb
def verificar_e_converter_xls(caminho_arquivo, novo_caminho):
    try:
        # Tentativa com pyxlsb
        with open(caminho_arquivo, 'rb') as f:
            workbook = pyxlsb.open_workbook(f)
            sheet = workbook.get_sheet(workbook.sheet_names()[0])
            data = []
            for row in sheet.rows():
                data.append([item.v for item in row])
        
        # Criar DataFrame e salvar como .xlsx
        df = pd.DataFrame(data[1:], columns=data[0])  # Presume que a primeira linha contém os cabeçalhos
        df.to_excel(novo_caminho, index=False, engine='openpyxl')
        print(f"Arquivo convertido e salvo como {novo_caminho} usando pyxlsb")
        return True
    except Exception as e:
        print(f"Erro com pyxlsb: {e}")
        try:
            # Tentativa com win32com
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(caminho_arquivo)
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)  # Excluir o arquivo existente
            workbook.SaveAs(novo_caminho, FileFormat=51)  # 51 é o código para .xlsx
            workbook.Close(False)
            excel.Application.Quit()
            print(f"Arquivo convertido e salvo como {novo_caminho} usando win32com")
            return True
        except Exception as e2:
            print(f"Erro com win32com: {e2}")
            return False

# Função para renomear a planilha dentro do arquivo Excel
def renomear_planilha(caminho_arquivo, novo_nome_planilha):
    try:
        workbook = load_workbook(caminho_arquivo)
        sheet = workbook.active
        sheet.title = novo_nome_planilha
        workbook.save(caminho_arquivo)
        print(f"Planilha renomeada para {novo_nome_planilha} no arquivo {caminho_arquivo}")
    except Exception as e:
        print(f"Erro ao renomear a planilha: {e}")

# Função para fazer login no Tiny ERP
def login_tiny(driver, email, senha):
    driver.get("https://erp.tiny.com.br/")
    time.sleep(2)
    driver.find_element(By.XPATH, "//*[@id='kc-content-wrapper']/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input").send_keys(email)
    driver.find_element(By.XPATH, "//*[@id='kc-content-wrapper']/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input").send_keys(senha)
    driver.find_element(By.XPATH, "//*[@id='kc-content-wrapper']/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button").click()
    time.sleep(5)

# Função para fazer logout no Tiny ERP
def logout_tiny(driver):
    driver.find_element(By.XPATH, "//*[@id=\"main-menu\"]/div[2]/div[1]/div[1]/div[1]").click()
    time.sleep(2)
    driver.find_element(By.XPATH, "//*[@id=\"main-menu\"]/div[2]/div[1]/div[1]/div[2]/ul/li[4]/a/div/div[2]/img").click()
    time.sleep(2)
    driver.find_element(By.XPATH, "//*[@id=\"main-menu\"]/div[2]/div[2]/nav[7]/div[3]/a")
    time.sleep(5)

# Configurações do navegador
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"download.default_directory": "C:\\Users\\AGREGAR\\Downloads"}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)

try:
    # Login na primeira conta
    login_tiny(driver, "agregarnegocios@gmail.com", "Transformar.vidas07")

    # Verificar se há um aviso de usuário já logado e clicar no botão para confirmar o logout do outro usuário
    try:
        login_button = driver.find_element(By.XPATH, "//*[@id='bs-modal-ui-popup']/div/div/div/div[3]/button[1]")
        if login_button.is_displayed():
            login_button.click()
            time.sleep(5)
    except:
        pass

    # Expandir menu principal
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/div[1]").click() 
    time.sleep(2)

    # Acessar Caixa
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/nav/ul/li[5]/a/span").click()
    time.sleep(2)

    # Acessar Caixa
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[2]/nav[5]/ul/li[1]/a/span").click()
    time.sleep(2)

    # Selecionar Período "Intervalo"
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[1]/div[3]/ul/li[1]/a").click() 
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id='opc-periodo']").click()  # Selecionar "Intervalo"
    time.sleep(2)

    # Apagar período 
    data_inicio = driver.find_element(By.XPATH, "//*[@id='data-ini']")
    data_inicio.send_keys(Keys.CONTROL, 'a', Keys.BACKSPACE)
    time.sleep(2)

    # Selecionar Período
    driver.find_element(By.XPATH, "//*[@id='data-ini']").send_keys("01/01/2024")
    time.sleep(2)

    # Clicar no Botão "Aplicar"
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[1]/div[3]/ul/li[1]/div/div[6]/button[1]").click()
    time.sleep(2)

    # Clicar nas contas
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[2]/div[1]/ul/li/a/span").click()
    time.sleep(2)

    # Remover Filtro de Conta Bancária
    driver.find_element(By.XPATH, "//*[@id='item-conta-todas']").click()
    time.sleep(7)

    # Clicar em mais ações
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[1]/div[1]/div/div[2]/button/span[1]").click()
    time.sleep(5)

    # Clicar em exportar relatórios
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[1]/div[1]/div/div[2]/ul/li[4]/a").click()
    time.sleep(5)

    # Exportar a planilha
    driver.find_element(By.XPATH, "//*[@id='bs-modal']/div/div/div/div[3]/button[1]").click()
    time.sleep(10)  # Aguardar o início do download

    # Caminhos dos arquivos
    pasta_downloads = "C:\\Users\\AGREGAR\\Downloads"
    pasta_cliente = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\FRAN MAKES\\3. Finanças\\3 - Relatórios Financeiros\\BASE DE DADOS - TINY\\CAIXA"

    # Verificar se o arquivo foi baixado
    tempo_espera = 60  # Tempo máximo de espera em segundos
    tempo_inicial = time.time()
    arquivo_baixado = None

    while (time.time() - tempo_inicial) < tempo_espera:
        arquivos = glob.glob(os.path.join(pasta_downloads, "*.xls"))
        if arquivos:
            arquivo_baixado = max(arquivos, key = os.path.getctime)
            break
        time.sleep(1)  # Esperar 1 segundo antes de tentar novamente

    if arquivo_baixado:
        # Caminho para salvar o novo arquivo .xlsx
        caminho_arquivo_xlsx = os.path.splitext(arquivo_baixado)[0] + ".xlsx"

        # Verificar e converter o arquivo .xls para .xlsx
        if verificar_e_converter_xls(arquivo_baixado, caminho_arquivo_xlsx):
            # Renomear a planilha dentro do arquivo Excel
            renomear_planilha(caminho_arquivo_xlsx, "caixa fran")

            # Renomear o arquivo
            novo_nome = f"caixa fran.xlsx"
            novo_caminho = os.path.join(pasta_downloads, novo_nome)

            # Renomear o arquivo
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)  # Excluir o arquivo existente
            os.rename(caminho_arquivo_xlsx, novo_caminho)

            # Caminho do arquivo de destino
            caminho_arquivo_cliente = os.path.join(pasta_cliente, novo_nome)

            # Mover e substituir o arquivo
            shutil.move(novo_caminho, caminho_arquivo_cliente, copy_function=shutil.copy2)  # Mover o novo arquivo, substituindo se necessário
            print(f"Arquivo movido para {caminho_arquivo_cliente}")
    else:
        print("Nenhum arquivo foi baixado.")
        
    # Acessar Tiny ERP
    driver.get("https://erp.tiny.com.br/")
    time.sleep(2)

    # Expandir menu principal
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/div[1]").click()
    time.sleep(2)

    # Acessar Vendas
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/nav/ul/li[4]/a/span").click()
    time.sleep(2)

    # Acessar Relatórios de Vendas
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[2]/nav[4]/ul/li[14]/a/span").click()
    time.sleep(2)

    # Selecionar Relatório de Vendas
    driver.find_element(By.XPATH, "//*[@id='resultado']/div[1]/div/a[1]/span[1]").click()
    time.sleep(2)

    # Agrupar por Ecommerce e Subagrupar por Produtos
    driver.find_element(By.XPATH, "//*[@id='agrupamento']").click()
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id='agrupamento']/option[6]").click()  # Selecionar Ecommerce
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id='subAgrupamento']").click()
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id='subAgrupamento']/option[3]").click()  # Selecionar Produtos
    time.sleep(1)

    # Opções avançadas
    driver.find_element(By.XPATH, "//*[@id='link-pesquisa']").click()
    time.sleep(2)
    driver.find_element(By.XPATH, "//*[@id='tipoLucro']/option[2]").click()

    # Gerar Relatório
    driver.find_element(By.XPATH, "//*[@id='btn-visualizar']").click()
    time.sleep(10)  # Aguardar o início do download

    # Download dos relatórios
    driver.find_element(By.XPATH, "//*[@id='btn-download']").click()
    time.sleep(15)  # Aguardar o download ser concluído

    # Caminhos dos arquivos
    pasta_downloads = "C:\\Users\\AGREGAR\\Downloads"
    pasta_cliente = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\FRAN MAKES\\3. Finanças\\3 - Relatórios Financeiros\\ECOMMERCE"

    # Verificar se o arquivo foi baixado
    tempo_espera = 60
    tempo_inicial = time.time()
    arquivo_baixado = None

    while (time.time() - tempo_inicial) < tempo_espera:
        arquivos = glob.glob(os.path.join(pasta_downloads, "*.xls"))
        if arquivos:
            arquivo_baixado = max(arquivos, key=os.path.getctime)
            break
        time.sleep(1)

    if arquivo_baixado:
        # Caminho para salvar o novo arquivo .xlsx
        caminho_arquivo_xlsx = os.path.splitext(arquivo_baixado)[0] + ".xlsx"

        # Verificar e converter o arquivo .xls para .xlsx
        if verificar_e_converter_xls(arquivo_baixado, caminho_arquivo_xlsx):
            # Renomear o arquivo com base no mês atual
            mes_atual = time.strftime("%m")
            novo_nome = f"ECOMMERCE {mes_atual}.xlsx"
            novo_caminho = os.path.join(pasta_downloads, novo_nome)

            # Renomear o arquivo
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)  # Excluir o arquivo existente
            os.rename(caminho_arquivo_xlsx, novo_caminho)

            # Caminho do arquivo de destino
            caminho_arquivo_cliente = os.path.join(pasta_cliente, novo_nome)

            # Mover e substituir o arquivo
            shutil.move(novo_caminho, caminho_arquivo_cliente, copy_function=shutil.copy2)  # Mover o novo arquivo, substituindo se necessário
            print(f"Arquivo movido para {caminho_arquivo_cliente}")
    else:
        print("Nenhum arquivo foi baixado.")

    driver.get("https://erp.tiny.com.br/")
    time.sleep(2)

    # Expandir menu principal
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/div[1]").click()
    time.sleep(2)

    # Acessar Suprimentos
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/nav/ul/li[3]/a/span").click()
    time.sleep(2)

    # Acessar Relatórios de Suprimentos
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[2]/nav[3]/ul/li[6]/a/span").click()
    time.sleep(2)

    # Selecionar Relatório de Estoque Multiempresa
    driver.find_element(By.XPATH, "//*[@id='resultado']/div[1]/div/a/span[1]").click()
    time.sleep(2)

    # Gerar Relatório
    driver.find_element(By.XPATH, "//*[@id='btn-visualizar']").click()
    time.sleep(45)  # Aguardar o início do download

    # Download dos relatórios
    driver.find_element(By.XPATH, "//*[@id='btn-download']").click()
    time.sleep(90)  # Aguardar o download ser concluído

    # Caminhos dos arquivos
    pasta_downloads = "C:\\Users\\AGREGAR\\Downloads"
    pasta_cliente = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\FRAN MAKES\\3. Finanças\\3 - Relatórios Financeiros\\BASE DE DADOS - TINY\\Estoque multi empresa"

    # Verificar se o arquivo foi baixado
    tempo_espera = 60
    tempo_inicial = time.time()
    arquivo_baixado = None

    while (time.time() - tempo_inicial) < tempo_espera:
        arquivos = glob.glob(os.path.join(pasta_downloads, "*.xls"))
        if arquivos:
            arquivo_baixado = max(arquivos, key=os.path.getctime)
            break
        time.sleep(1)

    if arquivo_baixado:
        # Caminho para salvar o novo arquivo .xlsx
        caminho_arquivo_xlsx = os.path.splitext(arquivo_baixado)[0] + ".xlsx"

        # Verificar e converter o arquivo .xls para .xlsx
        if verificar_e_converter_xls(arquivo_baixado, caminho_arquivo_xlsx):
            # Renomear a planilha dentro do arquivo Excel
            renomear_planilha(caminho_arquivo_xlsx, "Estoque Multiempresa fef")

            # Renomear o arquivo
            novo_nome = f"estoque_multi_empresa fran e fpml.xlsx"
            novo_caminho = os.path.join(pasta_downloads, novo_nome)

            # Renomear o arquivo
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)  # Excluir o arquivo existente
            os.rename(caminho_arquivo_xlsx, novo_caminho)

            # Caminho do arquivo de destino
            caminho_arquivo_cliente = os.path.join(pasta_cliente, novo_nome)

            # Mover e substituir o arquivo
            shutil.move(novo_caminho, caminho_arquivo_cliente, copy_function=shutil.copy2)  # Mover o novo arquivo, substituindo se necessário
            print(f"Arquivo movido para {caminho_arquivo_cliente}")
    else:
        print("Nenhum arquivo foi baixado.")

    # Logout da primeira conta
    logout_tiny(driver)

    # Login na segunda conta
    login_tiny(driver, "LAVRATTIASSESSORIA@GMAIL.COM", "Transformar.vidas07")

    # Verificar se há um aviso de usuário já logado e clicar no botão para confirmar o logout do outro usuário
    try:
        login_button = driver.find_element(By.XPATH, "//*[@id='bs-modal-ui-popup']/div/div/div/div[3]/button[1]")
        if login_button.is_displayed():
            login_button.click()
            time.sleep(5)
    except:
        pass

    # Expandir menu principal
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/div[1]").click() 
    time.sleep(2)

    # Acessar Caixa
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/nav/ul/li[5]/a/span").click()
    time.sleep(2)

    # Acessar Caixa
    driver.find_element(By.XPATH, "//*[@id='main-menu']/div[2]/div[2]/nav[5]/ul/li[1]/a/span").click()
    time.sleep(2)

    # Selecionar Período "Intervalo"
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[1]/div[3]/ul/li[1]/a").click() 
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id='opc-periodo']").click()  # Selecionar "Intervalo"
    time.sleep(2)

    # Apagar período 
    data_inicio = driver.find_element(By.XPATH, "//*[@id='data-ini']")
    data_inicio.send_keys(Keys.CONTROL, 'a', Keys.BACKSPACE)
    time.sleep(2)

    # Selecionar Período
    driver.find_element(By.XPATH, "//*[@id='data-ini']").send_keys("01/01/2024")
    time.sleep(2)

    # Clicar no Botão "Aplicar"
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[1]/div[3]/ul/li[1]/div/div[6]/button[1]").click()
    time.sleep(2)

    # Clicar nas contas
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[2]/div[1]/ul/li/a/span").click()
    time.sleep(2)

    # Remover Filtro de Conta Bancária
    driver.find_element(By.XPATH, "//*[@id='item-conta-todas']").click()
    time.sleep(7)

    # Clicar em mais ações
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[1]/div[1]/div/div[2]/button/span[1]").click()
    time.sleep(5)

    # Clicar em exportar relatórios
    driver.find_element(By.XPATH, "//*[@id='page-wrapper']/div[4]/div[1]/div[1]/div/div[2]/ul/li[4]/a").click()
    time.sleep(5)

    # Exportar a planilha
    driver.find_element(By.XPATH, "//*[@id='bs-modal']/div/div/div/div[3]/button[1]").click()
    time.sleep(10)  # Aguardar o início do download

    # Caminhos dos arquivos
    pasta_downloads = "C:\\Users\\AGREGAR\\Downloads"
    pasta_cliente_fpml = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\FPML\\3. Finanças\\3 - Relatórios Financeiros\\BASE DE DADOS - TINY\\CAIXA"

    # Verificar se o arquivo foi baixado
    tempo_espera = 60  # Tempo máximo de espera em segundos
    tempo_inicial = time.time()
    arquivo_baixado = None

    while (time.time() - tempo_inicial) < tempo_espera:
        arquivos = glob.glob(os.path.join(pasta_downloads, "*.xls"))
        if arquivos:
            arquivo_baixado = max(arquivos, key = os.path.getctime)
            break
        time.sleep(1)  # Esperar 1 segundo antes de tentar novamente

    if arquivo_baixado:
        # Caminho para salvar o novo arquivo .xlsx
        caminho_arquivo_xlsx = os.path.splitext(arquivo_baixado)[0] + ".xlsx"

        # Verificar e converter o arquivo .xls para .xlsx
        if verificar_e_converter_xls(arquivo_baixado, caminho_arquivo_xlsx):
            # Renomear a planilha dentro do arquivo Excel
            renomear_planilha(caminho_arquivo_xlsx, "caixa fpml")

            # Renomear o arquivo
            novo_nome = f"caixa fpml.xlsx"
            novo_caminho = os.path.join(pasta_downloads, novo_nome)

            # Renomear o arquivo
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)  # Excluir o arquivo existente
            os.rename(caminho_arquivo_xlsx, novo_caminho)

            # Caminho do arquivo de destino
            caminho_arquivo_cliente_fpml = os.path.join(pasta_cliente_fpml, novo_nome)

            # Mover e substituir o arquivo
            shutil.move(novo_caminho, caminho_arquivo_cliente_fpml, copy_function=shutil.copy2)  # Mover o novo arquivo, substituindo se necessário
            print(f"Arquivo movido para {caminho_arquivo_cliente_fpml}")
    else:
        print("Nenhum arquivo foi baixado.")

except Exception as e:
    print(f"Ocorreu um erro: {e}")
finally:
    driver.quit()

input("Pressione Enter para fechar o navegador...")