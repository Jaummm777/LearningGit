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

# Configurações do navegador
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Iniciar navegador maximizado
prefs = {"download.default_directory": "C:\\Users\\AGREGAR\\Downloads"}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)

try:
    # Acessar Tiny ERP
    driver.get("https://erp.tiny.com.br/")
    time.sleep(2)

    # Login
    driver.find_element(By.XPATH, "//*[@id=\"kc-content-wrapper\"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input").send_keys("agregarnegocios@gmail.com")
    driver.find_element(By.XPATH, "//*[@id=\"kc-content-wrapper\"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input").send_keys("Transformar.vidas07")
    driver.find_element(By.XPATH, "//*[@id=\"kc-content-wrapper\"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button").click()
    time.sleep(5)

    # Expandir menu principal
    driver.find_element(By.XPATH, "//*[@id=\"main-menu\"]/div[2]/div[1]/div[1]/div[1]").click()  # Ajuste o XPath conforme necessário
    time.sleep(2)

    driver.find_element(By.XPATH, "//*[@id=\"main-menu\"]/div[2]/div[1]/div[1]/nav/ul/li[5]/a/span").click()
    time.sleep(2)

    # Acessar Caixa
    driver.find_element(By.XPATH, "//*[@id=\"main-menu\"]/div[2]/div[2]/nav[5]/ul/li[1]/a/span").click()
    time.sleep(2)

    # Selecionar Período "Intervalo"
    driver.find_element(By.XPATH, "//*[@id=\"page-wrapper\"]/div[4]/div[1]/div[3]/ul/li[1]/a").click()  # Ajuste o XPath conforme necessário
    time.sleep(1)
    driver.find_element(By.XPATH, "//*[@id=\"opc-periodo\"]").click()  # Selecionar "Intervalo"
    time.sleep(2)

    # Apagar período 
    data_inicio = driver.find_element(By.XPATH, "//*[@id=\"data-ini\"]")
    data_inicio.send_keys(Keys.CONTROL, 'a', Keys.BACKSPACE)
    time.sleep(2)

    # Selecionar Período
    driver.find_element(By.XPATH, "//*[@id=\"data-ini\"]").send_keys("01/01/2024")
    time.sleep(2)

    # Clicar no Botão "Aplicar"
    driver.find_element(By.XPATH, "//*[@id=\"page-wrapper\"]/div[4]/div[1]/div[3]/ul/li[1]/div/div[6]/button[1]").click()
    time.sleep(2)

    # Clicar nas contas
    driver.find_element(By.XPATH, "//*[@id=\"page-wrapper\"]/div[4]/div[2]/div[1]/ul/li/a/span").click()
    time.sleep(2)

    # Remover Filtro de Conta Bancária
    driver.find_element(By.XPATH, "//*[@id=\"item-conta-todas\"]").click()
    time.sleep(7)

    # Clicar em mais ações
    driver.find_element(By.XPATH, "//*[@id=\"page-wrapper\"]/div[4]/div[1]/div[1]/div/div[2]/button/span[1]").click()
    time.sleep(5)

    # Clicar em exportar relatórios
    driver.find_element(By.XPATH, "//*[@id=\"page-wrapper\"]/div[4]/div[1]/div[1]/div/div[2]/ul/li[4]/a").click()
    time.sleep(5)

    # Exportar a planilha
    driver.find_element(By.XPATH, "//*[@id=\"bs-modal\"]/div/div/div/div[3]/button[1]").click()
    time.sleep(10)  # Aguardar o início do download

    # Caminhos dos arquivos
    pasta_downloads = "C:\\Users\\AGREGAR\\Downloads"
    pasta_cliente = "G:\\Drives compartilhados\\Agregar Negócios - Drive Geral\\Agregar Clientes Ativos\\FRAN MAKES\\3. Finanças\\3 - Relatórios Financeiros\\BASE DE DADOS - TINY\\TESTE AUTOMAÇÃO"

    # Loop para verificar se o arquivo foi baixado
    arquivo_baixado = None
    for _ in range(60):  # Tentar por 60 segundos
        lista_arquivos = glob.glob(os.path.join(pasta_downloads, "*.xlsx"))
        if lista_arquivos:
            arquivo_baixado = max(lista_arquivos, key=os.path.getctime)
            if "caixa_" in os.path.basename(arquivo_baixado):
                break
        time.sleep(1)  # Esperar 1 segundo antes de tentar novamente

    if arquivo_baixado and "caixa_" in os.path.basename(arquivo_baixado):
        # Caminho do arquivo de destino
        caminho_arquivo_cliente = os.path.join(pasta_cliente, os.path.basename(arquivo_baixado))

        # Mover e substituir o arquivo
        if os.path.exists(arquivo_baixado):
            if os.path.exists(caminho_arquivo_cliente):
                os.remove(caminho_arquivo_cliente)  # Remover o arquivo antigo
            shutil.move(arquivo_baixado, caminho_arquivo_cliente)  # Mover o novo arquivo
            print(f"Arquivo movido para {caminho_arquivo_cliente}")
        else:
            print(f"Arquivo {arquivo_baixado} não encontrado na pasta de downloads.")
    else:
        print("Nenhum arquivo foi baixado.")

except Exception as e:
    print(f"Ocorreu um erro: {e}")

input("Pressione Enter para fechar o navegador...")
driver.quit()