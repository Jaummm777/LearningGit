# Bibliotecas
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime

# Configuração do serviço do ChromeDriver
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

# Função para realizar o login
def realizar_login(navegador, usuario, senha):
    print("Acessando a página de login...")
    navegador.get("https://erp.tiny.com.br")
    
    try:
        # Esperar até que os campos de login estejam presentes
        print("Esperando pelo campo de e-mail...")
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input'))
        )
        
        # Preencher o campo de e-mail
        print("Preenchendo o campo de e-mail...")
        navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input').send_keys(usuario)
        
        # Preencher o campo de senha
        print("Preenchendo o campo de senha...")
        navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input').send_keys(senha)
        
        # Clicar no botão de login
        print("Clicando no botão de login...")
        navegador.find_element(By.XPATH, '//*[@id="kc-content-wrapper"]/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button').click()
        
        # Esperar até que a página principal esteja carregada
        print("Esperando a página principal carregar...")
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main-menu"]/div[2]/div[1]/div[1]/div[1]/i'))
        )
        
        print("Login realizado com sucesso!")
    
    except Exception as e:
        print(f"Erro ao realizar login: {e}")

# Função para navegar até o submenu e clicar em "Finanças" e "Caixa"
def navegar_para_caixa(navegador):
    try:
        # Esperar até que o menu principal esteja presente
        print("Esperando pelo menu principal...")
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[2]/div[1]/div[1]/div[1]/i"))
        )
        
        # Clicar no submenu "Finanças"
        print("Clicando no submenu 'Finanças'...")
        menu_financas = navegador.find_element(By.XPATH, "/html/body/div[3]/div[2]/div[1]/div[1]/nav/ul/li[5]/a/span")
        menu_financas.click()

        # Esperar até que o submenu "Caixa" esteja visível e clicável
        print("Esperando pelo submenu 'Caixa' dentro de 'Finanças'...")
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/div[2]/nav[5]/ul/li[1]/a/span"))
        )
        
        # Clicar no submenu "Caixa" dentro de "Finanças"
        print("Clicando no submenu 'Caixa' dentro de 'Finanças'...")
        submenu_caixa = navegador.find_element(By.XPATH, "/html/body/div[3]/div[2]/div[2]/nav[5]/ul/li[1]/a/span")
        submenu_caixa.click()
        
        print("Navegação para 'Caixa' realizada com sucesso!")
    
    except Exception as e:
        print(f"Erro ao navegar para 'Caixa': {e}")

# Função para filtrar a data do mês de janeiro até o mês atual
def filtrar_data(navegador):
    try:
        # Esperar até que os campos de data estejam presentes
        print("Esperando pelos campos de data...")
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@name='data_inicio']"))
        )
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@name='data_fim']"))
        )
        
        # Definir as datas de início e fim
        data_inicio = "01/01/" + str(datetime.now().year)
        data_fim = datetime.now().strftime("%d/%m/%Y")
        
        # Preencher o campo de data de início
        print(f"Preenchendo a data de início: {data_inicio}")
        campo_data_inicio = navegador.find_element(By.XPATH, "//input[@name='data_inicio']")
        campo_data_inicio.clear()
        campo_data_inicio.send_keys(data_inicio)
        
        # Preencher o campo de data de fim
        print(f"Preenchendo a data de fim: {data_fim}")
        campo_data_fim = navegador.find_element(By.XPATH, "//input[@name='data_fim']")
        campo_data_fim.clear()
        campo_data_fim.send_keys(data_fim)
        
        # Clicar no botão de filtrar
        print("Clicando no botão de filtrar...")
        botao_filtrar = navegador.find_element(By.XPATH, "//button[contains(text(), 'Filtrar')]")
        botao_filtrar.click()
    
    except Exception as e:
        print(f"Erro ao filtrar a data: {e}")

# Credenciais de login
usuario = "private"
senha = "private"

# Realizar login
realizar_login(navegador, usuario, senha)

# Navegar para "Caixa"
navegar_para_caixa(navegador)

# Filtrar a data do mês de janeiro até o mês atual
filtrar_data(navegador)

# Manter o navegador aberto
input("Pressione Enter para fechar o navegador...")
navegador.quit()
