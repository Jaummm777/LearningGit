# tiny_login.py

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Função para fazer login no Tiny ERP
def login_tiny(driver, email, senha):
    driver.get("https://erp.tiny.com.br/")
    time.sleep(2)
    driver.find_element(By.XPATH, "//*[@id='kc-content-wrapper']/react-login/div/div/div[1]/div[1]/div[1]/form/div[1]/div/input").send_keys(email)
    driver.find_element(By.XPATH, "//*[@id='kc-content-wrapper']/react-login/div/div/div[1]/div[1]/div[1]/form/div[2]/div/input").send_keys(senha)
    driver.find_element(By.XPATH, "//*[@id='kc-content-wrapper']/react-login/div/div/div[1]/div[1]/div[1]/form/div[3]/button").click()
    time.sleep(5)

    # Verificar se há um aviso de usuário já logado e clicar no botão para confirmar o logout do outro usuário
    try:
        login_button = driver.find_element(By.XPATH, "//*[@id='bs-modal-ui-popup']/div/div/div/div[3]/button[1]")
        if login_button.is_displayed():
            login_button.click()
            time.sleep(5)
    except:
        pass

# Função para fazer logout no Tiny ERP
def logout_tiny(driver):
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='main-menu']/div[2]/div[1]/div[1]/div[2]/ul/li[4]/a/div/div[2]/img"))).click()
        time.sleep(2)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='main-menu']/div[2]/div[2]/nav[7]/div[3]/a"))).click()
        time.sleep(5)
    except:
        # Verificar se a mensagem de sessão expirada aparece
        try:
            alert_message = driver.find_element(By.XPATH, "//*[contains(text(), 'Atenção! Sua sessão expirou ou este usuário efetuou login em outra máquina.')]")
            if alert_message.is_displayed():
                # Abrir uma nova aba e fazer login na conta secundária
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[1])
                login_tiny(driver, "LAVRATTIASSESSORIA@GMAIL.COM", "Transformar.vidas07")
        except:
            pass