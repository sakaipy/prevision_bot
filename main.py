from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from dotenv import load_dotenv
import os
import time

from utils.Prevision_Login import realizar_login
from utils.Prevision_Navigation import navegar_para_medicao

# Carrega variáveis do .env
load_dotenv()
USER = os.getenv("PREVISION_USER")
PASS = os.getenv("PREVISION_PASS")

# Configurações do Chrome (visível)
chrome_options = Options()
chrome_options.add_argument("--start-maximized")

# Inicializa o driver
service = Service()
driver = webdriver.Chrome(service=service, options=chrome_options)

# Acessa o site da Prevision
url = "https://app.prevision.com.br/app/home"
driver.get(url)
time.sleep(5)

# Função para verificar se está logado
def esta_logado(driver):
    try:
        driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/main/div/div[1]/div/div[3]/div/div/header/div/div/div[2]/button[2]/span/i")  # exemplo de seletor
        return True
    except NoSuchElementException:
        return False

# Verifica login
if esta_logado(driver):
    print("✅ Usuário já está logado. Prosseguindo...")
else:
    print("🔐 Não logado. Realizando login...")
    realizar_login(driver, USER, PASS)

navegar_para_medicao(driver)
