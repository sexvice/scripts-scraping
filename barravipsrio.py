from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from openpyxl import Workbook
import pandas as pd


def salvar_dados_excel(
    nome, telefone, nome_arquivo_excel="dados_coletados_rj.xlsx"
):
    # Verifica se o arquivo Excel já existe
    if os.path.exists(nome_arquivo_excel):
        # Se existir, carrega os dados existentes
        df = pd.read_excel(nome_arquivo_excel)
    else:
        # Se não existir, cria um novo DataFrame com as colunas
        df = pd.DataFrame(columns=["Nome", "Telefone1", "Site"])

    # Cria um DataFrame com o novo dado
    novo_dado = pd.DataFrame(
        {
            "Nome": [nome],
            "Telefone": [telefone],
            "Site": "Barravipsrio",
        }
    )

    # Concatena os dados existentes com o novo dado
    df = pd.concat([df, novo_dado], ignore_index=True)

    # Salva o DataFrame no arquivo Excel
    df.to_excel(nome_arquivo_excel, index=False)

# Definição das cores
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
RESET = "\033[0m"

# Configurações do ChromeDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Executa o Chrome em modo headless
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Inicializa o WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# URL da página que você deseja acessar
#https://barravipsrio.com/
#https://barravipsrio.com/index.php?f=acompanhantes-sao-paulo-sp
url = "https://barravipsrio.com/"
driver.get(url)
print("Esperando 30s")
time.sleep(30)
print("Iniciando...")
try:
    section = driver.find_element(By.CLASS_NAME, "items-home")
    cards = section.find_elements(By.CLASS_NAME, "home_info")
    if cards:
        for card in cards:
            # Coleta o nome do card
            name_element = card.find_element(By.CSS_SELECTOR, "div.bottom_info h2 a")
            nome = name_element.text
            
            # Coleta o link do perfil
            profile_link = card.find_element(By.CSS_SELECTOR, "a").get_attribute("href")
            
            # Abre o link do card em uma nova aba
            driver.execute_script("window.open(arguments[0]);", profile_link)
            driver.switch_to.window(driver.window_handles[1])
            
            # Aguarda até que o elemento do telefone esteja disponível
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "a.btn-wpp-default.btnwhats"))
            )
            
            # Coleta o telefone do link do WhatsApp
            whatsapp_link = driver.find_element(By.CSS_SELECTOR, "a.btn-wpp-default.btnwhats").get_attribute("href")
            telefone = whatsapp_link.split("phone=")[1].split("&")[0]
            
            # Printa o nome e o telefone
            print(f"{nome}: {telefone}")
            salvar_dados_excel(nome, telefone)
            # Fecha a aba e retorna para a aba original
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            
    
except:
    print("Erro")