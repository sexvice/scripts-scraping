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

# Definição das cores
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
RESET = "\033[0m"


def salvar_dados_excel(nome, telefone1, nome_arquivo_excel="dados_coletados_sp.xlsx"):
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
            "Telefone1": [telefone1],
            "Site": "Skokka",
        }
    )

    # Concatena os dados existentes com o novo dado
    df = pd.concat([df, novo_dado], ignore_index=True)

    # Salva o DataFrame no arquivo Excel
    df.to_excel(nome_arquivo_excel, index=False)


# Configurações do ChromeDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Executa o Chrome em modo headless
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Inicializa o WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

url = f"https://www.photoacompanhantes.com/acompanhantes/sao-paulo"
driver.get(url)


try:

    last_height = driver.execute_script("return document.body.scrollHeight")
    printed_names = set()  # Conjunto para armazenar nomes já impressos
    id_counter = 0  # Contador para IDs sequenciais
    time.sleep(500)
    while True:
        try:
            # Localiza a seção específica
            section = driver.find_element(By.ID, "resultados")
            print(section)

            break
        except:
            print("Erro ao rolar a página esperando 10s")
            time.sleep(10)
            driver.execute_script("window.open(arguments[0]);", url)

            driver.close()
            driver.switch_to.window(driver.window_handles[0])

finally:
    # Fecha o navegador
    driver.quit()
