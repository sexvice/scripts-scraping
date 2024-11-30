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


def salvar_dados_excel(
    nome, telefone1, nome_arquivo_excel="dados_coletados_goiania03.xlsx"
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
page = 1
contador = 1
while True:
    # URL da página que você deseja acessar
    url = f"https://br.skokka.com/acompanhantes/rio-de-janeiro/?p={page}"
    print(f"Acessando {url}")
    driver.get(url)

    if page == 1:
        print(f"INICIANDO EM 10s")
        time.sleep(10)
    # Aguarda os elementos carregarem
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, "item-card"))
    )

    # Coleta de dados
    cards = driver.find_elements(By.CLASS_NAME, "item-card")

    nome_arquivo_excel = "dados_coletados_skokka_rj.xlsx"
    nome = "Undefined"
    telefone = None

    print(f"Total de leads: {YELLOW}[{len(cards)}]{RESET} pag {page}")

    for card in cards:
        # Acessa o link do card
        link_element = card.find_element(
            By.CSS_SELECTOR, "p.listing-title.item-title a"
        )

        link = link_element.get_attribute("href")

        # Abre o link em uma nova aba
        driver.execute_script("window.open(arguments[0]);", link)
        driver.switch_to.window(driver.window_handles[1])

        # Aguarda a página carregar e obtém o título
        WebDriverWait(driver, 5).until(lambda d: d.title != "")
        telefone = driver.title.split(" ")[0]

        # Fecha a aba atual e volta para a aba original
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

        # Salva os dados coletados
        nome = link_element.text[:20]

        if telefone:
            status = f"{GREEN}COLETADO{RESET}"
        else:
            status = f"{RED}NÃO COLETADO{RESET}"

        print(f"Lead [{YELLOW}{contador}{RESET}] : {status} - {telefone} > {nome}")
        salvar_dados_excel(nome, telefone, nome_arquivo_excel)
        contador += 1
        time.sleep(1)
    page += 1
# Fecha o driver
driver.quit()
