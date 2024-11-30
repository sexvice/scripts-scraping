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
    nome, telefone1, nome_arquivo_excel="dados_coletados_lindas_go.xlsx"
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
    url = f"https://www.lindas.com.br/acompanhantes/goiania-go?page={page}"
    print(f"Acessando {url}")
    driver.get(url)

    if page == 1:
        print(f"INICIANDO EM 10s")
        time.sleep(10)
    # Aguarda o elemento específico carregar
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.row.justify-content-center")
        )
    )

    # Coleta de dados
    divs = driver.find_elements(By.CSS_SELECTOR, "div.row.justify-content-center div a")

    # Exibe os links encontrados no terminal, descartando links repetidos ou indesejados
    links_encontrados = set()
    for div in divs:
        link = div.get_attribute("href")
        if (
            link
            and link not in links_encontrados
            and not link.startswith(
                "https://www.lindas.com.br/acompanhantes/goiania-go"
            )
        ):
            links_encontrados.add(link)
            # Abre o link em uma nova aba
            driver.execute_script("window.open(arguments[0]);", link)
            driver.switch_to.window(driver.window_handles[-1])

            try:
                # Aguarda o elemento específico carregar
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "a.btn.btn-xs.border-radius-20.text-white")
                    )
                )

                # Coleta o número de telefone
                whatsapp_link = driver.find_element(
                    By.CSS_SELECTOR, "a.btn.btn-xs.border-radius-20.text-white"
                ).get_attribute("href")
                telefone = whatsapp_link.split("phone=")[1].split("&")[0]

                # Coleta o nome
                nome = link.split("/")[-1]

                # Printa o nome e telefone no terminal
                print(f"Nome: {nome}, Telefone: {telefone}")
                salvar_dados_excel(nome, telefone)
            except Exception as e:
                ...
            finally:
                # Fecha a aba atual e volta para a aba original
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

    print(f"Total de links encontrados: {len(divs)}")

    nome = "Undefined"
    telefone = None

    # print(f"Total de leads: {YELLOW}[{len(cards)}]{RESET} pag {page}")

    page += 1
# Fecha o driver
driver.quit()
