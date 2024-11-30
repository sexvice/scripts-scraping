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

url = input("Digite o link: ")
local = input("Digite o local: ")


def salvar_dados_excel(
    nome, telefone, nome_arquivo_excel
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
            "Site": "Private55",
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
#https://br.simpleescorts.com/acompanhantes/goiania/?gender=mulher&p=
#https://br.simpleescorts.com/acompanhantes/rio-de-janeiro/?gender=mulher&p=
#https://br.simpleescorts.com/acompanhantes/sao-paulo/?gender=mulher&p=

page = 1
contador = 1
while True:
    try:
        # URL da página que você deseja acessar
        url_acess = url + str(page)
        print(f"Acessando {url_acess}")
        driver.get(url_acess)
        if page == 1:
            print(f"INICIANDO EM 10s")
            time.sleep(10)
            
        # Aguarda todos os elementos com o atributo data-test="ad" carregarem antes de prosseguir
        elementos_ad = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[data-test="ad"]'))
        )
        
        if isinstance(elementos_ad, list):
            print(f"Total de {len(elementos_ad)} cards")
            for card in elementos_ad:
                try:
                    data_location = card.get_attribute("data-location")
                    if data_location:
                        link_completo = "https://br.simpleescorts.com" + data_location
                        driver.execute_script("window.open('');")
                        driver.switch_to.window(driver.window_handles[1])
                        driver.get(link_completo)
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.text-2xl.font-medium.text-gray-100'))
                        )
                        elemento_nome = driver.find_element(By.CSS_SELECTOR, 'div.text-2xl.font-medium.text-gray-100')
                        nome = elemento_nome.text
                        
                        elemento_telefone = driver.find_element(By.CSS_SELECTOR, 'a[data-test="phone-btn"]')
                        telefone_href = elemento_telefone.get_attribute("href")
                        telefone = telefone_href.split(":")[-1].replace("-", "").replace(" ", "")
                        
                        print(f"[{GREEN}{contador}{RESET}] Nome: {nome}, Telefone: {telefone}")
                        contador += 1
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        nome_arquivo_excel = f"dados_coletados_{local}.xlsx"
                        salvar_dados_excel(nome, telefone, nome_arquivo_excel)
                except:
                    pass
    except:
        pass
    page += 1
    if page > 70:
        print("Total batido de paginas")
        url = "https://br.simpleescorts.com/acompanhantes/goiania/?gender=mulher&p="
        page = 1
        local = "go"
        contador = 1
        print("Trocando de cidade")
        
    