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
    nome, telefone, nome_arquivo_excel="dados_coletados_go.xlsx"
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
            "Site": "Gemidos",
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
#https://gemidos.tv/brasil-rio-de-janeiro
#https://gemidos.tv/brasil-sao-paulo
#https://gemidos.tv/brasil-goias-goiania
url = "https://gemidos.tv/brasil-goias-goiania"
driver.get(url)

try:
    last_height = driver.execute_script("return document.body.scrollHeight")
    printed_names = set()  # Conjunto para armazenar nomes já impressos
    cont = 1
    time.sleep(10)
    while True:
        while True:
            try:
                # Localiza a seção específica
                section = driver.find_element(By.ID, "block-860")

                # Inicializa um conjunto para armazenar os nomes já processados
                processed_names = set()

                # Encontra todos os elementos <div> dentro da seção
                cards = section.find_elements(By.CLASS_NAME, "listing-pub")
                print(f"Total de cards encontrados: {len(cards)}")

                # Processa cada card
                for card in cards:
                    try:
                        # Rola a tela para o local onde o elemento está
                        driver.execute_script("arguments[0].scrollIntoView(true);", card)
                        time.sleep(1)  # Aguarda um pouco para garantir que o elemento esteja visível
                        
                        # Coleta o nome do card
                        name_element = card.find_element(By.CSS_SELECTOR, "a.listing-link")
                        nome = name_element.get_attribute("title")
                        
                        if nome not in processed_names:
                            # Abre o link do card em uma nova aba
                            profile_link = name_element.get_attribute("href")
                            if not profile_link:
                                print("link nao estava presente")
                                continue
                            
                            driver.execute_script("window.open(arguments[0]);", profile_link)
                            driver.switch_to.window(driver.window_handles[1])
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, "div.pub-info"))
                            )  # Aguarda até que o elemento esteja disponível

                            # Procura e coleta o telefone
                            try:
                                phone_element = driver.find_element(By.CSS_SELECTOR, "li.pub-phone.pub-tags-item span")
                                telefone = phone_element.text
                            except Exception as e:
                                telefone = "Telefone não encontrado"
                            
                            print(f"{GREEN}[{cont}]{RESET} - {nome}: {telefone}")
                            salvar_dados_excel(nome, telefone)
                            # Fecha a aba e retorna para a aba original
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            time.sleep(2)  # Aguarda um pouco antes de continuar
                            cont += 1
                            processed_names.add(nome)
                    except Exception as e:
                        print(f"Erro ao processar")
                        cont += 1
                        # Retorna para a primeira aba
                        driver.switch_to.window(driver.window_handles[0])
                        # Fecha todas as outras abas, se existirem
                        while len(driver.window_handles) > 1:
                            driver.switch_to.window(driver.window_handles[-1])
                            driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    
                driver.switch_to.window(driver.window_handles[0])       
                # Rola a página para baixo
                driver.execute_script("arguments[0].scrollIntoView(true);", section)

                # Aguarda o carregamento da nova seção
                time.sleep(2)

                # Calcula a nova altura da página e compara com a última altura
                new_height = driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height
               
            except Exception as e:
                print(f"Erro ao rolar a página: {e}, esperando 10s")
                time.sleep(10)
                driver.execute_script("window.open(arguments[0]);", url)

                driver.close()
                driver.switch_to.window(driver.window_handles[0])
except Exception as ex:
    print(f"Erro ao acessar a página: {ex}")
    