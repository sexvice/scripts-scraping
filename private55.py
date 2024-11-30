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

url = input("Digite o link:")
local = input("Digite o local: ")

nome_arquivo_excel = f"dados_coletados_{local}.xlsx"
def salvar_dados_excel(
    nome, telefone
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
#https://www.private55.com/acompanhantes-goiania-gyn
#https://www.private55.com/acompanhantes-rj-rio-de-janeiro
#https://www.private55.com/acompanhantes-sp-sao-paulo

driver.get(url)
print("Esperando 30s")
time.sleep(30)
print("Iniciando...")

try:
    sections = driver.find_elements(By.CLASS_NAME, "fw-row")
    
    if isinstance(sections, list):
        print(f"Total de {len(sections)} seções de putas encontradas")
    leads = 0
    for section in sections:
        cards = section.find_elements(By.ID, "column")
        if isinstance(cards, list):
            for card in cards:
                
                try:
                    
                    # Armazena o link do card
                    link_element = card.find_element(By.CSS_SELECTOR, "a")
                    link = link_element.get_attribute("href")
                    
                    # Tenta extrair o nome do link
                    try:
                        nome = link.split("/")[-1].replace("-", " ").title()
                    except:
                        # Se falhar, tenta extrair o nome do elemento <h4>
                        nome_element = card.find_element(By.CSS_SELECTOR, "h4.fw-imagebox-subtitle")
                        nome = nome_element.text.strip()
                        
                    if "Acompanhantes Sp Sao Paulo#" in nome:
                        continue
                    
                    # Acessa o link em uma nova aba
                    driver.execute_script("window.open(arguments[0]);", link)
                    driver.switch_to.window(driver.window_handles[-1])
                    
                    # Espera o conteúdo carregar
                    try:
                        # Espera até que o elemento do WhatsApp esteja presente
                        whatsapp_div = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div.whatsapp-fixo"))
                        )
                        whatsapp_link_element = whatsapp_div.find_element(By.CSS_SELECTOR, "a")
                        whatsapp_link = whatsapp_link_element.get_attribute("href")
                        
                        # Extrai o telefone do link
                        telefone = whatsapp_link.split("phone=")[-1].split("&")[0]
                    except:
                        telefone = None
                    
                    if telefone:
                        # Imprime o nome e o telefone
                        leads+=1
                        print(f"[{GREEN}{leads}{RESET}] Nome: {nome}, Telefone: {telefone}")
                        salvar_dados_excel(nome, telefone)
                        
                    # Fecha a aba atual e volta para a aba anterior
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    
                except Exception as e:
                    pass  
except:
    print("Error")