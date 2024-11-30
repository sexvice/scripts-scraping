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
        df = pd.DataFrame(columns=["Nome", "Telefone", "Site"])

    # Cria um DataFrame com o novo dado
    novo_dado = pd.DataFrame(
        {
            "Nome": [nome],
            "Telefone": [telefone],
            "Site": "Hotters",
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
#https://www.hotters.com.br/acompanhantes-goiania-go?page=
#https://www.hotters.com.br/acompanhantes-rio-de-janeiro-rj?page=
#https://www.hotters.com.br/acompanhantes-sao-paulo-sp?page=

page = 1
contador = 1
sair_loop = False
while sair_loop == False:
    try:
        # URL da página que você deseja acessar
        url_acess = url + str(page)
        print(f"Acessando {url_acess}")
        driver.get(url_acess)
        if page == 1:
            print(f"INICIANDO EM 10s")
            time.sleep(10)
            
        # Aguarda todos os elementos carregarem e procura na página todos os elementos com a classe "group relative"
        try:
            # Verifica se a mensagem de "Nenhum anúncio encontrado" está presente
            mensagens_erro = driver.find_elements(By.CSS_SELECTOR, 'h3.text-neutral-gray')
            for mensagem in mensagens_erro:
                if "Nenhum anúncio encontrado" in mensagem.text:
                    print("Nenhum anúncio encontrado nesta cidade, finalizando as páginas.")
                    sair_loop = True
                    break
            if sair_loop:
                break
            elementos_ad = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'group.relative'))
            )
        except:
            print("Demorou muito tempo, reabrindo a página...")
            driver.get(url_acess)
            elementos_ad = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'group.relative'))
            )
            
        
        
        if isinstance(elementos_ad, list):
            print(f"Total de {len(elementos_ad)} cards")
            for card in elementos_ad:
                try:
                    link_elemento = card.find_element(By.CSS_SELECTOR, 'a.block.w-32.shrink-0')
                    link = link_elemento.get_attribute("href")
                    elemento_nome = card.find_element(By.CSS_SELECTOR, 'h3.truncate.font-medium.leading-none.text-base')
                    nome = elemento_nome.text
                    
                    
                    driver.execute_script("window.open('');")
                    driver.switch_to.window(driver.window_handles[1])
                    driver.get(link)
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'contact'))
                    )
                    elemento_whatsapp = driver.find_element(By.CSS_SELECTOR, 'a[href^="https://api.whatsapp.com/send/?phone="]')
                    link_whatsapp = elemento_whatsapp.get_attribute("href")
                    telefone_whatsapp = link_whatsapp.split("phone=")[1].split("&")[0]
                    print(f"[{GREEN}{contador}{RESET}] Nome: {nome} Telefone: {telefone_whatsapp}")
                    nome_arquivo_excel = f"dados_coletados_{local}.xlsx"
                    salvar_dados_excel(nome, telefone_whatsapp, nome_arquivo_excel)
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])

                    contador +=1
                except:
                    pass
    except:
        pass
    page += 1
    if page > 70:
        print("Total de paginas batida...")
        sair_loop = True
        

driver.close()
    