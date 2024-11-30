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
    nome, telefone1, nome_arquivo_excel="dados_coletados_garotacomlocal_go.xlsx"
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

url = f"https://garotacomlocal.com/acompanhantes/goiania/"
driver.get(url)


try:

    last_height = driver.execute_script("return document.body.scrollHeight")
    printed_names = set()  # Conjunto para armazenar nomes já impressos
    id_counter = 0  # Contador para IDs sequenciais
    time.sleep(10)
    while True:
        try:
            encontrado = []
            nao_encontrado = 0
            # Localiza a seção específica
            section = driver.find_element(By.ID, "post-17739")
            uls = section.find_elements(By.TAG_NAME, "ul")
            qtd_lead = 0
            for ul in uls:
                lis = ul.find_elements(By.TAG_NAME, "li")
                for li in lis:
                    try:
                        a_tag = li.find_element(
                            By.CSS_SELECTOR, "a.job_listing-clickbox"
                        )
                        link_perfil = a_tag.get_attribute("href")
                        driver.execute_script("window.open(arguments[0]);", link_perfil)
                        driver.switch_to.window(driver.window_handles[-1])

                        try:
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located(
                                    (By.CLASS_NAME, "botoesChatAnunciante")
                                )
                            )
                            botao_chat = driver.find_element(
                                By.CLASS_NAME, "botoesChatAnunciante"
                            )
                            link_whatsapp = botao_chat.get_attribute("href")

                            telefone = link_whatsapp.split("phone=")[1].split("&")[0]
                            nome = link_perfil.split("/")[-2]
                            qtd_lead += 1
                            print(
                                f"{YELLOW}{qtd_lead}{RESET}: Nome: {nome}, Telefone: {telefone}"
                            )
                            salvar_dados_excel(nome, telefone)
                        except Exception as e:
                            print(f"Erro ao processar o perfil: {e}")
                        finally:
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])

                        encontrado.append(a_tag.get_attribute("href"))
                    except:
                        print("Nenhum encontrado")
                        nao_encontrado += 1
                        continue

            print(f"Encontrados: {len(encontrado)}, Nao encontrados: {nao_encontrado}")
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
