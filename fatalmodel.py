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
    nome, telefone1, telefone2, nome_arquivo_excel="dados_coletados_goiania03.xlsx"
):
    # Verifica se o arquivo Excel já existe
    if os.path.exists(nome_arquivo_excel):
        # Se existir, carrega os dados existentes
        df = pd.read_excel(nome_arquivo_excel)
    else:
        # Se não existir, cria um novo DataFrame com as colunas
        df = pd.DataFrame(columns=["Nome", "Telefone1", "Telefone2", "Site"])

    # Cria um DataFrame com o novo dado
    novo_dado = pd.DataFrame(
        {
            "Nome": [nome],
            "Telefone1": [telefone1],
            "Telefone2": [telefone2],
            "Site": "FatalModel",
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

# URL da página que você deseja acessar
url = "https://fatalmodel.com/acompanhantes-goiania-go"
driver.get(url)

try:

    last_height = driver.execute_script("return document.body.scrollHeight")
    printed_names = set()  # Conjunto para armazenar nomes já impressos
    id_counter = 0  # Contador para IDs sequenciais
    time.sleep(10)
    while True:
        try:
            # Localiza a seção específica
            section = driver.find_element(By.CSS_SELECTOR, "section.listing-section")

            # Encontra todos os elementos <h2> dentro da seção
            h2_elements = section.find_elements(By.CSS_SELECTOR, "h2.text-base")

            # Percorre os elementos <h2> e imprime o texto de cada um, se for novo
            for h2 in h2_elements:
                try:
                    name = h2.text
                    if name not in printed_names:
                        # Verifica se o ícone do telefone está presente no card
                        card = h2.find_element(
                            By.XPATH,
                            "./ancestor::div[contains(@class, 'no-tap-highlight')]",
                        )
                        try:
                            # Rola até o card para garantir que esteja visível
                            driver.execute_script(
                                "arguments[0].scrollIntoView(true);", card
                            )

                            # Procura pelo ícone do telefone
                            phone_icon = card.find_element(
                                By.XPATH, ".//i[contains(@class, 'icon-phone-outline')]"
                            )
                            button = phone_icon.find_element(
                                By.XPATH, "./ancestor::button"
                            )

                            button.click()

                            # Aguarda o modal aparecer
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "telephone1"))
                            )
                            phone1 = "Sem telefone"
                            phone2 = "Sem telefone"
                            # Coleta os números de telefone
                            try:
                                phone1 = driver.find_element(By.ID, "telephone1").text
                            except:
                                phone1 = "Sem telefone"

                            try:
                                phone2 = driver.find_element(By.ID, "telephone2").text
                            except:
                                phone2 = "Sem telefone"

                            # Fecha o modal
                            close_button = driver.find_element(
                                By.CSS_SELECTOR, "i.icon-close"
                            )
                            close_button.click()

                            # Formata a saída
                            if phone1 or phone2:
                                phone_info = f"Tel01: {phone1 if phone1 else 'N/A'} Tel02: {phone2 if phone2 else 'N/A'}"
                            else:
                                phone_info = "Sem telefone"

                            phone_status = "\033[92mSIM\033[0m"  # Verde
                        except:
                            try:
                                # data_event_params = card.get_attribute(
                                #    "data-event-params"
                                # )
                                ## Verifica se o card possui "Exclusivo premium"
                                # if (
                                #    data_event_params
                                #    and "exclusivo_premium" in data_event_params
                                # ):
                                #    phone1 = "Exclusivo premium"
                                #    phone_status = "\033[93mEXCLUSIVO\033[0m"  # Amarelo
                                #    phone_info = f"Tel01: {phone1} Tel02: Sem telefone"
                                #    continue  # Pula para o próximo card

                                print(f"botao de telefone nao encontrado para {name}")
                                # Se não encontrar o botão, abre o perfil em uma nova guia
                                profile_link = card.find_element(
                                    By.CSS_SELECTOR, "a.link-mapped"
                                ).get_attribute("href")
                                driver.execute_script(
                                    "window.open(arguments[0]);", profile_link
                                )
                                driver.switch_to.window(driver.window_handles[1])
                                time.sleep(1)
                                # Verifica o tipo de botão e clica se for o botão correto
                                try:
                                    whatsapp_button = WebDriverWait(driver, 10).until(
                                        EC.presence_of_element_located(
                                            (
                                                By.CSS_SELECTOR,
                                                "button.profile-contact__whatsapp",
                                            )
                                        )
                                    )
                                    # Verifica se o botão é "Telefone exclusivo Premium"
                                    if (
                                        "Telefone exclusivo Premium"
                                        in whatsapp_button.text
                                    ):
                                        print("Exclusivo premium")
                                        phone1 = "Exclusivo Premium"

                                    else:
                                        whatsapp_button.click()
                                        time.sleep(1)
                                        # Aguarda o modal aparecer e captura o telefone
                                        WebDriverWait(driver, 10).until(
                                            EC.presence_of_element_located(
                                                (
                                                    By.CSS_SELECTOR,
                                                    "p.phone-modal__whatsapp-number",
                                                )
                                            )
                                        )
                                        time.sleep(1)
                                        phone1 = driver.find_element(
                                            By.CSS_SELECTOR,
                                            "p.phone-modal__whatsapp-number",
                                        ).text
                                except Exception as e:
                                    print(f"Erro ao encontrar ou clicar no botão: {e}")
                                # Fecha a guia e volta para a anterior
                                time.sleep(3)
                                driver.close()

                                driver.switch_to.window(driver.window_handles[0])
                                if not phone1:
                                    phone_status = "\033[91mNÃO\033[0m"  # Vermelho
                                elif phone1 == "Exclusivo Premium":
                                    phone_status = "\033[93mEXCLUSIVO\033[0m"  # Amarelo
                                else:
                                    phone_status = "\033[92mSIM\033[0m"  # Verde

                                phone_info = f"Tel01: {phone1 if phone1 else 'N/A'} Tel02: Sem telefone"
                            except Exception as ex:
                                print(f"Erro ao abrir perfil de {ex}")
                                phone_status = "\033[91mNÃO\033[0m"  # Vermelho
                                phone_info = "Sem telefone"

                        print(
                            f"ID {id_counter}: {name} - Telefone: {phone_status} {phone_info}"
                        )
                        printed_names.add(name)
                        id_counter += 1
                        if phone_info == "Sem telefone":
                            phone1 = "Sem telefone"
                            phone2 = "Sem telefone"
                        salvar_dados_excel(name, phone1, phone2)
                        time.sleep(3)
                except:
                    print(f"Erro ao processar")
                    time.sleep(10)
                    driver.execute_script("window.open(arguments[0]);", url)

                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
            # Rola a página para baixo
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            # Aguarda o carregamento da nova seção
            time.sleep(2)

            # Calcula a nova altura da página e compara com a última altura
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
        except:
            print("Erro ao rolar a página esperando 10s")
            time.sleep(10)
            driver.execute_script("window.open(arguments[0]);", url)

            driver.close()
            driver.switch_to.window(driver.window_handles[0])
finally:
    # Fecha o navegador
    driver.quit()
