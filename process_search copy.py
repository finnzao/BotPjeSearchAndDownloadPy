from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import os
import json
import math
import re
import time
# Declarar as variáveis globais
driver = None
wait = None

def initialize_driver():
    global driver, wait
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 20)

def login(user, password):
    login_url = 'https://pje.tjba.jus.br/pje/login.seam'
    driver.get(login_url)
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ssoFrame')))
    username_input = wait.until(EC.presence_of_element_located((By.ID, 'username')))
    password_input = wait.until(EC.presence_of_element_located((By.ID, 'password')))
    login_button = wait.until(EC.presence_of_element_located((By.ID, 'kc-login')))
    username_input.send_keys(user)
    password_input.send_keys(password)
    login_button.click()
    driver.switch_to.default_content()

def skip_token():
    proceed_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Prosseguir sem o Token')]"))
    )
    proceed_button.click()

def select_profile(profile):
    dropdown = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'dropdown-toggle')))
    dropdown.click()
    button_xpath = f"//a[contains(text(), '{profile}')]"
    desired_button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
    driver.execute_script("arguments[0].scrollIntoView(true);", desired_button)
    driver.execute_script("arguments[0].click();", desired_button)

def search_process(optionSearch):
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
    icon_search_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'li#liConsultaProcessual i.fas'))
    ) 
    icon_search_button.click()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameConsultaProcessual')))

    ElementoNumOrgaoJutica = wait.until(EC.presence_of_element_located((By.ID, 'fPP:numeroProcesso:NumeroOrgaoJustica')))
    ElementoNumOrgaoJutica.send_keys(optionSearch.get('numOrgaoJustica'))

    # OAB
    if optionSearch.get('estadoOAB'):
        ElementoNumeroOAB = wait.until(EC.presence_of_element_located((By.ID, 'fPP:decorationDados:numeroOAB')))
        ElementoNumeroOAB.send_keys(optionSearch.get('numeroOAB'))
        ElementoEstadosOAB = wait.until(EC.presence_of_element_located((By.ID, 'fPP:decorationDados:ufOABCombo')))
        listaEstadosOAB = Select(ElementoEstadosOAB)
        listaEstadosOAB.select_by_value(optionSearch.get('estadoOAB'))  # Exemplo: 'BA'

    consulta_classe = wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id245:classeJudicial')))
    consulta_classe.send_keys(optionSearch.get('classeJudicial'))

    ElementonomeDaParte = wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id150:nomeParte')))
    ElementonomeDaParte.send_keys(optionSearch.get('nomeParte'))

    btnProcurarProcesso = wait.until(EC.presence_of_element_located((By.ID, 'fPP:searchProcessos')))
    btnProcurarProcesso.click()

def get_total_pages():
    try:
        # XPath atualizado para localizar o elemento corretamente
        total_results_element = wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "//table[contains(@id, 'processosTable')]//tfoot//span[contains(text(), 'resultados encontrados')]")
            )
        )
        total_results_text = total_results_element.text
        print(f"Texto dos resultados: '{total_results_text}'")
        
        # Usar expressões regulares para extrair o número de resultados
        match = re.search(r'(\d+)\s+resultados encontrados', total_results_text)
        if match:
            total_results_number = int(match.group(1))
            total_pages = math.ceil(total_results_number / 20)
            return total_pages
        else:
            print("Não foi possível extrair o número total de resultados.")
            return 0
    except Exception as e:
        print(f"Erro ao obter o número total de páginas: {e}")
        return 0

def collect_process_numbers():
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, 'fPP:processosTable:tb')))
    numProcessos = []
    max_itens_por_pagina = 20  # Ajuste este valor se o número máximo por página for diferente

    while True:
        print("Processando nova página")
        # Aguarde até que os links dos processos estejam visíveis
        WebDriverWait(driver, 50).until(
            EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "a.btn-link.btn-condensed"))
        )

        # Encontre todos os links dos processos na página atual
        links_dos_processos = driver.find_elements(By.CSS_SELECTOR, "a.btn-link.btn-condensed")

        #print(f"Número de processos encontrados na página: {len(links_dos_processos)}")
        print(len(links_dos_processos))
        # Extraia o número do processo de cada link e adicione à lista
        for link in links_dos_processos:
            numero_do_processo = link.get_attribute('title')
            numProcessos.append(numero_do_processo)

        # Se o número de processos na página for menor que o número máximo por página, estamos na última página
        if len(links_dos_processos) < max_itens_por_pagina:
            print("Última página alcançada.")
            time.sleep(2)  # Esperar 2 segundos para garantir que todos os itens sejam capturados
            break

        # Tentar clicar no botão 'Próxima'
        try:
            # Aguardar até que qualquer elemento de carregamento desapareça
            wait.until(
                EC.invisibility_of_element((By.ID, 'j_id136:modalStatusCDiv'))
            )

            # Tente localizar e clicar no botão 'Próxima página'
            next_page_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//td[contains(@onclick, 'next')]"))
            )
            next_page_button.click()
        except Exception as e:
            print(f"Erro ao navegar para a próxima página: {e}")
            break

    numUnico = set(numProcessos)
    numUnicosLista = list(numUnico)
    return numUnicosLista


def save_to_json(data, filename="ResultadoProcessosPesquisa.json"):
    # Definir o caminho do arquivo
    dir_path = "./docs"
    file_path = f"{dir_path}/{filename}"

    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

    with open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)

    print(f"Arquivo '{file_path}' criado com sucesso.")

def main():
    load_dotenv()
    initialize_driver()
    user, password = os.getenv("USER"), os.getenv("PASSWORD")
    profile = os.getenv("PROFILE")
    optionSearch = {
        'nomeParte': 'EDMILSON DO NASCIMENTO SANTOS',
        'numOrgaoJustica': '0216',
        'Assunto': '',
        'NomeDoRepresentante': '',
        'Alcunha': '',
        'classeJudicial': '',
        'numDoc': '',
        'estadoOAB': '',
        'numeroOAB': ''
    }

    login(user, password)
    # skip_token()
    select_profile(profile)
    search_process(optionSearch)
    time.sleep(20)
    process_numbers = collect_process_numbers()
    driver.quit()

    # Salvar os números dos processos em formato JSON
    if process_numbers:
        save_to_json(process_numbers)
    else:
        print("Nenhum processo encontrado para salvar.")

if __name__ == "__main__":
    main()
