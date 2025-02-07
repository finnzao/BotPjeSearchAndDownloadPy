import os
import json
import math
import re
import time
import logging
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from openpyxl import Workbook
from openpyxl.styles import Font

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Variáveis globais para o driver e o wait
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

    # Caso a busca utilize OAB
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
        total_results_element = wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "//table[contains(@id, 'processosTable')]//tfoot//span[contains(text(), 'resultados encontrados')]")
            )
        )
        try:
            total_results_text = driver.find_element(By.XPATH, "//table[contains(@id, 'processosTable')]//tfoot").text
            logging.info(f"Texto do rodapé da tabela: {total_results_text}")
            match = re.search(r'(\d+)\s+resultados encontrados', total_results_text)
            if match:
                total_results_number = int(match.group(1))
                total_pages = math.ceil(total_results_number / 20)
                return total_pages
            else:
                logging.warning("Não foi possível extrair o número total de resultados.")
                return 0
        except Exception as e:
            logging.error(f"Erro ao obter o número total de páginas: {e}")
            return 0
    except Exception as e:
        logging.error(f"Erro ao obter o número total de páginas: {e}")
        return 0

def collect_process_numbers():
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, 'fPP:processosTable:tb')))
    numProcessos = []
    max_itens_por_pagina = 20  # Ajuste esse valor se necessário

    total_pages = get_total_pages()
    logging.info(f"Total de páginas para processar: {total_pages}")

    for page_num in range(1, total_pages + 1):
        logging.info(f"Processando página {page_num} de {total_pages}")
        table_body = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.ID, 'fPP:processosTable:tb'))
        )

        rows = table_body.find_elements(By.XPATH, "./tr")
        logging.info(f"Número de processos encontrados na página: {len(rows)}")

        for row in rows:
            try:
                first_td = row.find_element(By.XPATH, "./td[1]")
                a_tag = first_td.find_element(By.TAG_NAME, "a")
                numero_do_processo = a_tag.get_attribute('title')
                logging.info(f"Processo encontrado: {numero_do_processo}")
                numProcessos.append(numero_do_processo)
            except Exception as e:
                logging.error(f"Não foi possível extrair o número do processo na linha: {e}")
                continue

        if page_num < total_pages:
            try:
                wait.until(EC.invisibility_of_element((By.ID, 'j_id136:modalStatusCDiv')))
                next_page_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//td[contains(@onclick, 'next')]"))
                )
                next_page_button.click()
                wait.until(EC.staleness_of(table_body))
            except Exception as e:
                logging.error(f"Erro ao navegar para a próxima página: {e}")
                break
        else:
            logging.info("Última página alcançada.")
            time.sleep(2)

    return numProcessos

def save_to_json(data, filename="ResultadoProcessosPesquisa.json"):
    dir_path = "./docs"
    file_path = f"{dir_path}/{filename}"

    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

    with open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)

    logging.info(f"Arquivo '{file_path}' criado com sucesso.")

def save_data_to_excel(data_list, filename="./docs/dados_partes.xlsx"):
    """
    Salva a lista de dicionários em um arquivo Excel.

    :param data_list: Lista de dicionários contendo os dados a serem salvos.
    :param filename: Nome do arquivo Excel de saída.
    """
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Dados das Partes"

        # Cabeçalhos atualizados em português
        headers = ['Número do Processo', 'Polo', 'Nome da Parte', 'CPF', 'Nome Civil', 
                   'Data de Nascimento', 'Genitor', 'Genitora', 'Classe', 'Assunto', 'Área']
        ws.append(headers)

        # Estilização dos cabeçalhos
        bold_font = Font(bold=True)
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = bold_font

        # Adicionar os dados
        for data in data_list:
            ws.append([
                data.get('Número do Processo', ''),
                data.get('Polo', ''),
                data.get('Nome da Parte', ''),
                data.get('CPF', ''),
                data.get('Nome Civil', ''),
                data.get('Data de Nascimento', ''),
                data.get('Genitor', ''),
                data.get('Genitora', ''),
                data.get('Classe', ''),
                data.get('Assunto', ''),
                data.get('Área', '')
            ])

        # Ajustar a largura das colunas
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except Exception:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width

        wb.save(filename)
        logging.info(f"Dados salvos com sucesso no arquivo '{filename}'.")

    except Exception as e:
        logging.error(f"Ocorreu uma exceção ao salvar os dados no Excel. Erro: {e}")
        raise e

def main():
    load_dotenv()
    initialize_driver()
    user, password = os.getenv("USER"), os.getenv("PASSWORD")
    profile = os.getenv("PROFILE")
    optionSearch = {
        'nomeParte': '',
        'numOrgaoJustica': '0216',
        'Assunto': '',
        'NomeDoRepresentante': '',
        'Alcunha': '',
        'classeJudicial': 'AUTO DE PRISÃO EM FLAGRANTE',
        'numDoc': '',
        'estadoOAB': '',
        'numeroOAB': ''
    }

    login(user, password)
    # Caso seja necessário pular o token, descomente a próxima linha:
    # skip_token()
    select_profile(profile)
    search_process(optionSearch)
    time.sleep(20)
    process_numbers = collect_process_numbers()
    driver.quit()

    logging.info(f"Números de processos coletados: {process_numbers}")

    # Salvar os números dos processos em formato JSON
    if process_numbers:
        save_to_json(process_numbers)

        # Montar a lista de dicionários para salvar em Excel.
        # Aqui, preenchemos apenas o 'Número do Processo' e deixamos os demais campos vazios.
        data_list = []
        for num in process_numbers:
            data_list.append({
                'Número do Processo': num,
                'Polo': '',
                'Nome da Parte': '',
                'CPF': '',
                'Nome Civil': '',
                'Data de Nascimento': '',
                'Genitor': '',
                'Genitora': '',
                'Classe': '',
                'Assunto': '',
                'Área': ''
            })

        save_data_to_excel(data_list)
    else:
        logging.info("Nenhum processo encontrado para salvar.")

if __name__ == "__main__":
    main()
