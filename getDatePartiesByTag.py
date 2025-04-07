from selenium import webdriver
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    TimeoutException,
    NoSuchElementException
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl.styles import Font
from functools import wraps
import time
from dotenv import load_dotenv
import os
import logging
import re

from utils.pje_automation import PjeConsultaAutomator

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("script_log_getDatePartiesByTag.log"),
        logging.StreamHandler()
    ]
)

class PJEAutomationGetInfoTagParties(PjeConsultaAutomator):
    @staticmethod
    def retry(max_retries=2):
        def decorator(func):
            @wraps(func)
            def wrapper(self, *args, **kwargs):
                retries = 0
                while retries < max_retries:
                    try:
                        return func(self, *args, **kwargs)
                    except (TimeoutException, StaleElementReferenceException) as e:
                        retries += 1
                        logging.warning(f"Tentativa {retries} falhou com erro: {e}. Tentando novamente...")
                        if retries >= max_retries:
                            logging.error(f"Falha ao executar {func.__name__} após {max_retries} tentativas")
                            raise TimeoutException(f"Falha ao executar {func.__name__} após {max_retries} tentativas")
            return wrapper
        return decorator

    def __init__(self):
        super().__init__()
        self.process_data_list = []

    @retry()
    def search_on_tag(self, search):
        self.switch_to_ng_frame()
        logging.info("Alternado para o frame 'ngFrame'.")
        original_handles = set(self.driver.window_handles)
        logging.info(f"Handles originais das janelas: {original_handles}")
        self.nav_tag()
        self.input_tag(search)

    def nav_tag(self):
        xpath = "/html/body/app-root/selector/div/div/div[1]/side-bar/nav/ul/li[5]/a"
        self.click_element(xpath)
        logging.info("Navegando para a seção de etiquetas.")

    def input_tag(self, search_text):
        search_input = self.wait.until(EC.element_to_be_clickable((By.ID, "itPesquisarEtiquetas")))
        search_input.clear()
        search_input.send_keys(search_text)
        logging.info(f"Texto de pesquisa inserido: {search_text}")
        current_handles = set(self.driver.window_handles)
        self.click_element("/html/body/app-root/selector/div/div/div[2]/right-panel/div/etiquetas/div[1]/div/div[1]/div[2]/div[1]/span/button[1]")
        logging.info("Botão de pesquisa de etiquetas clicado.")
        time.sleep(2)
        new_handles = set(self.driver.window_handles)
        if len(new_handles) > len(current_handles):
            data_window = (new_handles - current_handles).pop()
            self.driver.switch_to.window(data_window)
            logging.info(f"Alternado para a nova janela de pesquisa de etiquetas: {data_window}")
            self.driver.close()
            logging.info("Nova janela de pesquisa de etiquetas fechada.")
            self.driver.switch_to.default_content()
            logging.info("Retornado para a janela original após fechar a nova janela.")
        else:
            logging.info("Nenhuma nova janela foi aberta após a pesquisa de etiquetas.")
        self.click_element("/html/body/app-root/selector/div/div/div[2]/right-panel/div/etiquetas/div[1]/div/div[2]/ul/p-datalist/div/div/ul/li/div/li/div[2]/span/span")
        logging.info("Etiqueta selecionada na lista de resultados.")

    def get_process_list(self):
        try:
            process_xpath = "//processo-datalist-card"
            processes = self.wait.until(EC.presence_of_all_elements_located((By.XPATH, process_xpath)))
            logging.info(f"Número de processos encontrados: {len(processes)}")
            return processes
        except Exception as e:
            self.driver.save_screenshot("get_process_list_exception.png")
            logging.error(f"Ocorreu uma exceção ao obter a lista de processos. Captura de tela salva como 'get_process_list_exception.png'. Erro: {e}")
            raise e

    def click_on_process(self, process_element):
        try:
            original_handles = set(self.driver.window_handles)
            self.driver.execute_script("arguments[0].scrollIntoView(true);", process_element)
            self.driver.execute_script("arguments[0].click();", process_element)
            logging.info("Processo clicado com sucesso!")
            new_window_handle = self.switch_to_new_window(original_handles)
            logging.info(f"Alternado para a nova janela do processo: {new_window_handle}")
            return new_window_handle
        except Exception as e:
            self.driver.save_screenshot("click_on_process_exception.png")
            logging.error(f"Ocorreu uma exceção ao clicar no processo. Captura de tela salva como 'click_on_process_exception.png'. Erro: {e}")
            raise e

    def switch_to_new_window(self, original_handles, timeout=20):
        try:
            WebDriverWait(self.driver, timeout).until(
                lambda d: len(d.window_handles) > len(original_handles)
            )
            new_handles = set(self.driver.window_handles) - original_handles
            if new_handles:
                new_window = new_handles.pop()
                self.driver.switch_to.window(new_window)
                logging.info(f"Alternado para a nova janela: {new_window}")
                return new_window
            else:
                raise TimeoutException("Nova janela não foi encontrada dentro do tempo especificado.")
        except TimeoutException as e:
            self.driver.save_screenshot("switch_to_new_window_timeout.png")
            logging.error("TimeoutException: Não foi possível encontrar a nova janela. Captura de tela salva como 'switch_to_new_window_timeout.png'")
            raise e

    @retry()
    def click_element(self, xpath):
        try:
            element = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
            try:
                element.click()
                logging.info(f"Elemento clicado com sucesso: {xpath}")
            except (ElementClickInterceptedException, Exception) as e:
                logging.warning(f"Erro ao clicar no elemento normalmente: {e}. Tentando com JavaScript...")
                self.driver.execute_script("arguments[0].click();", element)
                logging.info(f"Elemento clicado com sucesso usando JavaScript: {xpath}")
        except Exception as e:
            self.driver.save_screenshot("click_element_exception.png")
            logging.error(f"Ocorreu uma exceção ao clicar no elemento. Captura de tela salva como 'click_element_exception.png'. Erro: {e}")
            raise e

    def collect_data_parties(self):
        try:
            logging.info("Iniciando coleta de dados das partes.")
            data = {}
            fields = {
                'CPF': '//*[@id="pessoaFisicaViewView:j_id58"]/div/div[2]',
                'Nome Civil': '//*[@id="pessoaFisicaViewView:j_id80"]/div/div[2]',
                'Data de Nascimento': '//*[@id="pessoaFisicaViewView:j_id157"]/div/div[2]',
                'Genitor': '//*[@id="pessoaFisicaViewView:j_id168"]/div/div[2]',
                'Genitora': '//*[@id="pessoaFisicaViewView:j_id179"]/div/div[2]',
            }
            for field_name, xpath in fields.items():
                try:
                    element = self.driver.find_element(By.XPATH, xpath)
                    data[field_name] = element.text.strip()
                    logging.info(f"{field_name}: {data[field_name]}")
                except NoSuchElementException:
                    logging.warning(f"{field_name} não encontrado.")
                    data[field_name] = ''
            logging.info(f"Dados coletados: {data}")
            return data
        except Exception as e:
            logging.error(f"Ocorreu uma exceção ao coletar dados das partes: {e}")
            self.driver.save_screenshot("collectDataParties_exception.png")
            raise e

    def collect_process_info(self):
        try:
            logging.info("Coletando informações adicionais do processo.")
            process_info = {}
            fields = {
                'Classe': '//*[@id="classeProcesso"]',
                'Assunto': '//*[@id="assuntoProcesso"]',
                'Área': '//*[@id="areaProcesso"]',
            }
            for field_name, xpath in fields.items():
                try:
                    element = self.driver.find_element(By.XPATH, xpath)
                    process_info[field_name] = element.text.strip()
                    logging.info(f"{field_name}: {process_info[field_name]}")
                except NoSuchElementException:
                    logging.warning(f"{field_name} não encontrado.")
                    process_info[field_name] = ''
            return process_info
        except Exception as e:
            logging.error(f"Ocorreu uma exceção ao coletar informações do processo: {e}")
            self.driver.save_screenshot("collectProcessInfo_exception.png")
            raise e

    def get_data_parties(self, process_window_handle, process_number, process_info):
        try:
            self.driver.switch_to.window(process_window_handle)
            logging.info("Alternado para a janela do processo.")
            self.click_element('//*[@id="navbar"]/ul/li/a[1]')
            logging.info("Elemento da navbar clicado com sucesso.")
            self.wait.until(EC.presence_of_element_located((By.ID, 'poloPassivo')))
            polo_div = self.driver.find_element(By.ID, 'poloPassivo')
            party_links = polo_div.find_elements(By.CSS_SELECTOR, 'tbody tr td a')
            logging.info(f"Encontrado {len(party_links)} partes no polo passivo")
            logging.info(f"Links das partes: {party_links}")
            for index in range(len(party_links)):
                polo_div = self.driver.find_element(By.ID, 'poloPassivo')
                party_links = polo_div.find_elements(By.CSS_SELECTOR, 'tbody tr td a')
                party_link = party_links[index]
                party_name = party_link.text.strip()
                handles_before_click = set(self.driver.window_handles)
                self.driver.execute_script("arguments[0].click();", party_link)
                logging.info("Link da parte clicado")
                WebDriverWait(self.driver, 10).until(EC.new_window_is_opened(handles_before_click))
                handles_after_click = set(self.driver.window_handles)
                new_handles = handles_after_click - handles_before_click
                party_window_handle = new_handles.pop()
                self.driver.switch_to.window(party_window_handle)
                logging.info("Aba de dados da parte aberta")
                data = self.collect_data_parties()
                data['Número do Processo'] = process_number
                data['Polo'] = 'Passivo'
                data['Nome da Parte'] = party_name
                data.update(process_info)
                self.process_data_list.append(data)
                self.driver.close()
                logging.info("Aba de dados da parte fechada")
                self.driver.switch_to.window(process_window_handle)
                logging.info("Retornando para a janela do processo")
                time.sleep(0.5)
        except Exception as e:
            logging.error(f"Falha em coletar dados das partes. Erro: {e}")
            self.driver.save_screenshot("getDataParties_exception.png")
            raise e

    def switch_to_ng_frame(self):
        try:
            self.driver.switch_to.default_content()
            self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
            logging.info("Alternado para o frame 'ngFrame'.")
        except TimeoutException:
            logging.error("Timeout ao tentar alternar para o frame 'ngFrame'.")
            self.driver.save_screenshot("switch_to_ngFrame_timeout.png")
            raise

    def info_parties_process_on_tag_search(self):
        try:
            original_window = self.driver.current_window_handle
            self.switch_to_ng_frame()
            process_list = self.get_process_list()
            total_processes = len(process_list)
            logging.info(f"Total de processos a serem processados: {total_processes}")
            for index in range(1, total_processes + 1):
                logging.info(f"Iniciando o processamento do processo {index} de {total_processes}")
                self.switch_to_ng_frame()
                process_xpath = f"(//processo-datalist-card)[{index}]//a/div/span[2]"
                logging.info(f"XPath gerado: {process_xpath}")
                try:
                    process_element = self.wait.until(EC.element_to_be_clickable((By.XPATH, process_xpath)))
                except TimeoutException:
                    logging.error(f"Timeout ao localizar o elemento do processo no índice {index} com XPath: {process_xpath}")
                    self.driver.save_screenshot(f"process_element_{index}_timeout.png")
                    continue
                raw_process_number = process_element.text.strip()
                process_number = re.sub(r'\D', '', raw_process_number)
                if len(process_number) >= 17:
                    process_number = f"{process_number[:7]}-{process_number[7:9]}.{process_number[9:13]}.{process_number[13]}.{process_number[14:16]}.{process_number[16:]}"
                else:
                    process_number = raw_process_number
                logging.info(f"Número do Processo: {process_number}")
                print(process_number)
                try:
                    process_window_handle = self.click_on_process(process_element)
                except Exception as e:
                    logging.error(f"Falha ao clicar no processo no índice {index}: {e}")
                    self.driver.save_screenshot(f"click_process_{index}_exception.png")
                    continue
                process_info = self.collect_process_info()
                try:
                    self.get_data_parties(process_window_handle=process_window_handle, process_number=process_number, process_info=process_info)
                except Exception as e:
                    logging.error(f"Falha ao coletar dados para o processo {process_number}: {e}")
                    self.driver.save_screenshot(f"getDataParties_{process_number}_exception.png")
                try:
                    self.driver.close()
                    logging.info("Aba do processo fechada com sucesso.")
                except Exception as e:
                    logging.error(f"Falha ao fechar a aba do processo {process_number}: {e}")
                try:
                    self.driver.switch_to.window(original_window)
                    logging.info("Retornado para a janela original.")
                except Exception as e:
                    logging.error(f"Falha ao retornar para a janela original: {e}")
                time.sleep(1)
            logging.info("Processamento concluído.")
            return self.process_data_list
        except Exception as e:
            self.driver.save_screenshot("InfoPartiesProcessOnTagSearch_exception.png")
            logging.error(f"Ocorreu uma exceção em 'InfoPartiesProcessOnTagSearch'. Captura de tela salva como 'InfoPartiesProcessOnTagSearch_exception.png'. Erro: {e}")
            raise e

def save_data_to_excel(data_list, filename="dados_partes.xlsx"):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Dados das Partes"
        headers = ['Número do Processo', 'Polo', 'Nome da Parte', 'CPF', 'Nome Civil', 'Data de Nascimento', 'Genitor', 'Genitora', 'Classe', 'Assunto', 'Área']
        ws.append(headers)
        bold_font = Font(bold=True)
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = bold_font
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
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
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
    automation = PJEAutomationGetInfoTagParties()
    try:
        user, password = os.getenv("USER"), os.getenv("PASSWORD")
        automation.login(user, password)
        profile = os.getenv("PROFILE")
        automation.select_profile("VARA CRIMINAL DE RIO REAL / Direção de Secretaria / Diretor de Secretaria")
        automation.search_on_tag("Possivel OBT")
        process_data_list = automation.info_parties_process_on_tag_search()
        save_data_to_excel(process_data_list)
        time.sleep(5)
    finally:
        automation.driver.quit()
        logging.info("Driver encerrado.")

if __name__ == "__main__":
    main()
