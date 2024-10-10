from selenium import webdriver
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    TimeoutException,
    NoSuchElementException
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl.styles import Font
from functools import wraps
import time
from dotenv import load_dotenv
import os
import logging
import re

# Configuração básica do logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("script_log.log"),
        logging.StreamHandler()
    ]
)

# Variáveis globais para driver, wait e lista de dados
driver = None
wait = None
process_data_list = []  # Lista para armazenar os dados coletados

# Decorador para gerenciar tentativas de execução com retry
def retry(max_retries=2):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            retries = 0
            while retries < max_retries:
                try:
                    return func(*args, **kwargs)
                except (TimeoutException, StaleElementReferenceException) as e:
                    retries += 1
                    logging.warning(f"Tentativa {retries} falhou com erro: {e}. Tentando novamente...")
                    if retries >= max_retries:
                        logging.error(f"Falha ao executar {func.__name__} após {max_retries} tentativas")
                        raise TimeoutException(f"Falha ao executar {func.__name__} após {max_retries} tentativas")
        return wrapper
    return decorator

# Inicialização do driver e wait
def initialize_driver():
    global driver, wait
    # Configurações do Chrome
    chrome_options = webdriver.ChromeOptions()

    # Diretório para onde os downloads serão salvos
    download_directory = os.path.abspath("C:/Users/lfmdsantos/Downloads/processosBaixados")  # Altere para o caminho desejado

    # Criar o diretório se não existir
    os.makedirs(download_directory, exist_ok=True)

    # Configurações de preferências
    prefs = {
        "plugins.always_open_pdf_externally": True,  # Baixar PDFs em vez de abrir no navegador
        "download.default_directory": download_directory,  # Diretório padrão de download
        "download.prompt_for_download": False,  # Não perguntar onde salvar
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True  # Habilitar navegação segura
    }

    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 120, poll_frequency=2, ignored_exceptions=[NoSuchElementException])
    logging.info("Driver inicializado com sucesso.")

@retry()
def login(user, password):
    login_url = 'https://pje.tjba.jus.br/pje/login.seam'
    driver.get(login_url)
    logging.info(f"Navegando para a URL de login: {login_url}")

    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ssoFrame')))
    logging.info("Alternado para o frame 'ssoFrame'.")

    username_input = wait.until(EC.presence_of_element_located((By.ID, 'username')))
    password_input = wait.until(EC.presence_of_element_located((By.ID, 'password')))
    login_button = wait.until(EC.presence_of_element_located((By.ID, 'kc-login')))

    username_input.send_keys(user)
    password_input.send_keys(password)
    logging.info("Credenciais inseridas.")

    login_button.click()
    logging.info("Botão de login clicado.")

    driver.switch_to.default_content()
    logging.info("Voltando para o conteúdo principal após o login.")

@retry()
def select_profile(profile):
    dropdown = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'dropdown-toggle')))
    dropdown.click()
    logging.info("Dropdown de perfil clicado.")

    button_xpath = f"//a[contains(text(), '{profile}')]"
    desired_button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
    driver.execute_script("arguments[0].scrollIntoView(true);", desired_button)
    driver.execute_script("arguments[0].click();", desired_button)
    logging.info(f"Perfil '{profile}' selecionado com sucesso.")

@retry()
def search_on_tag(search):
    switch_to_ngFrame()
    logging.info("Alternado para o frame 'ngFrame'.")

    original_handles = set(driver.window_handles)
    logging.info(f"Handles originais das janelas: {original_handles}")    

    nav_tag()
    input_tag(search)

def nav_tag():
    xpath = "/html/body/app-root/selector/div/div/div[1]/side-bar/nav/ul/li[5]/a"
    click_element(xpath)
    logging.info("Navegando para a seção de etiquetas.")

def input_tag(search_text):
    search_input = wait.until(EC.element_to_be_clickable((By.ID, "itPesquisarEtiquetas")))
    search_input.clear()
    search_input.send_keys(search_text)
    logging.info(f"Texto de pesquisa inserido: {search_text}")

    # Capturar os handles antes do clique
    current_handles = set(driver.window_handles)

    click_element("/html/body/app-root/selector/div/div/div[2]/right-panel/div/etiquetas/div[1]/div/div[1]/div[2]/div[1]/span/button[1]")
    logging.info("Botão de pesquisa de etiquetas clicado.")

    # Esperar até que uma nova janela seja aberta, se aplicável
    time.sleep(2)  # Ajuste conforme necessário

    # Capturar os handles após o clique
    new_handles = set(driver.window_handles)

    # Verificar se uma nova janela foi aberta
    if len(new_handles) > len(current_handles):
        data_window = (new_handles - current_handles).pop()
        driver.switch_to.window(data_window)
        logging.info(f"Alternado para a nova janela de pesquisa de etiquetas: {data_window}")

        # Fechar a nova janela após a interação, se necessário
        driver.close()
        logging.info("Nova janela de pesquisa de etiquetas fechada.")

        # Voltar para a janela original
        driver.switch_to.default_content()
        logging.info("Retornado para a janela original após fechar a nova janela.")
    else:
        logging.info("Nenhuma nova janela foi aberta após a pesquisa de etiquetas.")

    # Continuar com a navegação
    click_element("/html/body/app-root/selector/div/div/div[2]/right-panel/div/etiquetas/div[1]/div/div[2]/ul/p-datalist/div/div/ul/li/div/li/div[2]/span/span")
    logging.info("Etiqueta selecionada na lista de resultados.")

def get_process_list():
    """
    Retorna uma lista de elementos que representam os processos encontrados.
    """
    try:
        # XPath para localizar todos os itens da lista de processos
        process_xpath = "//processo-datalist-card"
        processes = wait.until(EC.presence_of_all_elements_located((By.XPATH, process_xpath)))
        logging.info(f"Número de processos encontrados: {len(processes)}")
        return processes
    except Exception as e:
        driver.save_screenshot("get_process_list_exception.png")
        logging.error(f"Ocorreu uma exceção ao obter a lista de processos. Captura de tela salva como 'get_process_list_exception.png'. Erro: {e}")
        raise e

def click_on_process(process_element):
    try:
        original_handles = set(driver.window_handles)
        driver.execute_script("arguments[0].scrollIntoView(true);", process_element)
        driver.execute_script("arguments[0].click();", process_element)
        logging.info("Processo clicado com sucesso!")

        # Alternar para a nova janela aberta após clicar no processo
        new_window_handle = switch_to_new_window(original_handles)
        logging.info(f"Alternado para a nova janela do processo: {new_window_handle}")
        return new_window_handle

    except Exception as e:
        driver.save_screenshot("click_on_process_exception.png")
        logging.error(f"Ocorreu uma exceção ao clicar no processo. Captura de tela salva como 'click_on_process_exception.png'. Erro: {e}")
        raise e

def switch_to_new_window(original_handles, timeout=20):
    """
    Alterna para a nova janela que foi aberta após a execução de uma ação.

    :param original_handles: Set contendo os handles das janelas originais.
    :param timeout: Tempo máximo de espera para a nova janela aparecer.
    :return: Handle da nova janela.
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: len(d.window_handles) > len(original_handles)
        )
        new_handles = set(driver.window_handles) - original_handles
        if new_handles:
            new_window = new_handles.pop()
            driver.switch_to.window(new_window)
            logging.info(f"Alternado para a nova janela: {new_window}")
            return new_window
        else:
            raise TimeoutException("Nova janela não foi encontrada dentro do tempo especificado.")
    except TimeoutException as e:
        driver.save_screenshot("switch_to_new_window_timeout.png")
        logging.error("TimeoutException: Não foi possível encontrar a nova janela. Captura de tela salva como 'switch_to_new_window_timeout.png'")
        raise e

@retry()
def click_element(xpath):
    """
    Clica em um elemento localizado pelo XPath fornecido.
    Usa JavaScript como fallback se o clique normal falhar.
    
    :param xpath: O XPath do elemento a ser clicado.
    """
    try:
        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        try:
            element.click()
            logging.info(f"Elemento clicado com sucesso: {xpath}")
        except (ElementClickInterceptedException, Exception) as e:
            logging.warning(f"Erro ao clicar no elemento normalmente: {e}. Tentando com JavaScript...")
            driver.execute_script("arguments[0].click();", element)
            logging.info(f"Elemento clicado com sucesso usando JavaScript: {xpath}")
    except Exception as e:
        driver.save_screenshot("click_element_exception.png")
        logging.error(f"Ocorreu uma exceção ao clicar no elemento. Captura de tela salva como 'click_element_exception.png'. Erro: {e}")
        raise e

@retry()
def select_tipo_documento(tipoDocumento):
    """
    Seleciona o tipo de documento no dropdown com base no 'tipoDocumento' fornecido.

    :param tipoDocumento: O tipo de documento a ser selecionado (e.g., 'Despacho').
    """
    try:
        select_element = wait.until(EC.presence_of_element_located(
            (By.ID, 'navbar:cbTipoDocumento')
        ))
        select = Select(select_element)
        select.select_by_visible_text(tipoDocumento)
        logging.info(f"Tipo de documento '{tipoDocumento}' selecionado com sucesso.")
    except Exception as e:
        driver.save_screenshot("select_tipo_documento_exception.png")
        logging.error(f"Ocorreu uma exceção ao selecionar o tipo de documento. Captura de tela salva como 'select_tipo_documento_exception.png'. Erro: {e}")
        raise e

from selenium.common.exceptions import NoSuchElementException

def collectDataParties():
    """
    Coleta informações específicas das partes de um processo e as armazena em um dicionário.
    Se um elemento não for encontrado, atribui um valor padrão e continua.
    
    Retorna:
        dict: Dicionário contendo os dados coletados.
    """
    try:
        logging.info("Iniciando coleta de dados das partes.")
        # Alternar para o iframe correto, se necessário
        frames = driver.find_elements(By.TAG_NAME, 'iframe')
        logging.info(f"Número de iframes encontrados: {len(frames)}")
        if frames:
            driver.switch_to.frame(frames[0])  # Ajuste o índice conforme necessário
            logging.info("Alternado para o primeiro iframe.")
        else:
            logging.warning("Nenhum iframe encontrado. Continuando no contexto atual.")

        data = {}

        # Lista de campos e seus XPaths com IDs dinâmicos
        fields = {
            'CPF': '//*[@id="pessoaFisicaViewView:j_id58"]/div/div[2]',
            'Nome Civil': '//*[@id="pessoaFisicaViewView:j_id80"]/div/div[2]',
            'Data de Nascimento': '//*[@id="pessoaFisicaViewView:j_id157"]/div/div[2]',
            'Genitor': '//*[@id="pessoaFisicaViewView:j_id168"]/div/div[2]',
            'Genitora': '//*[@id="pessoaFisicaViewView:j_id179"]/div/div[2]',
        }

        for field_name, xpath in fields.items():
            try:
                element = driver.find_element(By.XPATH, xpath)
                data[field_name] = element.text.strip()
                logging.info(f"{field_name}: {data[field_name]}")
            except NoSuchElementException:
                logging.warning(f"{field_name} não encontrado.")
                data[field_name] = ''  # Atribuir valor padrão

        # Garantir que voltamos ao conteúdo padrão
        driver.switch_to.default_content()

        logging.info(f"Dados coletados: {data}")
        return data

    except Exception as e:
        logging.error(f"Ocorreu uma exceção ao coletar dados das partes: {e}")
        driver.save_screenshot("collectDataParties_exception.png")
        with open("collectDataParties_exception.html", "w", encoding='utf-8') as f:
            f.write(driver.page_source)
        raise e

def getDataParties(original_handles, process_number, process_window_handle):
    try:
        # Já estamos na janela do processo

        # Clicar no elemento da navbar para navegar até a seção de dados das partes
        click_element('//*[@id="navbar"]/ul/li/a[1]')
        logging.info("Elemento da navbar clicado com sucesso após sair do frame.")

        # Clicar no botão específico para acessar os dados das partes
        click_element('/html/body/div/div[1]/div/form/ul/li/ul/li/div[4]/table/tbody/tr/td/a')
        logging.info("Navegado para a página de dados das partes com sucesso.")

        # Esperar até que uma nova janela seja aberta
        try:
            WebDriverWait(driver, 10).until(EC.new_window_is_opened(original_handles))
            logging.info("Nova janela detectada após clicar em 'dados das partes'.")
        except TimeoutException:
            logging.error("Timeout ao esperar pela nova janela de 'dados das partes'.")
            driver.save_screenshot("getDataParties_new_window_timeout.png")
            raise

        # Obter os novos handles das janelas
        new_handles = set(driver.window_handles) - original_handles
        if new_handles:
            data_window = new_handles.pop()
            time.sleep(1)
            driver.switch_to.window(data_window)
            logging.info(f"Alternado para a nova janela: {data_window}")

            # Alternar para o frame correto dentro da nova janela, se necessário
            # Exemplo: se houver um frame chamado 'dataFrame'
            # driver.switch_to.frame('dataFrame')

            # Coletar os dados das partes
            data = collectDataParties()
            data['Process Number'] = process_number  # Adicionar o número do processo aos dados
            process_data_list.append(data)
            logging.info("Dados das partes coletados e adicionados à lista.")

            # Fechar a janela de dados
            driver.close()
            logging.info("Janela de dados das partes fechada com sucesso.")

            # Retornar para a janela do processo
            driver.switch_to.window(process_window_handle)
            logging.info("Retornado para a janela do processo.")

        else:
            logging.warning("Nenhuma nova janela foi aberta após clicar em 'dados das partes'.")

    except TimeoutException as te:
        logging.error(f"TimeoutException: Elementos demoraram para carregar. Erro: {te}")
        driver.save_screenshot("getDataParties_timeout.png")
        raise te
    except Exception as e:
        logging.error(f"Falha em coletar dados das partes. Erro: {e}")
        driver.save_screenshot("getDataParties_exception.png")
        raise e

def save_data_to_excel(data_list, filename="dados_partes.xlsx"):
    """
    Salva a lista de dicionários em um arquivo Excel.
    
    :param data_list: Lista de dicionários contendo os dados a serem salvos.
    :param filename: Nome do arquivo Excel de saída.
    """
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Dados das Partes"

        # Cabeçalhos (incluindo 'Process Number')
        headers = ['Process Number', 'CPF', 'Nome Civil', 'Data de Nascimento', 'Genitor', 'Genitora']
        ws.append(headers)

        # Estilização dos cabeçalhos
        bold_font = Font(bold=True)
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = bold_font

        # Adicionar os dados
        for data in data_list:
            ws.append([
                data.get('Process Number', ''),
                data.get('CPF', ''),
                data.get('Nome Civil', ''),
                data.get('Data de Nascimento', ''),
                data.get('Genitor', ''),
                data.get('Genitora', '')
            ])

        # Ajustar a largura das colunas
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Obtém a letra da coluna
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Salvar o arquivo
        wb.save(filename)
        logging.info(f"Dados salvos com sucesso no arquivo '{filename}'.")

    except Exception as e:
        driver.save_screenshot("save_data_to_excel_exception.png")
        logging.error(f"Ocorreu uma exceção ao salvar os dados no Excel. Captura de tela salva como 'save_data_to_excel_exception.png'. Erro: {e}")
        raise e

def switch_to_ngFrame():
    """
    Alterna o contexto do Selenium para o frame 'ngFrame'.
    """
    try:
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
        logging.info("Alternado para o frame 'ngFrame'.")
    except TimeoutException:
        logging.error("Timeout ao tentar alternar para o frame 'ngFrame'.")
        driver.save_screenshot("switch_to_ngFrame_timeout.png")
        raise

def InfoPartiesProcessOnTagSearch():
    try:
        original_window = driver.current_window_handle  # Salva o handle da janela original

        # Obter o número total de processos
        switch_to_ngFrame()
        process_list = get_process_list()
        total_processes = len(process_list)
        logging.info(f"Total de processos a serem processados: {total_processes}")

        for index in range(1, total_processes + 1):  # Ajustar índice para começar em 1
            logging.info(f"Iniciando o processamento do processo {index} de {total_processes}")

            # Garantir que estamos no frame 'ngFrame' antes de cada iteração
            switch_to_ngFrame()

            # Gerar o XPath relativo para localizar o processo
            process_xpath = f"(//processo-datalist-card)[{index}]//a/div/span[2]"
            logging.info(f"XPath gerado: {process_xpath}")

            try:
                # Localizar o elemento do processo
                process_element = wait.until(EC.element_to_be_clickable((By.XPATH, process_xpath)))
            except TimeoutException:
                logging.error(f"Timeout ao localizar o elemento do processo no índice {index} com XPath: {process_xpath}")
                driver.save_screenshot(f"process_element_{index}_timeout.png")
                continue  # Pular para o próximo processo

            # Extrair o número do processo
            raw_process_number = process_element.text.strip()
            process_number = re.sub(r'\D', '', raw_process_number)
            if len(process_number) >= 17:
                process_number = f"{process_number[:7]}-{process_number[7:9]}.{process_number[9:13]}.{process_number[13]}.{process_number[14:16]}.{process_number[16:]}"
            else:
                process_number = raw_process_number  # Fallback caso o formato esperado não seja encontrado

            logging.info(f"Número do Processo: {process_number}")
            print(process_number)

            # Clicar no elemento do processo e obter o handle da nova janela
            try:
                process_window_handle = click_on_process(process_element)
            except Exception as e:
                logging.error(f"Falha ao clicar no processo no índice {index}: {e}")
                driver.save_screenshot(f"click_process_{index}_exception.png")
                continue  # Pular para o próximo processo

            # Navegar para a página de dados das partes e coletar dados
            try:
                original_handles = set(driver.window_handles)
                getDataParties(original_handles=original_handles.copy(), process_number=process_number, process_window_handle=process_window_handle)
            except Exception as e:
                logging.error(f"Falha ao coletar dados para o processo {process_number}: {e}")
                driver.save_screenshot(f"getDataParties_{process_number}_exception.png")

            # Fechar a janela do processo
            try:
                driver.close()
                logging.info("Janela do processo fechada com sucesso.")
            except Exception as e:
                logging.error(f"Falha ao fechar a janela do processo {process_number}: {e}")

            # Retornar para a janela original
            try:
                driver.switch_to.window(original_window)
                logging.info("Retornado para a janela original.")
            except Exception as e:
                logging.error(f"Falha ao retornar para a janela original: {e}")

            # Esperar antes de processar o próximo processo
            time.sleep(1)  # Ajuste conforme necessário

        logging.info("Processamento concluído.")

        # Salvar os dados coletados em um arquivo Excel após o processamento
        if process_data_list:
            save_data_to_excel(process_data_list)

    except Exception as e:
        driver.save_screenshot("InfoPartiesProcessOnTagSearch_exception.png")
        logging.error(f"Ocorreu uma exceção em 'InfoPartiesProcessOnTagSearch'. Captura de tela salva como 'InfoPartiesProcessOnTagSearch_exception.png'. Erro: {e}")
        raise e

def main():
    load_dotenv()
    initialize_driver()
    try:
        user, password = os.getenv("USER"), os.getenv("PASSWORD")
        login(user, password)
        profile = os.getenv("PROFILE")
        select_profile(profile)

        search_on_tag("Possivel OBT")
        InfoPartiesProcessOnTagSearch()
        time.sleep(5)
    finally:
        driver.quit()
        logging.info("Driver encerrado.")

if __name__ == "__main__":
    main()
