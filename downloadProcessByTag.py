import re 
from selenium import webdriver
from selenium.webdriver.support.expected_conditions import StaleElementReferenceException
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
from dotenv import load_dotenv
import os
from functools import wraps

# Variáveis globais para driver e wait
driver = None
wait = None

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
            print(f"Alternado para a nova janela: {new_window}")
            return new_window
        else:
            raise TimeoutException("Nova janela não foi encontrada dentro do tempo especificado.")
    except TimeoutException as e:
        save_exception_screenshot("switch_to_new_window_timeout.png")
        print("TimeoutException: Não foi possível encontrar a nova janela. Captura de tela salva como 'switch_to_new_window_timeout.png'")
        raise e

def switch_to_original_window(original_handle):
    """
    Alterna de volta para a janela original.

    :param original_handle: Handle da janela original.
    """
    try:
        driver.switch_to.window(original_handle)
        print(f"Retornado para a janela original: {original_handle}")
    except Exception as e:
        save_exception_screenshot("switch_to_original_window_exception.png")
        print(f"Exception: Ocorreu um erro ao retornar para a janela original. Captura de tela salva como 'switch_to_original_window_exception.png'. Erro: {e}")
        raise e

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
                    print(f"Tentativa {retries} falhou com erro: {e}. Tentando novamente...")
                    if retries >= max_retries:
                        raise TimeoutException(f"Falha ao executar {func.__name__} após {max_retries} tentativas")
        return wrapper
    return decorator

# Inicialização do driver e wait
def initialize_driver():
    global driver, wait
    # Configurações do Chrome
    chrome_options = webdriver.ChromeOptions()

    # Diretório para onde os downloads serão salvos de forma genérica
    user_home = os.path.expanduser("~")
    download_directory = os.path.join(user_home, "Downloads", "processosBaixados")

    # Criar o diretório se não existir
    os.makedirs(download_directory, exist_ok=True)
    print(f"Diretório de download configurado para: {download_directory}")

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
    wait = WebDriverWait(driver, 50)
def save_exception_screenshot(filename):
    """
    Salva o screenshot atual do driver na pasta '.logs/exception'.
    
    :param filename: Nome do arquivo (ex: "meuarquivo.png")
    """
    # Define o diretório de destino
    directory = ".logs/exception"
    # Cria o diretório se não existir
    if not os.path.exists(directory):
        os.makedirs(directory)
    # Gera o caminho completo do arquivo
    filepath = os.path.join(directory, filename)
    save_exception_screenshot(filepath)
    print(f"Screenshot salvo em: {filepath}")

@retry()
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

@retry()
def skip_token():
    proceed_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Prosseguir sem o Token')]"))
    )
    proceed_button.click()

@retry()
def select_profile(profile):
    dropdown = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'dropdown-toggle')))
    dropdown.click()
    button_xpath = f"//a[contains(text(), '{profile}')]"
    desired_button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
    driver.execute_script("arguments[0].scrollIntoView(true);", desired_button)
    driver.execute_script("arguments[0].click();", desired_button)

@retry()
def search_process(classeJudicial='', nomeParte='', numOrgaoJustica='0216', numeroOAB='', estadoOAB=''):
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
    icon_search_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'li#liConsultaProcessual i.fas')))
    icon_search_button.click()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameConsultaProcessual')))
    ElementoNumOrgaoJutica = wait.until(EC.presence_of_element_located((By.ID, 'fPP:numeroProcesso:NumeroOrgaoJustica')))
    ElementoNumOrgaoJutica.send_keys(numOrgaoJustica)
   
    # OAB
    if estadoOAB:
        ElementoNumeroOAB = wait.until(EC.presence_of_element_located((By.ID, 'fPP:decorationDados:numeroOAB')))
        ElementoNumeroOAB.send_keys(numeroOAB)
        ElementoEstadosOAB = wait.until(EC.presence_of_element_located((By.ID, 'fPP:decorationDados:ufOABCombo')))
        listaEstadosOAB = Select(ElementoEstadosOAB)
        listaEstadosOAB.select_by_value(estadoOAB)
    
    consulta_classe = wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id245:classeJudicial')))
    consulta_classe.send_keys(classeJudicial)
    
    ElementonomeDaParte = wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id150:nomeParte')))
    ElementonomeDaParte.send_keys(nomeParte)
    
    btnProcurarProcesso = wait.until(EC.presence_of_element_located((By.ID, 'fPP:searchProcessos')))
    btnProcurarProcesso.click()

@retry()
def preencher_formulario(numProcesso=None, Comp=None, Etiqueta=None):
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.CLASS_NAME, 'ng-frame')))
    """
    Preenche o formulário com base nos parâmetros opcionais fornecidos e clica no botão Pesquisar.
    
    Parâmetros:
    - numProcesso: O número do processo (opcional)
    - Comp: A competência (opcional)
    - Etiqueta: A etiqueta (opcional)
    """
    
    # Preencher campo "Número do processo" se fornecido
    if numProcesso:
        num_processo_input = wait.until(EC.presence_of_element_located((By.ID, "itNrProcesso")))
        driver.execute_script("arguments[0].scrollIntoView(true);", num_processo_input)
        num_processo_input.clear()
        num_processo_input.send_keys(numProcesso)
    
    # Preencher campo "Competência" se fornecido
    if Comp:
        competencia_input = wait.until(EC.presence_of_element_located((By.ID, "itCompetencia")))
        driver.execute_script("arguments[0].scrollIntoView(true);", competencia_input)
        competencia_input.clear()
        competencia_input.send_keys(Comp)
    
    # Preencher campo "Etiqueta" se fornecido
    if Etiqueta:
        etiqueta_input = wait.until(EC.presence_of_element_located((By.ID, "itEtiqueta")))
        driver.execute_script("arguments[0].scrollIntoView(true);", etiqueta_input)
        etiqueta_input.clear()
        etiqueta_input.send_keys(Etiqueta)
    
    # Clicar no botão de "Pesquisar"
    pesquisar_button_xpath = "//button[text()='Pesquisar']"
    click_element(pesquisar_button_xpath)

    print("Formulário preenchido e pesquisa iniciada com sucesso!")
    time.sleep(10)

def input_tag(search_text):
    search_input = wait.until(EC.element_to_be_clickable((By.ID, "itPesquisarEtiquetas")))
    search_input.clear()
    search_input.send_keys(search_text)
    click_element("/html/body/app-root/selector/div/div/div[2]/right-panel/div/etiquetas/div[1]/div/div[1]/div[2]/div[1]/span/button[1]")
    time.sleep(1)
    print(f"Pesquisa realizada com o texto: {search_text}")
    click_element("/html/body/app-root/selector/div/div/div[2]/right-panel/div/etiquetas/div[1]/div/div[2]/ul/p-datalist/div/div/ul/li/div/li/div[2]/span/span")

@retry()
def search_on_tag(search):
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
    original_handles = set(driver.window_handles)
    print(f"Handles originais das janelas: {original_handles}")    
    nav_tag()
    input_tag(search)

def nav_tag():
    xpath = "/html/body/app-root/selector/div/div/div[1]/side-bar/nav/ul/li[5]/a"
    click_element(xpath)

def get_process_list():
    """
    Retorna uma lista de elementos que representam os processos encontrados.
    """
    try:
        # XPath para localizar todos os itens da lista de processos
        process_xpath = "//processo-datalist-card"
        processes = wait.until(EC.presence_of_all_elements_located((By.XPATH, process_xpath)))
        print(f"Número de processos encontrados: {len(processes)}")
        return processes
    except Exception as e:
        save_exception_screenshot("get_process_list_exception.png")
        print(f"Ocorreu uma exceção ao obter a lista de processos. Erro: {e}")
        raise e

def click_on_process(process_element):
    try:
        original_handles = set(driver.window_handles)
        driver.execute_script("arguments[0].scrollIntoView(true);", process_element)
        driver.execute_script("arguments[0].click();", process_element)
        print("Processo clicado com sucesso!")
        new_window_handle = switch_to_new_window(original_handles)
    except Exception as e:
        save_exception_screenshot("click_on_process_exception.png")
        print(f"Ocorreu uma exceção ao clicar no processo. Erro: {e}")
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
            print(f"Elemento clicado com sucesso: {xpath}")
        except (ElementClickInterceptedException, Exception) as e:
            print(f"Erro ao clicar no elemento normalmente: {e}. Tentando com JavaScript...")
            driver.execute_script("arguments[0].click();", element)
            print(f"Elemento clicado com sucesso usando JavaScript: {xpath}")
    except Exception as e:
        save_exception_screenshot("click_element_exception.png")
        print(f"Ocorreu uma exceção ao clicar no elemento. Captura de tela salva como 'click_element_exception.png'. Erro: {e}")
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
        print(f"Tipo de documento '{tipoDocumento}' selecionado com sucesso.")
    except Exception as e:
        save_exception_screenshot("select_tipo_documento_exception.png")
        print(f"Ocorreu uma exceção ao selecionar o tipo de documento. Captura de tela salva como 'select_tipo_documento_exception.png'. Erro: {e}")
        raise e

def downloadProcessOnTagSearch(typeDocument):
    try:
        original_window = driver.current_window_handle  # Salva o handle da janela original

        # Certificar-se de que estamos dentro do frame 'ngFrame'
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
        print("Dentro do frame 'ngFrame'.")

        process_numbers = []  # Lista para armazenar os números de processo

        # Obter o número total de processos
        total_processes = len(get_process_list())

        for index in range(1, total_processes + 1):
            print(f"\nIniciando o download para o processo {index} de {total_processes}")

            # XPath relativo para localizar o processo
            process_xpath = f"(//processo-datalist-card)[{index}]//a/div/span[2]"
            print(f"XPath gerado: {process_xpath}")

            # Localizar o elemento antes de passar para a função
            process_element = wait.until(EC.element_to_be_clickable((By.XPATH, process_xpath)))
            raw_process_number = process_element.text.strip()
            process_number = re.sub(r'\D', '', raw_process_number)
            if len(process_number) >= 17:
                process_number = f"{process_number[:7]}-{process_number[7:9]}.{process_number[9:13]}.{process_number[13]}.{process_number[14:16]}.{process_number[16:]}"
            else:
                process_number = raw_process_number
            print(f"Número do processo: {process_number}")
            process_numbers.append(process_number)

            click_on_process(process_element)

            driver.switch_to.default_content()
            print("Saiu do frame 'ngFrame'.")

            click_element("//*[@id='navbar:ajaxPanelAlerts']/ul[2]/li[5]/a")
            select_tipo_documento(typeDocument)
            click_element("/html/body/div/div[1]/div/form/span/ul[2]/li[5]/div/div[5]/input")

            # Esperar pelo download ser concluído, se necessário
            time.sleep(5)  # Ajuste conforme necessário

            # Fechar a janela atual
            driver.close()
            print("Janela atual fechada com sucesso.")

            # Alternar de volta para a janela original
            driver.switch_to.window(original_window)
            print("Retornado para a janela original.")

            # Entrar novamente no frame 'ngFrame' para a próxima iteração
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
            print("Alternado para o frame 'ngFrame'.")

        print("Processamento concluído.")
        return process_numbers  # Retorna a lista de números de processo

    except Exception as e:
        save_exception_screenshot("downloadProcessOnTagSearch_exception.png")
        print(f"Ocorreu uma exceção em 'downloadProcessOnTagSearch'. Captura de tela salva. Erro: {e}")
        raise e

def download_requested_processes(process_numbers):
    """
    Navega até a página de requisições de download e baixa os processos cujo número está em 'process_numbers',
    evitando baixar processos com numeração repetida.
    """
    try:
        # Navega até a página de download
        driver.get('https://pje.tjba.jus.br/pje/AreaDeDownload/listView.seam')

        # Espera o iframe 'ngFrame' estar disponível e alterna para ele
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
        print("Dentro do iframe 'ngFrame'.")

        # Espera a tabela carregar dentro do iframe
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        print("Tabela carregada.")

        # Obtém todas as linhas da tabela (exceto o cabeçalho)
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table//tbody//tr")))
        print(f"Número total de processos na lista de downloads: {len(rows)}")

        # Conjunto para rastrear os números de processos já baixados
        downloaded_process_numbers = set()

        for row in rows:
            # Obtém o número do processo na primeira coluna
            process_number_td = row.find_element(By.XPATH, "./td[1]")
            process_number = process_number_td.text.strip()
            print(f"Verificando o processo: {process_number}")
            print(f"Números de processos a serem baixados: {process_numbers}")

            # Verifica se o número do processo está na lista e não foi baixado ainda
            if process_number in process_numbers and process_number not in downloaded_process_numbers:
                print(f"Processo {process_number} encontrado na lista e ainda não baixado. Iniciando download...")
                # Clica no botão de download nesta linha
                # Considerando que o botão está na última coluna
                download_button = row.find_element(By.XPATH, "./td[last()]//button")
                driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
                download_button.click()
                # Adiciona o número do processo ao conjunto de processados
                downloaded_process_numbers.add(process_number)
                # Espera pelo download ser iniciado
                time.sleep(5)
            else:
                print(f"Processo {process_number} não está na lista ou já foi baixado. Pulando...")

        # Volta para o conteúdo principal
        driver.switch_to.default_content()
        print("Voltando para o conteúdo principal.")

    except Exception as e:
        save_exception_screenshot("download_requested_processes_exception.png")
        print(f"Ocorreu uma exceção em 'download_requested_processes'. Captura de tela salva. Erro: {e}")
        raise e

def main():
    load_dotenv()
    initialize_driver()
    try:
        user, password = os.getenv("USER"), os.getenv("PASSWORD")
        login(user, password)
        profile = os.getenv("PROFILE")
        select_profile(profile)
        search_on_tag("etiqueta teste")
        process_numbers = downloadProcessOnTagSearch("Selecione")  # Captura os números de processo
        download_requested_processes(process_numbers)  # Chama a nova função com a lista de números
        time.sleep(10)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
