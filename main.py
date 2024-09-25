from selenium import webdriver
from selenium.webdriver.support.expected_conditions import StaleElementReferenceException
from selenium.common.exceptions import StaleElementReferenceException, ElementClickInterceptedException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from openpyxl import Workbook
from openpyxl.styles import Font
import time
from dotenv import load_dotenv
import os
from functools import wraps

# Variáveis globais para driver e wait
driver = None
wait = None

def switch_to_new_window(original_handles, timeout=10):
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
        driver.save_screenshot("switch_to_new_window_timeout.png")
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
        driver.save_screenshot("switch_to_original_window_exception.png")
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

    # Diretório para onde os downloads serão salvos
    download_directory = os.path.abspath("C:/Users/lfmdsantos/Downloads/processosBaixados")  # Altere "downloads" para o caminho desejado

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
    click_element("/html/body/app-root/selector/div/div/div[2]/right-panel/div/etiquetas/div[1]/div/div[2]/ul/p-datalist/div/div/ul/li/div/li/div[2]/span/span")
    print(f"Pesquisa realizada com o texto: {search_text}")

@retry()
def search_on_tag(search):
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
    original_handles = set(driver.window_handles)
    print(f"Handles originais das janelas: {original_handles}")    
    nav_tag()
    input_tag(search)

def nav_tag():
    xpath = "/html/body/app-root/selector/div/div/div[1]/side-bar/nav/ul/li[5]/a/i"
    click_element(xpath)

def click_element_on_result():
    try:
        print("Tentando localizar o elemento...")
        target_xpath = "/html/body/app-root/selector/div/div/div[2]/right-panel/div/etiquetas/div[1]/div/div[2]/ul/p-datalist/div/div/ul/li/div/li/div[2]/span/span"
        click_element(target_xpath)
        print("Elemento clicado com sucesso!")
        
    except TimeoutException as e:
        driver.save_screenshot("timeout_exception.png")
        print("TimeoutException: Elemento não encontrado. Captura de tela salva como 'timeout_exception.png'")
        raise e

@retry()
def click_processo_numero():
    try:
        print("Tentando localizar o elemento de processo número...")
        processo_numero_selector = "#divListaProcessos > div:nth-child(3) > processo-datalist-card > div > div.col-sm-11.no-padding.pt-5 > a > div > span.tarefa-numero-processo.process"

        target_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, processo_numero_selector)))
        print("Elemento é clicável. Clicando no elemento...")

        original_handles = set(driver.window_handles)
        print(f"Handles originais das janelas: {original_handles}")

        target_element.click()
        print("Elemento clicado com sucesso!")

        new_window_handle = switch_to_new_window(original_handles)

    except Exception as e:
        driver.save_screenshot("click_processo_numero_exception.png")
        print(f"Exception: Ocorreu um erro ao tentar clicar no elemento de processo número. Captura de tela salva como 'click_processo_numero_exception.png'. Erro: {e}")
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
        driver.save_screenshot("select_tipo_documento_exception.png")
        print(f"Ocorreu uma exceção ao selecionar o tipo de documento. Captura de tela salva como 'select_tipo_documento_exception.png'. Erro: {e}")
        raise e

@retry()
def scroll_div_to_bottom():
    try:
        print("Localizando a div 'divTimeLine:eventosTimeLineElement' na nova janela...")
        timeline_div = wait.until(EC.presence_of_element_located((By.ID, 'divTimeLine:eventosTimeLineElement')))
        print("Div 'divTimeLine:eventosTimeLineElement' localizada.")
    
        last_height = driver.execute_script("return arguments[0].scrollTop;", timeline_div)
        print(f"Altura inicial de rolagem: {last_height}")
        max_attempts = 10
        attempts = 0
    
        while attempts < max_attempts:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", timeline_div)
            print(f"Rolagem tentativa {attempts + 1} realizada.")
    
            time.sleep(2)
    
            new_height = driver.execute_script("return arguments[0].scrollTop;", timeline_div)
            print(f"Nova altura de rolagem: {new_height}")
    
            if new_height == last_height:
                print("Chegamos ao final da div. Não há mais conteúdo para carregar.")
                break
            last_height = new_height
            attempts += 1
        else:
            print("Número máximo de tentativas de rolagem atingido.")
    
    except TimeoutException as e:
        driver.save_screenshot("scroll_div_timeout.png")
        print("TimeoutException: Não foi possível localizar ou rolar a div. Captura de tela salva como 'scroll_div_timeout.png'")
        raise e
    except Exception as e:
        driver.save_screenshot("scroll_div_exception.png")
        print(f"Exception: Ocorreu um erro ao tentar rolar a div. Captura de tela salva como 'scroll_div_exception.png'. Erro: {e}")
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
        driver.save_screenshot("click_element_exception.png")
        print(f"Ocorreu uma exceção ao clicar no elemento. Captura de tela salva como 'click_element_exception.png'. Erro: {e}")
        raise e

def main():
    load_dotenv()
    initialize_driver()
    try:
        user, password = os.getenv("USER"), os.getenv("PASSWORD")
        login(user, password)
        profile = os.getenv("PROFILE")
        select_profile(profile)

        search_on_tag("Cobrar Custas")
        click_element_on_result()
        click_processo_numero()

        # Clicar no botão de download usando a função genérica
        click_element("//a[@title='Download autos do processo']")

        # Selecionar o tipo de documento 'Despacho'
        select_tipo_documento("Despacho")

        # Clicar no botão de download input usando a função genérica
        click_element("/html/body/div/div[1]/div/form/span/ul[2]/li[5]/div/div[5]/input")

        time.sleep(10)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
