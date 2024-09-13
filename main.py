from selenium import webdriver
from selenium.webdriver.support.expected_conditions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from openpyxl import Workbook
from openpyxl.styles import Font
import time
from dotenv import load_dotenv
import os
import math
from functools import wraps

# Variáveis globais para driver e wait
driver = None
wait = None

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
    driver = webdriver.Chrome()
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
def search_process(classeJudicial:str='', nomeParte:str='',numOrgaoJustica:int='0216',numeroOAB:int='',estadoOAB:int=''):
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
def click_filtros_button():
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
    target_element = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "div.filtros.vcenter.pa-5[data-toggle='collapse'][data-target='.group-filtro-tarefas-pendentes-tela']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", target_element)
    target_element = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, "div.filtros.vcenter.pa-5[data-toggle='collapse'][data-target='.group-filtro-tarefas-pendentes-tela']")))
    target_element.click()
    print("Elemento 'Filtros' clicado com sucesso!")
    driver.switch_to.default_content()
@retry()
def click_element_on_result():
    target_element = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "span.nome-etiqueta.selecionada[title='Ver lista de processos vinculados a esta etiqueta']")))
    #driver.execute_script("arguments[0].scrollIntoView(true);", target_element)
    target_element = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, "span.nome-etiqueta.selecionada[title='Ver lista de processos vinculados a esta etiqueta']")))
    target_element.click()
    print("Elemento clicado com sucesso!")

@retry()
def nav_tag():
    tag_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'li#liEtiquetas i.fas'))
    ) 
    tag_button.click()


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
        driver.execute_script("arguments[0].scrollIntoView(true);", num_processo_input)  # Rolar até o elemento
        num_processo_input.clear()
        num_processo_input.send_keys(numProcesso)
    
    # Preencher campo "Competência" se fornecido
    if Comp:
        competencia_input = wait.until(EC.presence_of_element_located((By.ID, "itCompetencia")))
        driver.execute_script("arguments[0].scrollIntoView(true);", competencia_input)  # Rolar até o elemento
        competencia_input.clear()
        competencia_input.send_keys(Comp)
    
    # Preencher campo "Etiqueta" se fornecido
    if Etiqueta:
        etiqueta_input = wait.until(EC.presence_of_element_located((By.ID, "itEtiqueta")))
        driver.execute_script("arguments[0].scrollIntoView(true);", etiqueta_input)  # Rolar até o elemento
        etiqueta_input.clear()
        etiqueta_input.send_keys(Etiqueta)
    
    # Clicar no botão de "Pesquisar"
    pesquisar_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Pesquisar']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", pesquisar_button)  # Rolar até o botão
    pesquisar_button.click()

    print("Formulário preenchido e pesquisa iniciada com sucesso!")
    time.sleep(10)

@retry()
def input_tag(search_text):
    search_input = wait.until(EC.element_to_be_clickable((By.ID, "itPesquisarEtiquetas")))
    search_input.clear()
    search_input.send_keys(search_text)
    search_button = driver.find_element(By.CSS_SELECTOR, "button.btn.btn-default.btn-pesquisa[title='Pesquisar']")
    search_button.click()

@retry()
def search_on_tag(search):
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
    nav_tag()
    input_tag(search)
    click_element_on_result()
def main():
    load_dotenv()
    initialize_driver()
    user, password = os.getenv("USER"), os.getenv("PASSWORD")
    login(user, password)
    profile = os.getenv("PROFILE")
    select_profile("V DOS FEITOS DE REL DE CONS CIV E COMERCIAIS DE RIO REAL / Direção de Secretaria / Diretor de Secretaria")
    #click_filtros_button()
    #preencher_formulario(Etiqueta="Cobrar Custas")
    search_on_tag("Cobrar Custas")
    driver.quit()
    #driver.switch_to.default_content()

if __name__ == "__main__":
    main()
