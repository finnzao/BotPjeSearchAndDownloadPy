from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import os
import json

def initialize_driver():
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 20)
    return driver, wait

def login(driver, wait, user, password):
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

def skip_token(driver, wait):
    proceed_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Prosseguir sem o Token')]"))
    )
    proceed_button.click()

def select_profile(driver, wait, profile):
    dropdown = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'dropdown-toggle')))
    dropdown.click()
    button_xpath = f"//a[contains(text(), '{profile}')]"
    desired_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, button_xpath))
    )
    desired_button.click()

def search_process(driver, wait, classeJudicial='', nomeParte='', numOrgaoJustica='0216', numeroOAB='', estadoOAB=''):
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
    icon_search_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'li#liConsultaProcessual i.fas'))
    ) 
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
        listaEstadosOAB.select_by_value(estadoOAB) # 4 .:. BA
    
    consulta_classe = wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id245:classeJudicial')))
    consulta_classe.send_keys(classeJudicial)
    
    ElementonomeDaParte = wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id150:nomeParte')))
    ElementonomeDaParte.send_keys(nomeParte)
    
    btnProcurarProcesso = wait.until(EC.presence_of_element_located((By.ID, 'fPP:searchProcessos')))
    btnProcurarProcesso.click()

def collect_process_numbers(driver, wait):
    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, 'fPP:processosTable:tb')))
    numProcessos = []
    while True:
        # Aguarde até que os links dos processos estejam visíveis
        WebDriverWait(driver, 50).until(
            EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "a.btn-link.btn-condensed"))
        )

        # Encontre todos os links dos processos na página atual
        links_dos_processos = driver.find_elements(By.CSS_SELECTOR, "a.btn-link.btn-condensed")

        # Extraia o número do processo de cada link e adicione à lista
        for link in links_dos_processos:
            numero_do_processo = link.get_attribute('title')
            numProcessos.append(numero_do_processo)

        # Tenta encontrar o botão de próxima página e clicar nele
        try:    
            # Aguardar até que o elemento que está interceptando o clique desapareça
            wait.until(
                EC.invisibility_of_element((By.ID, 'j_id136:modalStatusCDiv'))
            )
            
            # Depois que o elemento desaparecer, tente clicar no botão de próxima página novamente
            next_page_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//td[contains(@onclick, 'next')]"))
            )
            next_page_button.click()
        except Exception as e:
            print("Fim da pagina")
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
    driver, wait = initialize_driver()
    user, password = os.getenv("USER"), os.getenv("PASSWORD")
    profile = os.getenv("PROFILE")
    optionSearch = {'classeJudicial': "Curatela", 'nomeParte': ''}
    login(driver, wait, user, password)
    #skip_token(driver, wait)
    #select_profile(driver, wait, profile)
    
    search_process(driver, wait, optionSearch['classeJudicial'], optionSearch['nomeParte'])
    process_numbers = collect_process_numbers(driver, wait)
    driver.quit()
    
    # Salvar os números dos processos em formato JSON
    save_to_json(process_numbers)

if __name__ == "__main__":
    main()