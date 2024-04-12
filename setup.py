"automation"
from openpyxl import Workbook
from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import login

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)
login_url = 'https://pje.tjba.jus.br/pje/login.seam'
driver.get(login_url)

## DEF LOGIN
def login(user,password):
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ssoFrame')))
    username_input = wait.until(EC.presence_of_element_located((By.ID, 'username')))
    password_input = wait.until(EC.presence_of_element_located((By.ID, 'password')))
    login_button = wait.until(EC.presence_of_element_located((By.ID, 'kc-login')))
    username_input.send_keys(user)
    password_input.send_keys(password)
    login_button.click()
    driver.switch_to.default_content()


login('21920907572','L16@28m18')


##TOKEN
proceed_button = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Prosseguir sem o Token')]"))
)
proceed_button.click()

## DEF SELECIONAR PERFIL
def selectProfile(profile):
    dropdown=wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'dropdown-toggle')))
    dropdown.click()
    button_xpath = f"//a[contains(text(), '{profile}')]"
    desired_button = wait.until(
    EC.element_to_be_clickable((By.XPATH, button_xpath))
    )
    desired_button.click()


selectProfile('V DOS FEITOS DE REL DE CONS CIV E COMERCIAIS DE RIO REAL / Direção de Secretaria / Diretor de Secretaria')
## Clicando na Area de pesquisa Completa
wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))

icon_search_button = wait.until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'li#liConsultaProcessual i.fas.fa-search'))
)

icon_search_button.click()

## Pesquisando 
wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameConsultaProcessual')))

#consultaAdvogado=wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id168:nomeAdvogado')))
def queryjudicialClass(inputClass):
    consultaClasse=wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id245:classeJudicial')))
    consultaClasse.send_keys(inputClass)

def namePart(inputPartie):
    nomeDaParte=wait.until(EC.presence_of_element_located((By.ID, 'fPP:j_id150:nomeParte')))
    nomeDaParte.send_keys(inputPartie)

def orgaoJudicial():
    consultaNumeroOrgaoJustica=wait.until(EC.presence_of_element_located((By.ID, 'fPP:numeroProcesso:NumeroOrgaoJustica')))
    consultaNumeroOrgaoJustica.send_keys('0216')
    
queryjudicialClass('EXECUÇÃO FISCAL')
namePart('MUNICIPIO DE RIO REAL BAHIA')

btnProcurarProcesso=wait.until(EC.presence_of_element_located((By.ID, 'fPP:searchProcessos')))
btnProcurarProcesso.click()

## Resultado da pesquisa
WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, 'fPP:processosTable:tb')))
numeros_dos_processos = []

# Use um loop para navegar através das páginas de processos
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
        numeros_dos_processos.append(numero_do_processo)

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
        print(f"Erro ao clicar no botão de próxima página: {e}")
        break   

# Imprime todos os números dos processos coletados

numeros_unicos = set(numeros_dos_processos)

# Converter o conjunto de volta para uma lista, se necessário
numeros_unicos_lista = list(numeros_unicos)

# Imprimir os números dos processos sem duplicações
for numero in numeros_unicos_lista:
    print(numero)


driver.quit()



# Cria um novo arquivo Excel (workbook) e seleciona a planilha ativa
wb = Workbook()
ws = wb.active

# Define o título da planilha
ws.title = "Processos"

# Insere cada número do processo em uma nova linha na primeira coluna (A)
for row, processo in enumerate(numeros_unicos_lista, start=1):
    ws[f"A{row}"] = processo

# Salva o arquivo Excel
nome_arquivo = "./docs/numeros_dos_processos.xlsx"
wb.save(filename=nome_arquivo)

print(f"Arquivo '{nome_arquivo}' criado com sucesso.")