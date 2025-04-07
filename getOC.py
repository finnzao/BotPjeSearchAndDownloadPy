import os
import time
import logging
from urllib.parse import urlparse, parse_qs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv

from utils.pje_automation import PjeConsultaAutomator
class PjeTJBA:
    def __init__(self, driver: webdriver.Chrome):
        self.driver = driver
        self.wait = WebDriverWait(self.driver, 30)
        self.logger = logging.getLogger(self.__class__.__name__)

    def login(self, user, password):
        try:
            login_url = 'https://pje.tjba.jus.br/pje/login.seam'
            self.driver.get(login_url)
            self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ssoFrame')))
            self.wait.until(EC.presence_of_element_located((By.ID, 'username'))).send_keys(user)
            self.wait.until(EC.presence_of_element_located((By.ID, 'password'))).send_keys(password)
            self.wait.until(EC.element_to_be_clickable((By.ID, 'kc-login'))).click()
            self.driver.switch_to.default_content()
            self.logger.info("Login efetuado com sucesso")
        except Exception as e:
            self.logger.exception("Erro no login")
            raise
            
    def abrir_primeiro_processo(self):
        try:
            self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ngFrame")))
            self.logger.info("Acessando iframe 'ngFrame'")
    
            self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "menuItem")))
    
            elementos = self.wait.until(EC.presence_of_all_elements_located((
                By.XPATH, "//div[contains(@class, 'menuItem')]//a[contains(@href, 'lista-processos-tarefa')]"
            )))
    
            for elemento in elementos:
                if elemento.is_displayed() and elemento.is_enabled():
                    elemento.click()
                    self.logger.info("Clicado no primeiro processo")
                    return
    
            self.logger.error("Nenhum processo clicável encontrado")
            raise Exception("Nenhum processo clicável encontrado")
        except Exception as e:
            self.logger.exception("Erro ao abrir o primeiro processo")
            raise
        finally:
            self.driver.switch_to.default_content()

    def abrir_autos_do_processo(self):
        try:
            self.wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ngFrame")))
            self.logger.info("Contexto do iframe 'ngFrame' acessado para abrir autos")
    
            botao_autos = self.wait.until(EC.presence_of_all_elements_located((
                By.XPATH, "//button[@title='Abrir autos']"
            )))
    
            for botao in botao_autos:
                if botao.is_displayed() and botao.is_enabled():
                    self.wait.until(EC.element_to_be_clickable(botao))
                    botao.click()
                    self.logger.info("Botão 'Abrir autos' clicado com sucesso")
                    return
    
            self.logger.error("Nenhum botão 'Abrir autos' visível e clicável encontrado")
            raise Exception("Nenhum botão 'Abrir autos' visível e clicável encontrado")
        except Exception as e:
            self.logger.exception("Erro ao abrir autos do processo")
            raise
        finally:
            self.driver.switch_to.default_content()


    def capturar_url_com_oc(self):
        try:
            time.sleep(3)
            abas = self.driver.window_handles
            if len(abas) > 1:
                self.driver.switch_to.window(abas[-1])
            url_atual = self.driver.current_url
            self.logger.info("URL capturada na nova aba: %s", url_atual)
            return url_atual
        except Exception as e:
            self.logger.exception("Erro ao capturar URL com OC")
            raise

    def extrair_oc_ou_ca(self, url):
        try:
            query = urlparse(url).query
            params = parse_qs(query)
            oc = params.get('oc', [None])[0] or params.get('ca', [None])[0]
            self.logger.info("Token OC/CA extraído: %s", oc)
            return oc
        except Exception as e:
            self.logger.exception("Erro ao extrair token OC/CA")
            raise

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    driver = webdriver.Chrome(options=options)
    load_dotenv()
    user, password = os.getenv("USER"), os.getenv("PASSWORD")
    bot = PjeConsultaAutomator()
    try:
        pje = PjeTJBA(driver)
        pje.login(user, password)
        pje.abrir_primeiro_processo()
        pje.abrir_autos_do_processo()
        url = pje.capturar_url_com_oc()
        oc = pje.extrair_oc_ou_ca(url)
        bot.update_config({"LoginInfo": {"oc": oc}})
        print("Token OC capturado:", oc)
    finally:
        #input("Pressione Enter para fechar o navegador...")
        driver.quit()
