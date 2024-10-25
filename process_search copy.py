from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException
from dotenv import load_dotenv
import os
import json
import math
import re
import time
import logging
from dataclasses import dataclass
from typing import List, Optional, Dict

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pje_automacao.log'),
        logging.StreamHandler()
    ]
)

@dataclass
class OpcoesBusca:
    """Classe para armazenar as opções de busca do processo"""
    nomeParte: str
    numOrgaoJustica: str
    assunto: str = ''
    nomeRepresentante: str = ''
    alcunha: str = ''
    classeJudicial: str = ''
    numDocumento: str = ''
    estadoOAB: str = ''
    numeroOAB: str = ''

class AutomacaoPJE:
    def __init__(self, modo_background: bool = False):
        self.navegador = None
        self.espera = None
        self.modo_background = modo_background

    def inicializar_navegador(self) -> None:
        """Inicializa o navegador Chrome com as configurações necessárias"""
        try:
            opcoes_chrome = Options()
            if self.modo_background:
                opcoes_chrome.add_argument('--headless')
            opcoes_chrome.add_argument('--start-maximized')
            opcoes_chrome.add_argument('--disable-gpu')
            opcoes_chrome.add_argument('--no-sandbox')
            opcoes_chrome.add_argument('--disable-dev-shm-usage')
            
            self.navegador = webdriver.Chrome(options=opcoes_chrome)
            self.espera = WebDriverWait(self.navegador, 20)
            logging.info("Navegador inicializado com sucesso")
        except WebDriverException as e:
            logging.error(f"Erro ao inicializar navegador: {e}")
            raise

    def fazer_login(self, usuario: str, senha: str) -> None:
        """Realiza o login no sistema PJE"""
        try:
            url_login = 'https://pje.tjba.jus.br/pje/login.seam'
            self.navegador.get(url_login)
            self.espera.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ssoFrame')))
            
            campo_usuario = self.espera.until(EC.presence_of_element_located((By.ID, 'username')))
            campo_senha = self.espera.until(EC.presence_of_element_located((By.ID, 'password')))
            botao_login = self.espera.until(EC.presence_of_element_located((By.ID, 'kc-login')))
            
            campo_usuario.send_keys(usuario)
            campo_senha.send_keys(senha)
            botao_login.click()
            self.navegador.switch_to.default_content()
            logging.info("Login realizado com sucesso")
        except Exception as e:
            logging.error(f"Erro durante o login: {e}")
            raise

    def pular_token(self) -> None:
        """Pula a etapa de verificação por token"""
        try:
            botao_prosseguir = self.espera.until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Prosseguir sem o Token')]"))
            )
            botao_prosseguir.click()
            logging.info("Token ignorado com sucesso")
        except TimeoutException:
            logging.warning("Botão de pular token não encontrado")

    def selecionar_perfil(self, perfil: str) -> None:
        """Seleciona o perfil do usuário no sistema"""
        try:
            menu_dropdown = self.espera.until(EC.presence_of_element_located((By.CLASS_NAME, 'dropdown-toggle')))
            menu_dropdown.click()
            
            xpath_botao = f"//a[contains(text(), '{perfil}')]"
            botao_perfil = self.espera.until(EC.element_to_be_clickable((By.XPATH, xpath_botao)))
            self.navegador.execute_script("arguments[0].scrollIntoView(true);", botao_perfil)
            self.navegador.execute_script("arguments[0].click();", botao_perfil)
            logging.info(f"Perfil '{perfil}' selecionado com sucesso")
        except Exception as e:
            logging.error(f"Erro ao selecionar perfil: {e}")
            raise

    def buscar_processo(self, opcoes_busca: OpcoesBusca) -> None:
        """Realiza a busca de processos com as opções fornecidas"""
        try:
            self.espera.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'ngFrame')))
            botao_busca = self.espera.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'li#liConsultaProcessual i.fas'))
            )
            botao_busca.click()
            
            self.espera.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'frameConsultaProcessual')))
            
            # Preencher campos de busca
            self._preencher_campo_busca('fPP:numeroProcesso:NumeroOrgaoJustica', opcoes_busca.numOrgaoJustica)
            self._preencher_campo_busca('fPP:j_id150:nomeParte', opcoes_busca.nomeParte)
            self._preencher_campo_busca('fPP:j_id245:classeJudicial', opcoes_busca.classeJudicial)
            
            if opcoes_busca.estadoOAB and opcoes_busca.numeroOAB:
                self._preencher_campos_oab(opcoes_busca.numeroOAB, opcoes_busca.estadoOAB)
            
            botao_pesquisar = self.espera.until(EC.presence_of_element_located((By.ID, 'fPP:searchProcessos')))
            botao_pesquisar.click()
            logging.info("Busca iniciada com sucesso")
        except Exception as e:
            logging.error(f"Erro durante a busca: {e}")
            raise

    def _preencher_campo_busca(self, id_campo: str, valor: str) -> None:
        """Auxiliar para preencher campos de busca"""
        if valor:
            elemento = self.espera.until(EC.presence_of_element_located((By.ID, id_campo)))
            elemento.send_keys(valor)

    def _preencher_campos_oab(self, numero_oab: str, estado_oab: str) -> None:
        """Auxiliar para preencher campos específicos de OAB"""
        campo_numero = self.espera.until(EC.presence_of_element_located((By.ID, 'fPP:decorationDados:numeroOAB')))
        campo_numero.send_keys(numero_oab)
        
        campo_estado = self.espera.until(EC.presence_of_element_located((By.ID, 'fPP:decorationDados:ufOABCombo')))
        Select(campo_estado).select_by_value(estado_oab)

    def coletar_numeros_processo(self) -> List[str]:
        """Coleta os números dos processos de todas as páginas"""
        numeros_processo = set()
        numero_pagina = 1
        
        while True:
            try:
                self.espera.until(EC.presence_of_element_located((By.ID, 'fPP:processosTable:tb')))
                links = self.espera.until(
                    EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "a.btn-link.btn-condensed"))
                )
                
                for link in links:
                    numeros_processo.add(link.get_attribute('title'))
                
                logging.info(f"Página {numero_pagina}: {len(links)} processos encontrados")
                
                if len(links) < 20 or not self._ir_proxima_pagina():
                    break
                    
                numero_pagina += 1
                
            except Exception as e:
                logging.error(f"Erro ao coletar processos da página {numero_pagina}: {e}")
                break
        
        return list(numeros_processo)

    def _ir_proxima_pagina(self) -> bool:
        """Tenta navegar para a próxima página de resultados"""
        try:
            self.espera.until(EC.invisibility_of_element((By.ID, 'j_id136:modalStatusCDiv')))
            botao_proximo = self.espera.until(
                EC.element_to_be_clickable((By.XPATH, "//td[contains(@onclick, 'next')]"))
            )
            botao_proximo.click()
            time.sleep(2)  # Pausa para carregamento
            return True
        except TimeoutException:
            return False

    @staticmethod
    def salvar_json(dados: List[str], nome_arquivo: str = "ResultadoProcessos.json") -> None:
        """Salva os resultados em um arquivo JSON"""
        try:
            pasta_docs = "./docs"
            os.makedirs(pasta_docs, exist_ok=True)
            
            caminho_arquivo = os.path.join(pasta_docs, nome_arquivo)
            with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo_json:
                json.dump(dados, arquivo_json, ensure_ascii=False, indent=4)
            
            logging.info(f"Resultados salvos em {caminho_arquivo}")
        except Exception as e:
            logging.error(f"Erro ao salvar arquivo JSON: {e}")
            raise

    def fechar(self) -> None:
        """Fecha o navegador"""
        if self.navegador:
            self.navegador.quit()
            logging.info("Navegador fechado com sucesso")

def limpar_cache(self):
    """Limpa cache, cookies e dados de sessão do navegador"""
    try:
        self.navegador.execute_cdp_cmd('Network.clearBrowserCache', {})
        self.navegador.execute_cdp_cmd('Network.clearBrowserCookies', {})
        self.navegador.delete_all_cookies()
        logging.info("Cache do navegador limpo com sucesso")
    except Exception as e:
        logging.error(f"Erro ao limpar cache: {e}")
        raise
    
def main():
    load_dotenv()
    
    # Configuração das opções de busca
    opcoes_busca = OpcoesBusca(
        nomeParte='EDMILSON DO NASCIMENTO SANTOS',
        numOrgaoJustica='0216'
    )
    
    automacao = AutomacaoPJE(modo_background=True)
    
    try:
        limpar_cache()
        automacao.inicializar_navegador()
        automacao.fazer_login(os.getenv("USUARIO"), os.getenv("SENHA"))
        automacao.selecionar_perfil(os.getenv("PERFIL"))
        automacao.buscar_processo(opcoes_busca)
        
        time.sleep(5)  # Aguarda carregamento inicial dos resultados
        numeros_processo = automacao.coletar_numeros_processo()
        
        if numeros_processo:
            automacao.salvar_json(numeros_processo)
        else:
            logging.warning("Nenhum processo encontrado")
            
    except Exception as e:
        logging.error(f"Erro durante a execução: {e}")
    finally:
        automacao.fechar()

if __name__ == "__main__":
    main()