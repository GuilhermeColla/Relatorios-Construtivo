'''
Utilizando PANDAS esse módulo vai formatar os relatórios extraídos do construtivo,
adequando-os para a ferramenta em desenvolvimento.
'''

import re
import pandas as pd
from datetime import datetime
import os
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC, wait
from selenium.webdriver.chrome.service import Service
import time


class Relatórios_Construtivo():
    def __init__(self, caminho_relatorio_gerencial: str, caminho_relatorio_planejamento):
        #TODO: Processar o relatório de Planejamento.
        '''
        Cria um objeto "Relatório".
        '''
        self.relatorio_gerencial = pd.read_csv(caminho_relatorio_gerencial, sep=";")
        self.relatorio_planejamento = pd.read_csv(caminho_relatorio_planejamento, sep=";", index_col="Codificação")
        self.obter_título_dos_documentos()
    

    @staticmethod
    def remover_colunas(relatorio, colunas_a_remover: list) -> None:
        relatorio.drop(columns=colunas_a_remover, inplace=True)
    
    
    @staticmethod
    def remover_estados(relatorio: pd.DataFrame, nome_coluna_estado_atual: str, estados_a_remover: list) -> None:
        '''
        A função remove os estados contidos na lista "estados_a_remover" do dataframe.
        Isto é feito pois, no momento, documentos neste estado não tem utilidade.
        '''
        
        for estado in estados_a_remover:
            relatorio.drop(relatorio[relatorio[nome_coluna_estado_atual] == estado].index, inplace=True)
        
    
    def obter_título_dos_documentos(self) -> None:
        '''
        A função obtém o título (código) dos documentos removendo a extensão do arquivo
        (.doc, .docx, .xls, etc...) e a versão do documento no relatório gerencial.
        '''
        códigos = []
        revisões = []
        for título in self.relatorio_gerencial["Título"]:
            
            match = re.search("(?P<codigo>[A-Z|0-9|-]+)-(?P<versao>[\w]+)", título)

            if match:
                código, revisão = match.groups()

            else:
                código, revisão = None, None

            códigos.append(código)
            revisões.append(revisão)
                
        self.relatorio_gerencial.drop(columns="Título", inplace=True)
        self.relatorio_gerencial.insert(0, "Revisão",revisões)
        self.relatorio_gerencial.insert(0, "Código", códigos)


    def dias_com_agentes(self, coluna_estado_atual: str) -> None:

        #TODO: Os estados de "Alerta" acontecem após o estado normal. Devemos fazer 
        #o calculo de acordo com o normal ou com o alerta?
        '''Adiciona uma coluna com os dias que se passaram desde a mudança de
        estado do documento'''
        
        diferencas = []

        for codigo, _ in self.relatorio_planejamento.iterrows():

            estado_atual = self.relatorio_planejamento.loc[codigo][coluna_estado_atual]
            dia_estado_atual = datetime.strptime(self.relatorio_gerencial.loc[codigo, estado_atual][estado_atual][0], "%d/%m/%Y %H:%M")
            
            match = re.search("(^[0-9]+) day", str((datetime.now() - dia_estado_atual)))

            if match:
                diferencas.append(match.groups()[0])
            else:
                diferencas.append("0")

            
        self.relatorio_planejamento["Dias"] = diferencas


    def selecionar_estados(self, estados: list) -> pd.DataFrame:
        return self.relatorio_planejamento.query("`Estado Workflow` == @estados")
    

    def com_projetista(self) -> pd.DataFrame:
        return self.selecionar_estados([
            "Em Análise",
            "Alerta de Em Análise",
            "Para Revisar - pre-analise",
            "Alerta de Para Revisar - pre-analise",
            "Não Aprovado",
            "Alerta de Não Aprovado",
            "Para Revisão",
            "Alerta para Revisão",
            "Aprovado",
            "Alerta Aprovado",
            "Aprovado com Comentários",
            "Alerta Aprovado com Comentários"
            ])


    def com_CPFL(self) -> pd.DataFrame:
        return self.selecionar_estados([
            "Para Análise CPFL",
            "Alerta de Para Análise CPFL",
            "Liberação CPFL",
        ])

    
    def com_acessadas(self) -> pd.DataFrame:
        return self.selecionar_estados([
            "Para Análise Acessado",
            "Para Análise Acessado_",
        ])


def download_gerencial_planejamento(nome_pasta_empreendimento: str) -> None:
    
    #Carregando os dados de usuário
    load_dotenv()

    #Criando uma janela do Chrome.
    service = Service(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.implicitly_wait(5)
    wait = WebDriverWait(driver, 60)

    #Conectando ao site colaborativo.
    driver.get("https://incorporadora.colaborativo.com/ssf/a/do?p_name=ss_forum&p_action=1&binderId=37&action=view_permalink&entityType=folder&novl_url=1&novl_landing=1?novl_root=1#1634731550743")
    
    #Obtendo as caixas de texto para login, senha e botão de entrar.
    wait.until(EC.element_to_be_clickable((By.ID, "loginOkBtn")))
    login = driver.find_element(By.ID, "j_usernameId")
    senha = driver.find_element(By.ID, "j_passwordId")
    botao_entrar = driver.find_element(By.ID, "loginOkBtn")

    #Preenchendo as caixas com dados de login, senha e pressionando entrar.
    login.send_keys(os.getenv("LOGIN"))
    senha.send_keys(os.getenv("SENHA"))
    botao_entrar.click()
    time.sleep(5)

    #A partir daqui já estamos conectados na área de trabalho do colaborativo.
    #Navegando os menus.
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[class='gwt-InlineLabel vibe-dataTableEntry-title']")))
    CPFL_Geracao = driver.find_element(By.CSS_SELECTOR, "span[class='gwt-InlineLabel vibe-dataTableEntry-title']")
    CPFL_Geracao.click()
    time.sleep(5)
    projetos = driver.find_elements(By.CSS_SELECTOR, "span[class='gwt-InlineLabel childBindersWidget_ListOfFoldersPanel_folderLabel']")
    #time.sleep(3)
    for element in projetos:
        if nome_pasta_empreendimento in element.text:
            element.click()
            break
            
    time.sleep(5)
    frame = driver.find_element(By.XPATH, "//*[@id='contentControl']")
    driver.switch_to.frame(frame)
    relatorio = driver.find_element(By.ID, "menu2")
    planejamento = driver.find_element(By.ID, "menu3")

    # Download relatório gerencial
    relatorio.click()
    relatorio_gerencial = driver.find_element(By.ID, "botaoRelatorioGerencial")
    relatorio_gerencial.click()
    driver.switch_to.window(driver.window_handles[1])
    wait.until(EC.element_to_be_clickable((By.ID, "gerarCSV")))
    time.sleep(5)
    exportar_csv = driver.find_element(By.ID, "gerarCSV")
    exportar_csv.click()
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "baixarCsv")))
    time.sleep(10)
    baixar_csv = driver.find_element(By.CLASS_NAME, "baixarCsv")
    baixar_csv.click()
    time.sleep(10)

    # Download vizualiza planejamento
    driver.switch_to.window(driver.window_handles[0])
    driver.switch_to.frame(frame)
    planejamento.click()
    visualiza_planejamento = driver.find_element(By.ID, "visuItensPub")
    visualiza_planejamento.click()
    driver.switch_to.window(driver.window_handles[2])
    wait.until(EC.element_to_be_clickable((By.ID, "gerarCSV")))
    time.sleep(2)
    exportar_csv = driver.find_element(By.ID, "gerarCSV")
    exportar_csv.click()
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "baixarCsv")))
    time.sleep(10)
    baixar_csv = driver.find_element(By.CLASS_NAME, "baixarCsv")
    baixar_csv.click()
    time.sleep(5)
    driver.quit()