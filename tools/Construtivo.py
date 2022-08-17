'''
Utilizando PANDAS esse módulo vai formatar os relatórios extraídos do construtivo,
adequando-os para a ferramenta em desenvolvimento.
'''

import re
import shutil
import pandas as pd
from datetime import datetime, timedelta
import os
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time
from glob import glob


class Relatórios_Construtivo():
    def __init__(self, caminho_relatorio_gerencial: str, caminho_relatorio_planejamento):
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


    def dias_para_liberar(self):
        """Função que determina quantos dias o documento ficou com cada agente.
        Ela possui um problema gritante de excesso de 'if statements' devido a
        complexidade do fluxo que um documento pode apresentar."""
        #TODO Reduzir os if statements ou deixar mais profissional.


        def _calculo_de_delta(valor1: str, valor2: str) -> timedelta:
            """Retorna valor1 - valor2.
            Os valores são processados com base no formato %d/%m/%Y %H:%M"""

            ans = datetime.strptime(valor1, "%d/%m/%Y %H:%M") - datetime.strptime(valor2, "%d/%m/%Y %H:%M")
            return ans

        
        def _retorno_do_agente() -> str:
            for retorno in ["Aprovado", "Aprovado com Comentários", "Não Aprovado"]:
                if isinstance(datas[retorno], str):
                    return retorno

        deltas = {
        "Com CPFL" : [],
        "Tempo total de fluxo": [],
        "Com Projetista": [],
        "Com Acessadas": [],
        }


        for código, _ in self.relatorio_planejamento.iterrows():

            for key in deltas.keys():
                deltas[key].append(timedelta())
            filtro = self.relatorio_gerencial.query("Código == @código", inplace=False)
            
            for _, alteracao_no_fluxo in filtro.groupby(level=0):
            
                for indice, datas in alteracao_no_fluxo.iterrows():
                    
                    if indice[1] == "Liberado para Execução":
                        deltas["Tempo total de fluxo"][-1] += _calculo_de_delta(datas["Liberado para Execução"], datas["Primeira Emissão"])
                        
                        if isinstance(datas["Liberação CPFL"], str):
                            deltas["Com Projetista"][-1] += _calculo_de_delta(datas["Liberação CPFL"], datas["Distribuição"])
                            
                            if isinstance(datas["Para Análise Acessado_"], str):
                                deltas["Com CPFL"][-1] += _calculo_de_delta(datas["Para Análise Acessado_"], datas["Liberação CPFL"])
                                deltas["Com Acessadas"][-1] += _calculo_de_delta(datas["Liberado para Execução"], datas["Para Análise Acessado_"])
                                
                            else:
                                deltas["Com CPFL"][-1] += _calculo_de_delta(datas["Liberado para Execução"], datas["Liberação CPFL"])


                    if indice[1] == "Obsoleto":
                        
                        if isinstance(datas["Para Análise CPFL"], str):
                            deltas["Com Projetista"][-1] += _calculo_de_delta(datas["Para Análise CPFL"], datas["Distribuição"])
                            
                            if isinstance(datas["Para Análise Acessado"], str):
                                deltas["Com CPFL"][-1] += _calculo_de_delta(datas["Para Análise Acessado"],datas["Para Análise CPFL"])
                                deltas["Com Acessadas"][-1] += _calculo_de_delta(datas[_retorno_do_agente()], datas["Para Análise Acessado"])

                            else:
                                deltas["Com CPFL"][-1] += _calculo_de_delta(datas[_retorno_do_agente()], datas["Para Análise CPFL"])
                                deltas["Com Projetista"][-1] += _calculo_de_delta(datas["Obsoleto"], datas[_retorno_do_agente()])
 
                        
                        elif isinstance(datas["Liberação CPFL"], str):
                            deltas["Com Projetista"][-1] +=_calculo_de_delta(datas["Liberação CPFL"], datas["Distribuição"])
                            
                            if isinstance(datas["Para Análise Acessado_"], str):
                                deltas["Com CPFL"][-1] += _calculo_de_delta(datas["Para Análise Acessado_"], datas["Liberação CPFL"])
                                
                                if isinstance(datas["Para Revisão"], str) and isinstance(datas["Liberado para Execução"], str):
                                    if datetime.strptime(datas["Para Revisão"], "%d/%m/%Y %H:%M") > datetime.strptime(datas["Liberado para Execução"], "%d/%m/%Y %H:%M"):
                                        deltas["Com Acessadas"][-1] += _calculo_de_delta(datas["Liberado para Execução"], datas["Para Análise Acessado_"])
                                    else:
                                        deltas["Com Acessadas"][-1] += _calculo_de_delta(datas["Para Revisão"], datas["Para Análise Acessado_"])

                            if isinstance(datas["Liberado para Execução"], str):
                                deltas["Com CPFL"][-1] += _calculo_de_delta(datas["Liberado para Execução"], datas["Liberação CPFL"])
                            
                            if isinstance(datas["Para Revisão"], str):
                                deltas["Com Projetista"][-1] += _calculo_de_delta(datas["Obsoleto"], datas["Para Revisão"])
                                
                        else:
                            deltas["Com Projetista"][-1] += _calculo_de_delta(datas["Obsoleto"], datas["Distribuição"])
        
        for key in deltas.keys():
            self.relatorio_planejamento[key] = deltas[key]



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

    
    def aprovados(self) -> pd.DataFrame:
        return self.selecionar_estados(["Liberado para Execução"])



class Downloader():

    """Esta classe cria um chrome webdriver (janela de navegador chrome) para
    fazer o download dos relatórios gerenciais e de planejamento dos 
    empreendimentos da empresa no Construtivo."""

    
    def __init__(self) -> None:
        #Carregando os dados de usuário
        load_dotenv()
        self.caminho_pasta_download = os.getenv("DOWNLOAD_PATH")
        #Criando uma janela do Chrome.
        self.service = Service(executable_path=ChromeDriverManager().install())
        #Alterando 'log-level' para que o console levante apenas tipos de erro 'FATAL'.
        #Auth. note: Se eu lembrasse disso no começo do desenvolvimento desse módulo, tudo teria sido mais fácil de debugar....
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--log-level=3")
        self.driver = webdriver.Chrome(service=self.service, options=chrome_options)
        self.driver.implicitly_wait(10)
        self.wait = WebDriverWait(self.driver, 10)
        self.limpar_arquivos_antigos()        
        self.login()

    def limpar_arquivos_antigos(self):
        if os.path.isdir(os.getenv("SAVE_PATH")):
            shutil.rmtree(os.getenv("SAVE_PATH"))
        os.mkdir(os.getenv("SAVE_PATH"))


    def login(self) -> None:
        #Guardando a janela principal para depois.
        self.janela_principal = self.driver.current_window_handle
        #Conectando ao site colaborativo.
        self.driver.get(os.getenv("CONSTRUTIVO_URL"))
        #Obtendo as caixas de texto para login, senha e botão de entrar.
        self.wait.until(EC.element_to_be_clickable((By.ID, "loginOkBtn")))
        #Inserindo Login.
        self.driver.find_element(By.ID, "j_usernameId").send_keys(os.getenv("LOGIN"))
        #Inserindo Senha.
        self.driver.find_element(By.ID, "j_passwordId").send_keys(os.getenv("SENHA"))
        #Click no botão de login.
        self.driver.find_element(By.ID, "loginOkBtn").click()
        time.sleep(3)


    def acessar_empreendimento(self, nome_pasta_empreendimento: str) -> None:
        """O nome da pasta deve ser inserido por completo devido a logica
        de procura que é feito. Isso pode ser alterado em versões futuras
        dependendo da necessidade."""
        #A partir daqui já estamos conectados na área de trabalho do colaborativo.
        #Navegaremos as pastas pela hierarquia que fica na esquerda da pagina.
        self.driver.switch_to.default_content()
        self.nome_pasta_empreendimento = nome_pasta_empreendimento
        #self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[class='gwt-InlineLabel vibe-dataTableEntry-title']")))
        pastas = self.driver.find_elements(By.CSS_SELECTOR, "div[class='gwt-Label workspaceTreeBinderAnchor']")
        try:
            for pasta in pastas:
                if self.nome_pasta_empreendimento == pasta.text:
                    pasta.click()
                    break
        except Exception:
            print("Pasta não encontrada")
        time.sleep(3)
            

    def download_relatorio_gerencial(self) -> None:
        self.frame = self.driver.find_element(By.XPATH, "//*[@id='contentControl']")
        self.driver.switch_to.frame(self.frame)
        self.wait.until(EC.element_to_be_clickable((By.ID, "menu2")))
        #click no botão de relatório
        self.driver.find_element(By.ID, "menu2").click()
        
        # Download relatório gerencial
        self.driver.find_element(By.ID, "botaoRelatorioGerencial").click()
        self.driver.switch_to.window(self.driver.window_handles[1])
        self.wait.until(EC.element_to_be_clickable((By.ID, "gerarCSV")))
        self.driver.find_element(By.ID, "gerarCSV").click()
        self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "baixarCsv")))
        baixar_csv = self.driver.find_element(By.CLASS_NAME, "baixarCsv")
        
        check = len(os.listdir(self.caminho_pasta_download))
        _timeout = 0
        while check == len(os.listdir(self.caminho_pasta_download)):
            baixar_csv.click()
            time.sleep(2)
            _timeout += 1
            if _timeout > 10:
                self.driver.close()
                self.driver.switch_to.window(self.janela_principal)
                raise TimeoutError("Muitas tentativas de baixar documento. Reabrindo janela de relatório...")
        time.sleep(5)
        print(f"\n{self.nome_pasta_empreendimento} gerencial ok.\n")

        arquivo_novo = max(glob(self.caminho_pasta_download+f"/*.csv"), key=os.path.getctime)
        renomear = os.getenv("SAVE_PATH")
        renomear = renomear + "/gerencial_" + self.nome_pasta_empreendimento + ".csv"
        try:
            os.remove(renomear)
        except Exception:
            pass
        os.rename(arquivo_novo, renomear)
        self.driver.close()
        self.driver.switch_to.window(self.janela_principal)


    def download_visualiza_planejamento(self) -> None:
        # Download visualiza planejamento
        self.driver.switch_to.frame(self.frame)
        # Click no botão planejamento.
        self.driver.find_element(By.ID, "menu3").click()
        # Click no botão visualiza planejamento.
        self.driver.find_element(By.ID, "visuItensPub").click()
        self.driver.switch_to.window(self.driver.window_handles[1])
        self.wait.until(EC.element_to_be_clickable((By.ID, "gerarCSV")))
        self.driver.find_element(By.ID, "gerarCSV").click()
        self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "baixarCsv")))
        baixar_csv = self.driver.find_element(By.CLASS_NAME, "baixarCsv")

        check = len(os.listdir(self.caminho_pasta_download))
        _timeout = 0
        while check == len(os.listdir(self.caminho_pasta_download)):
            baixar_csv.click()
            print("click baixar.")
            time.sleep(2)
            _timeout += 1
            if _timeout > 10:
                self.driver.close()
                self.driver.switch_to.window(self.janela_principal)
                raise TimeoutError("Muitas tentativas de baixar documento. Reabrindo janela de relatório...")
        time.sleep(5)
        print(f"\n{self.nome_pasta_empreendimento} planejamento ok.\n")
        
        arquivo_novo = max(glob(self.caminho_pasta_download+f"/*.csv"), key=os.path.getctime)
        renomear = os.getenv("SAVE_PATH")
        renomear = renomear + "/planejamento_" + self.nome_pasta_empreendimento + ".csv"
        try:
            os.remove(renomear)
        except Exception:
            pass
        os.rename(arquivo_novo, renomear)
        self.driver.close()
        self.driver.switch_to.window(self.janela_principal)

    def fechar(self) -> None:
        self.driver.quit()
