import os
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time


def download_gerencial_planejamento(nome_pasta_empreendimento: str) -> None:
    
    """Essa função é como eu estudei selenium para realizar o download
    dos relatórios brutos que preciso. Esse arquivo está aqui apenas para 
    referência. Essa função será implementada no Construtivo.py."""
 
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
    #time.sleep(5)

    #A partir daqui já estamos conectados na área de trabalho do colaborativo.
    #Navegando os menus.





    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[class='gwt-InlineLabel vibe-dataTableEntry-title']")))
    pastas = driver.find_elements(By.CSS_SELECTOR, "div[class='gwt-Label workspaceTreeBinderAnchor']")
    for pasta in pastas:
        if nome_pasta_empreendimento == pasta.text:
            pasta.click()
            break
    

    time.sleep(5)
    frame = driver.find_element(By.XPATH, "//*[@id='contentControl']")
    driver.switch_to.frame(frame)
    wait.until(EC.element_to_be_clickable((By.ID, "menu2")))
    relatorio = driver.find_element(By.ID, "menu2")
    planejamento = driver.find_element(By.ID, "menu3")

    # Download relatório gerencial
    relatorio.click()
    relatorio_gerencial = driver.find_element(By.ID, "botaoRelatorioGerencial")
    relatorio_gerencial.click()
    driver.switch_to.window(driver.window_handles[1])
    wait.until(EC.element_to_be_clickable((By.ID, "gerarCSV")))
    #time.sleep(5)
    exportar_csv = driver.find_element(By.ID, "gerarCSV")
    exportar_csv.click()
    #time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "baixarCsv")))
    #time.sleep(10)
    baixar_csv = driver.find_element(By.CLASS_NAME, "baixarCsv")
    baixar_csv.click()
    #time.sleep(10)

    # Download vizualiza planejamento
    driver.switch_to.window(driver.window_handles[0])
    driver.switch_to.frame(frame)
    planejamento.click()
    visualiza_planejamento = driver.find_element(By.ID, "visuItensPub")
    visualiza_planejamento.click()
    driver.switch_to.window(driver.window_handles[2])
    wait.until(EC.element_to_be_clickable((By.ID, "gerarCSV")))
    #time.sleep(2)
    exportar_csv = driver.find_element(By.ID, "gerarCSV")
    exportar_csv.click()
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "baixarCsv")))
    #time.sleep(10)
    baixar_csv = driver.find_element(By.CLASS_NAME, "baixarCsv")
    baixar_csv.click()
    #time.sleep(5)
    driver.quit()
    exit()
    
    
    
    
    CPFL_Geracao = driver.find_element(By.CSS_SELECTOR, "span[class='gwt-InlineLabel vibe-dataTableEntry-title']")
    CPFL_Geracao.click()
    projetos = []
    while len(projetos) <= 1:
        projetos = driver.find_elements(By.CSS_SELECTOR, "span[class='gwt-InlineLabel childBindersWidget_ListOfFoldersPanel_folderLabel']")
    
    print(f"{type(projetos)}\n{len(projetos)}")
    for element in projetos:
        if nome_pasta_empreendimento in element.text:
            element.click()
            break
            
    #time.sleep(5)
    frame = driver.find_element(By.XPATH, "//*[@id='contentControl']")
    print(frame)
    driver.switch_to.frame(frame)
    relatorio = driver.find_element(By.ID, "menu2")
    planejamento = driver.find_element(By.ID, "menu3")

    # Download relatório gerencial
    relatorio.click()
    relatorio_gerencial = driver.find_element(By.ID, "botaoRelatorioGerencial")
    relatorio_gerencial.click()
    driver.switch_to.window(driver.window_handles[1])
    wait.until(EC.element_to_be_clickable((By.ID, "gerarCSV")))
    #time.sleep(5)
    exportar_csv = driver.find_element(By.ID, "gerarCSV")
    exportar_csv.click()
    #time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "baixarCsv")))
    #time.sleep(10)
    baixar_csv = driver.find_element(By.CLASS_NAME, "baixarCsv")
    baixar_csv.click()
    #time.sleep(10)

    # Download vizualiza planejamento
    driver.switch_to.window(driver.window_handles[0])
    driver.switch_to.frame(frame)
    planejamento.click()
    visualiza_planejamento = driver.find_element(By.ID, "visuItensPub")
    visualiza_planejamento.click()
    driver.switch_to.window(driver.window_handles[2])
    wait.until(EC.element_to_be_clickable((By.ID, "gerarCSV")))
    #time.sleep(2)
    exportar_csv = driver.find_element(By.ID, "gerarCSV")
    exportar_csv.click()
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "baixarCsv")))
    #time.sleep(10)
    baixar_csv = driver.find_element(By.CLASS_NAME, "baixarCsv")
    baixar_csv.click()
    #time.sleep(5)
    driver.quit()

if __name__ == "__main__":
    download_gerencial_planejamento("CPFL_Sul_II_VMT_PE")
    #download_gerencial_planejamento("OSO3_PE")
    #download_gerencial_planejamento("PAL1_PE")
    #download_gerencial_planejamento("Sul_I_PE")