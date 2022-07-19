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
    driver.get(os.getenv("CONSTRUTIVO_URL"))
    

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
    time.sleep(1)
    exportar_csv = driver.find_element(By.ID, "gerarCSV")
    exportar_csv.click()
    #time.sleep(5)
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "baixarCsv")))
    baixar_csv = driver.find_element(By.CLASS_NAME, "baixarCsv")
    
    caminho_pasta_download = os.getenv("DOWNLOAD_PATH")
    check = len(os.listdir(caminho_pasta_download))

    while check == len(os.listdir(caminho_pasta_download)):
        baixar_csv.click()
        time.sleep(2)
    time.sleep(2)

    arquivo_novo = max(glob(caminho_pasta_download+f"/*.csv"), key=os.path.getctime)
    renomear = "C:/Users/guilherme.colla/Documents/Python Scripts/env1/Construtivo/gerencial_"+nome_pasta_empreendimento+".csv"
    try:
        os.remove(renomear)
    except Exception:
        pass
    os.rename(arquivo_novo, renomear)

    #os.rename(os.path.join(caminho_pasta_download, novo_arquivo), "gerencial".join(nome_pasta_empreendimento))

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
    while check == len(os.listdir(caminho_pasta_download)):
        baixar_csv.click()
        print("click baixar.")
        time.sleep(2)
    #time.sleep(5)
    
    arquivo_novo = max(glob(caminho_pasta_download+f"/*.csv"), key=os.path.getctime)
    renomear = "C:/Users/guilherme.colla/Documents/Python Scripts/env1/Construtivo/planejamento_"+nome_pasta_empreendimento+".csv"
    try:
        os.remove(renomear)
    except Exception:
        pass
    os.rename(arquivo_novo, renomear)
    
    driver.quit()

    
    
if __name__ == "__main__":
    download_gerencial_planejamento("CPFL_Sul_II_VMT_PE")
    download_gerencial_planejamento("CPFL_Sul_II_OSO3_PE")
    download_gerencial_planejamento("CPFL_Sul_II_PAL1_PE")
    download_gerencial_planejamento("CPFL_Sul_I_PE")
