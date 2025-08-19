from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui as py
import pandas as pd
import os

# Configuração inicial
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)

# Carregar planilha
# data = pd.read_excel(r'C:\\Users\\Servidor_Cameras\\Desktop\\Vendas Consorcio Diaria.xlsx')
data = pd.read_excel(r'C:\\xampp\\htdocs\\embracon\\Vendas Consorcio Diaria.xlsx')

contratos = data.iloc[:, 5].dropna().astype(str).str.replace("\.0$", "", regex=True)  # Removendo ".0" do final

# Abrir página de login
driver.get("https://parceirosweb.embraconnet.com.br/Newcon_Intranet/frmCorCCCnsLogin.aspx")
time.sleep(5)

# Fazer login
wait.until(EC.presence_of_element_located((By.ID, "i0116"))).send_keys("usecred.eireli@embracon.com.br")
driver.find_element(By.ID, "idSIButton9").click()
time.sleep(5)
driver.find_element(By.ID, "passwordInput").send_keys("Us3Cr3d3@2036")
driver.find_element(By.ID, "submitButton").click()
time.sleep(10)
driver.find_element(By.ID, "idBtn_Back").click()
time.sleep(5)

# garante que a pasta existe
os.makedirs(r"C:\xampp\htdocs\embracon\Boleto", exist_ok=True)

def emitir_boletos(contratos):
    # Navegar até a emissão de cobrança
    wait.until(EC.element_to_be_clickable((By.ID, "rptUnidadeNegocio_ctl03_txt_CD_Unidade_Negocio"))).click()
    time.sleep(5)
    driver.find_element(By.ID, "ctl00_Conteudo_tvwEmpresat4").click()
    time.sleep(5)
    driver.get("https://parceirosweb.embraconnet.com.br/Newcon_Intranet/CONPV/frmConPvRelContrato.aspx?applicationKey=1NmFNusRtxJxFVv3u8rP20QkKmoNa1QEYnUNwJ66m7o=")

    driver.find_element(By.ID, "ctl00_Conteudo_rblTipoPesquisa_1").send_keys("Número Contrato")
    time.sleep(3)

    for contrato in contratos:
        try:
            search_box = driver.find_element(By.ID, "ctl00_Conteudo_edtNO_Contrato")
            time.sleep(1)
            search_box.clear()
            search_box.send_keys(str(contrato))
            time.sleep(1)

            driver.find_element(By.ID, "ctl00_Conteudo_cbxTipoDocumento").click()
            time.sleep(2)
            driver.find_element(By.ID, "ctl00_Conteudo_btnLocalizar").click()
            time.sleep(3)

            # Se não encontrou, a página geralmente mostra um label de erro
            if "ctl00_Conteudo_lblErro" in driver.page_source:
                print(f"Contrato {contrato} não encontrado, tentando o próximo…")
                continue

            driver.find_element(By.ID, "ctl00_Conteudo_btnImprimir").click()
            time.sleep(5)
            
            

            # <<< SEU BLOCO DENTRO DO TRY >>>
            
            
            py.click(x=675, y=161)   # ajusta o foco/aba do PDF se necessário
            time.sleep(2)

            # Salvar (Ctrl+S) e nomear com o número do contrato
            py.hotkey("ctrl", "s")
            time.sleep(2)
            caminho_pdf = fr"C:\xampp\htdocs\embracon\Contrato\{contrato}.pdf"
            py.typewrite(caminho_pdf)
            time.sleep(1)
            py.press("enter")
            time.sleep(2)
            print(f"Arquivo salvo como {contrato}.pdf")
            wait.until(EC.element_to_be_clickable((By.ID, "download"))).click()
            time.sleep(5)
            py.click(x=909, y=629) #clica em cancelar
            time.sleep(5)
            driver.find_element(By.ID, "ctl00_Conteudo_rblTipoPesquisa_1").send_keys("Número Contrato")
            time.sleep(3)

        except Exception as e:
            print(f"Erro ao processar contrato {contrato}: {str(e)}")
            continue  # continua para o próximo contrato mesmo se der erro

# Loop pelos contratos (atenção: aqui você estava lendo outra planilha/coluna diferente)
# Se isso for proposital, ok; se não, use a mesma fonte/coluna que acima
data = pd.read_excel(r'C:\\Users\\Servidor_Cameras\\Desktop\\Vendas Consorcio Diaria.xlsx')
contratos = data.iloc[:, 5].dropna().astype(str).str.replace("\.0$", "", regex=True)


emitir_boletos(contratos)
driver.quit()
