import os
import time
import pyautogui as py
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import SessionNotCreatedException, WebDriverException


class Config:
    """Centraliza as configurações da automação de contratos."""

    DOWNLOAD_DIR = r"C:\\xampp\\htdocs\\embracon\\Contrato"
    USERNAME = "usecred.eireli@embracon.com.br"
    PASSWORD = "Us3Cr3d3@2036"
    LOGIN_URL = "https://parceirosweb.embraconnet.com.br/Newcon_Intranet/frmCorCCCnsLogin.aspx"
    URL_CONTRATO = (
        "https://parceirosweb.embraconnet.com.br/Newcon_Intranet/CONPV/frmConPvRelContrato.aspx?applicationKey="
        "1NmFNusRtxJxFVv3u8rP20QkKmoNa1QEYnUNwJ66m7o="
    )
    MAX_WAIT = 30
    SCREEN_WIDTH = 1920
    SCREEN_HEIGHT = 1080


def init_browser():
    """Configura e inicia o navegador Chrome."""

    options = webdriver.ChromeOptions()
    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('--log-level=3')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument(f'--window-size={Config.SCREEN_WIDTH},{Config.SCREEN_HEIGHT}')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": Config.DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
        },
    )

    service = Service(executable_path='chromedriver.exe')

    try:
        driver = webdriver.Chrome(service=service, options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        wait = WebDriverWait(driver, Config.MAX_WAIT)
        return driver, wait
    except SessionNotCreatedException as e:
        print(f"Erro de versão do ChromeDriver: {e}")
        raise
    except WebDriverException as e:
        print(f"Erro ao iniciar navegador: {e}")
        raise


class EmbraconAutomation:
    """Gerencia o fluxo de automação do contrato."""

    def __init__(self, driver, wait):
        self.driver = driver
        self.wait = wait

    def login(self):
        """Realiza o login na plataforma."""

        self.driver.get(Config.LOGIN_URL)
        time.sleep(5)

        self.wait.until(EC.presence_of_element_located((By.ID, "i0116"))).send_keys(Config.USERNAME)
        self.driver.find_element(By.ID, "idSIButton9").click()
        time.sleep(5)

        self.driver.find_element(By.ID, "passwordInput").send_keys(Config.PASSWORD)
        self.driver.find_element(By.ID, "submitButton").click()
        time.sleep(10)

        self.driver.find_element(By.ID, "idBtn_Back").click()
        time.sleep(5)

    def navigate_to_contract(self):
        """Navega até a página de emissão de contrato."""

        self.wait.until(
            EC.element_to_be_clickable((By.ID, "rptUnidadeNegocio_ctl03_txt_CD_Unidade_Negocio"))
        ).click()
        time.sleep(5)
        self.driver.find_element(By.ID, "ctl00_Conteudo_tvwEmpresat4").click()
        time.sleep(5)
        self.driver.get(Config.URL_CONTRATO)
        self.driver.find_element(By.ID, "ctl00_Conteudo_rblTipoPesquisa_1").send_keys("Número Contrato")
        time.sleep(3)

    def issue_contract(self, contratos):
        """Emite o PDF do contrato mantendo as etapas originais do robô."""

        for contrato in contratos:
            try:
                numero = str(contrato).strip()
                if not numero:
                    continue

                search_box = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_edtNO_Contrato"))
                )
                search_box.click()
                search_box.clear()
                search_box.send_keys(numero)
                time.sleep(1)

                self.driver.find_element(By.ID, "ctl00_Conteudo_cbxTipoDocumento").click()
                time.sleep(2)
                self.driver.find_element(By.ID, "ctl00_Conteudo_btnLocalizar").click()
                time.sleep(3)

                if "ctl00_Conteudo_lblErro" in self.driver.page_source:
                    print(f"Contrato {contrato} não encontrado, tentando o próximo…")
                    continue

                self.driver.find_element(By.ID, "ctl00_Conteudo_btnImprimir").click()
                time.sleep(5)

                py.click(x=675, y=161)  # ajusta o foco/aba do PDF se necessário
                time.sleep(2)

                py.hotkey("ctrl", "s")
                time.sleep(2)
                caminho_pdf = fr"{Config.DOWNLOAD_DIR}\\{numero}.pdf"
                py.typewrite(caminho_pdf)
                time.sleep(1)
                py.press("enter")
                time.sleep(2)
                print(f"Arquivo salvo como {numero}.pdf")
                self.wait.until(EC.element_to_be_clickable((By.ID, "download"))).click()
                time.sleep(5)
                py.click(x=909, y=629)  # clica em cancelar
                time.sleep(5)
                self.driver.find_element(By.ID, "ctl00_Conteudo_rblTipoPesquisa_1").send_keys("Número Contrato")
                time.sleep(3)
            except Exception as e:
                print(f"Erro ao processar contrato {numero}: {str(e)}")
                continue


def main(contrato: str) -> str:
    """Executa a automação para o contrato informado."""

    os.makedirs(Config.DOWNLOAD_DIR, exist_ok=True)
    driver, wait = init_browser()
    automator = EmbraconAutomation(driver, wait)

    contratos = [c.strip() for c in str(contrato).splitlines() if c.strip()]

    try:
        automator.login()
        automator.navigate_to_contract()
        automator.issue_contract(contratos)
    finally:
        driver.quit()

    return os.path.join(Config.DOWNLOAD_DIR, f"{contratos[0]}.pdf")


if __name__ == "__main__":
    contrato_num = input("Número do contrato: ")
    main(contrato_num)

