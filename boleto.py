import os
import time
import random
import pyautogui as py
import socket
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, NoSuchElementException, WebDriverException, SessionNotCreatedException)

# ==============================================
# CONFIGURAÇÕES GERAIS
# ==============================================

class Config:
    """Centraliza todas as configurações do sistema"""

    # Diretórios
    DOWNLOAD_DIR = r"C:\xampp\htdocs\embracon\Boleto"
    CONTRATOS_DIR = r"C:\xampp\htdocs\embracon\Boleto"
    
    # Credenciais
    USERNAME = "usecred.eireli@embracon.com.br"
    PASSWORD = "Us3Cr3d3@2036"
    
    # URLs
    LOGIN_URL = "https://parceirosweb.embraconnet.com.br/Newcon_Intranet/frmCorCCCnsLogin.aspx"
    URL_CONTRATO = "https://parceirosweb.embraconnet.com.br/Newcon_Intranet/CONCO/frmConCoRelBoletoAvulsoEmissao.aspx?applicationKey=uaor0K5WheI/vPIGShIf/qLcrqK1UXqfukobJzMeDs4="
    
    # Comportamento
    TYPING_SPEED = 0.1  # segundos entre caracteres
    ERROR_RATE = 0.03    # 3% chance de erro de digitação
    MAX_WAIT = 30        # segundos para timeout
    LOGIN_RETRIES = 3    # tentativas de login
    HUMAN_DELAY = (0.1, 0.5)  # intervalo de atraso entre ações

    # Resolução fixa utilizada pelo robô (executa sempre em 1920x1080)
    SCREEN_WIDTH = 1920
    SCREEN_HEIGHT = 1080
    
    COORDINATES = {
        'login_button': (967, 675),
        'search_button': (1287, 326),
        # Adicione outras coordenadas fixas que você usa
    }


# ==============================================
# INICIALIZAÇÃO DO NAVEGADOR
# ==============================================

def init_browser():
    """Configura e inicializa o navegador Chrome de forma robusta"""
    
    print("🖥️ Inicializando navegador com configurações avançadas...")
    
    options = webdriver.ChromeOptions()

    # Configurações básicas
    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('--log-level=3')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    # options.add_argument('--headless=new')  # Descomente para modo headless
    options.add_argument(f'--window-size={Config.SCREEN_WIDTH},{Config.SCREEN_HEIGHT}')
    
    # Prevenção contra detecção
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
    
    # Configurações de download
    # Nas opções do ChromeDriver, adicione:
    options.add_experimental_option("prefs", {
        "download.default_directory": r"C:\xampp\htdocs\embracon\Boleto",
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
})
    
    # Configurar serviço do ChromeDriver
    service = Service(executable_path='chromedriver.exe')
    
    # Tentativa de inicialização com tratamento de erros
    try:
        driver = webdriver.Chrome(service=service, options=options)

        # Remover indicadores de automação
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        driver.execute_cdp_cmd('Network.setUserAgentOverride', {
            "userAgent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        })

        return driver, WebDriverWait(driver, Config.MAX_WAIT)
    
    except SessionNotCreatedException as e:
        print(f"🚨 Erro de versão do ChromeDriver: {str(e)}")
        print("Por favor, verifique se a versão do ChromeDriver é compatível com seu Chrome.")
        raise
    except WebDriverException as e:
        print(f"🚨 Erro ao iniciar navegador: {str(e)}")
        raise

# ==============================================
# COMPORTAMENTO HUMANIZADO
# ==============================================

class HumanBehavior:
    """Simula comportamentos humanos realistas"""
    
    @staticmethod
    def get_screen_scale():
        """Calcula o fator de escala baseado na resolução atual vs configurada"""
        screen_width, screen_height = py.size()
        width_scale = screen_width / Config.SCREEN_WIDTH
        height_scale = screen_height / Config.SCREEN_HEIGHT
        return width_scale, height_scale
    
    @staticmethod
    def scale_coordinates(x, y):
        """Ajusta as coordenadas para a resolução atual"""
        width_scale, height_scale = HumanBehavior.get_screen_scale()
        scaled_x = int(x * width_scale)
        scaled_y = int(y * height_scale)
        return scaled_x, scaled_y
    
    @staticmethod
    def random_delay(min_delay=None, max_delay=None):
        """Pausa aleatória entre ações humanas"""
        if min_delay is None or max_delay is None:
            min_delay, max_delay = Config.HUMAN_DELAY
        
        delay = random.uniform(min_delay, max_delay)
        time.sleep(delay)
    
    @staticmethod
    def occasional_long_pause():
        """Pausa mais longa ocasional para simular distração"""
        if random.random() < 0.05:  # 5% chance
            pause = random.uniform(0.8, 2.5)
            print(f"⏳ Pausa humana de {pause:.1f} segundos...")
            time.sleep(pause)
    
    @staticmethod
    def human_move(x, y, duration=None):
        """Movimento do mouse com trajetória natural"""
        if duration is None:
            duration = random.uniform(0.3, 1.0)
        
        # Variação na posição final para parecer mais humano
        x_final = x + random.randint(-5, 5)
        y_final = y + random.randint(-5, 5)
        
        py.moveTo(x_final, y_final, duration=duration, tween=py.easeInOutQuad)
        HumanBehavior.random_delay()
    
    @staticmethod
    def human_move_to_element(element):
        """Move o mouse para um elemento de forma humanizada"""
        location = element.location
        size = element.size
        x = location['x'] + size['width'] // 2 + random.randint(-5, 5)
        y = location['y'] + size['height'] // 2 + random.randint(-5, 5)
        HumanBehavior.human_move(x, y)
    
    @staticmethod
    def human_click(x=None, y=None, element=None, coordinate_name=None):
        """Clique humanizado com suporte a diferentes resoluções"""
        if coordinate_name:
            # Obter coordenadas da configuração
            x, y = Config.COORDINATES[coordinate_name]
        
        if element is not None:
            # Clicar via Selenium (melhor opção)
            HumanBehavior.human_move_to_element(element)
            element.click()
        elif x is not None and y is not None:
            # Escalar coordenadas
            x, y = HumanBehavior.scale_coordinates(x, y)
            HumanBehavior.human_move(x, y)
            py.click()
        
        HumanBehavior.random_delay()
        HumanBehavior.occasional_long_pause()
    
    @staticmethod
    def human_type(text, speed=Config.TYPING_SPEED, error_chance=Config.ERROR_RATE):
        """Digitação humanizada com erros ocasionais"""
        random.seed(datetime.now().timestamp())
        
        for char in text:
            # Variação na velocidade de digitação
            typing_speed = max(0, speed * random.uniform(0.5, 1.5))
            
            # Simulação de erro de digitação
            if random.random() < error_chance and len(text) > 3:
                wrong_char = chr(ord(char) + random.randint(-2, 2))
                py.typewrite(wrong_char)
                time.sleep(typing_speed * 0.7)
                
                # Correção do erro
                py.press('backspace')
                time.sleep(typing_speed * 0.3)
                
                py.typewrite(char)
            else:
                py.typewrite(char)
            
            time.sleep(typing_speed)
            
            # Pausa ocasional durante a digitação
            if random.random() < 0.03:
                pause = random.uniform(0.3, 1.0)
                time.sleep(pause)

# ==============================================
# FUNÇÕES PRINCIPAIS DE AUTOMAÇÃO
# ==============================================

class EmbraconAutomation:
    """Classe principal que gerencia a automação"""
    
    def __init__(self, driver, wait):
        self.driver = driver
        self.wait = wait
        self.login_attempts = 0
    
    def check_browser_alive(self):
        """Verifica se o navegador ainda está respondendo"""
        try:
            self.driver.current_url
            return True
        except WebDriverException:
            return False
    
    def restart_browser(self):
        """Reinicia o navegador em caso de falha"""
        print("🔄 Reiniciando navegador...")
        try:
            self.driver.quit()
        except:
            pass
        
        try:
            self.driver, self.wait = init_browser()
            return True
        except Exception as e:
            print(f"🚨 Falha ao reiniciar navegador: {str(e)}")
            return False
    
    def login(self):
        """Fluxo de login totalmente revisado e testado"""
        while self.login_attempts < Config.LOGIN_RETRIES:
            try:
                print(f"\n🔑 Tentativa de login {self.login_attempts + 1} de {Config.LOGIN_RETRIES}")
                
                # 1. Acessar página de login
                self.driver.get(Config.LOGIN_URL)
                time.sleep(2)  # Espera inicial obrigatória

                # 2. Preencher email - Método ultraconfiável
                print("📩 Inserindo email...")
                try:
                    # Tentar localizar o campo de email de várias formas diferentes
                    email_field = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='email'], input[name='loginfmt'], #i0116")),
                        message="Campo de email não encontrado"
                    )
                    
                    # Rolagem até o elemento e clique manual
                    self.driver.execute_script("arguments[0].scrollIntoView();", email_field)
                    email_field.click()
                    time.sleep(0.5)
                    
                    # Limpar campo e digitar caractere por caractere
                    email_field.clear()
                    for char in Config.USERNAME:
                        email_field.send_keys(char)
                        time.sleep(random.uniform(0.05, 0.15))
                    
                    time.sleep(1)
                    
                    # 3. Clicar no botão Avançar
                    next_button = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit'][value='Avançar'], #idSIButton9")),
                        message="Botão Avançar não encontrado"
                    )
                    next_button.click()
                    time.sleep(3)  # Espera crítica

                except Exception as e:
                    print(f"⚠️ Erro no preenchimento do email: {str(e)}")
                    self.driver.save_screenshot("erro_email.png")
                    raise

                # 4. Preencher senha - Método garantido
                print("🔒 Inserindo senha...")
                try:
                    password_field = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='password'], input[name='passwd'], #i0118")),
                        message="Campo de senha não encontrado"
                    )
                    
                    # Digitação humana com foco garantido
                    password_field.click()
                    time.sleep(0.5)
                    password_field.clear()
                    for char in Config.PASSWORD:
                        password_field.send_keys(char)
                        time.sleep(random.uniform(0.05, 0.2))
                        if random.random() < 0.05:  # 5% chance de erro
                            password_field.send_keys(Keys.BACK_SPACE)
                            time.sleep(0.3)
                            password_field.send_keys(char)
                    
                    time.sleep(1)
                    
                    # 5. Clicar no botão Entrar
                    signin_button = self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit'][value*='Entrar'], #submitButton")),
                        message="Botão Entrar não encontrado"
                    )
                    signin_button.click()
                    time.sleep(5)

                except Exception as e:
                    print(f"⚠️ Erro no preenchimento da senha: {str(e)}")
                    self.driver.save_screenshot("erro_senha.png")
                    raise

                # 6. Pular "Manter conectado" se aparecer
                try:
                    no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "idBtn_Back"))
                    )
                    no_button.click()
                    time.sleep(2)
                except:
                    print("ℹ️ Página de 'permanecer conectado' não apareceu")
                    pass

                # 7. Verificar login bem-sucedido
                try:
                    no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "rptUnidadeNegocio_ctl03_txt_CD_Unidade_Negocio"))
                    )
                    no_button.click()
                    time.sleep(2)
                    
                    print("🎉 Login realizado com sucesso!")
                    return True

                except TimeoutException:
                    print("⚠️ Possível falha no login - verificando mensagens de erro...")
                    self.check_login_error()
                    raise

            except Exception as e:
                self.login_attempts += 1
                print(f"⚠️ Falha no login (tentativa {self.login_attempts}): {str(e)}")
                
                # Tirar screenshot para debug
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                self.driver.save_screenshot(f"login_error_{timestamp}.png")
                
                if self.login_attempts >= Config.LOGIN_RETRIES:
                    print("🚫 Número máximo de tentativas atingido")
                    return False
                
                time.sleep(5 * self.login_attempts)  # Espera exponencial
        
        return False
    
    
    def navigate_to_billing(self):
        """Navega até a seção de boletos"""
        print("🧭 Navegando para a seção de boletos...")
        try:
                
                no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_tvwEmpresat4"))
                    )
                no_button.click()
                time.sleep(2)
                
                
                no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "CO"))
                    )
                no_button.click()
                time.sleep(2)       
                
                
            #  Acessar página de login
                self.driver.get(Config.URL_CONTRATO)
                time.sleep(3)  # Espera inicial obrigatória
                    
        except Exception as e:
            self.navigate_to_billing += 1
            print(f"⚠️ Falha no login (tentativa {self.navigate_to_billing}): {str(e)}")
                                   
    def issue_billing(self, contrato):
        for i, contrato in enumerate(contrato, 1):
            try:
                #clica em Busca avançada
                no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_identificacao_cota_btnBuscaAvancada"))
                    )
                no_button.click()
                time.sleep(2)
                
                # Clicar no Nome
                dropdown = self.wait.until(
                    EC.presence_of_element_located((By.ID, "ctl00_Conteudo_cbxCriterioBusca"))
                )
                # Usa a classe Select para manipular o dropdown
                select = Select(dropdown)
        
                # Seleciona por valor "C" (Contrato)
                select.select_by_value("C")
                time.sleep(3)
                
                search_field = self.driver.find_element(By.ID, "ctl00_Conteudo_edtContextoBusca")
                HumanBehavior.human_type(str(contrato), speed=0.15, error_chance=0.03)
                time.sleep(2)
                
                
                no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_btnBuscar"))
                    )
                no_button.click()
                time.sleep(2)
                
                
                href = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_grdBuscaAvancada_ctl02_lnkNM_Pessoa"))
                    )
                href.click()
                time.sleep(2)
                
                no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_btnConfirma"))
                    )
                no_button.click()
                time.sleep(2)
                
                no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_identificacao_cota_btnLocaliza"))
                    )
                no_button.click()
                time.sleep(2)
                
                
                # Clicar na caixa do boleto
                input = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_grdBoleto_Avulso_ctl03_imgEmite_Boleto"))
                    )
                input.click()
                time.sleep(2)
                
                
                
                # abaixa a tela
                HumanBehavior.human_click(x=1917, y=927)
                time.sleep(2)
                
                no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_btnEmitir"))
                    )
                no_button.click()
                time.sleep(2)
                
                # Salvar (Ctrl+S) e nomear com o número do contrato
                py.hotkey("ctrl", "s")
                time.sleep(2)
                caminho_pdf = fr"C:\xampp\htdocs\embracon\Boleto\{contrato}.pdf"
                py.typewrite(caminho_pdf)
                time.sleep(1)
                py.press("enter")
                time.sleep(2)
                print(f"Arquivo salvo como {contrato}.pdf")
                py.click(x=763, y=15) #fecha a aba do download
                time.sleep(5)
                py.click(x=909, y=629) #clica em cancelar
                time.sleep(5)
                no_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_Conteudo_rblTipoPesquisa_1ctl00_Conteudo_rblTipoPesquisa_1"))
                    )
                no_button.click()
                time.sleep(2)
                
            except Exception as e:
                print(f"⚠️ Erro ao processar contrato {contrato}: {str(e)}")
                continue

def _return_to_home(self):
        """Retorna para a tela inicial após emitir um boleto"""
        print("🏠 Voltando para tela inicial...")
        
        HumanBehavior.human_click(x=763, y=15)
        time.sleep(10)
        HumanBehavior.human_click(x=909, y=629)
        time.sleep(10)
        
        # Alterar tipo de pesquisa
        search_type = self.driver.find_element(By.ID, "ctl00_Conteudo_rblTipoPesquisa_1")
        HumanBehavior.human_click(element=search_type)
        time.sleep(10)        

# ==============================================
# EXECUÇÃO PRINCIPAL
# ==============================================

def main(contrato: str) -> str:
    """Emite o boleto para o contrato informado."""
    print("="*50)
    print("INICIANDO AUTOMAÇÃO EMBRACON - EMISSÃO DE BOLETOS")
    print("="*50)

    os.makedirs(Config.DOWNLOAD_DIR, exist_ok=True)
    os.makedirs(Config.CONTRATOS_DIR, exist_ok=True)
    print("✅ Diretórios verificados/criados")

    # Inicializar navegador
    try:
        print("\n🖥️ Inicializando navegador...")
        driver, wait = init_browser()
        automator = EmbraconAutomation(driver, wait)
    except Exception as e:
        print(f"🚨 Falha crítica ao iniciar navegador: {str(e)}")
        return ""

    # Executar fluxo principal
    try:
        print("\n🚀 Iniciando fluxo principal...")
        if automator.login():
            automator.navigate_to_billing()
            pdf_path = automator.issue_billing([contrato])
        else:
            print("🚫 Não foi possível fazer login após várias tentativas")
    except Exception as e:
        print(f"🚨 Erro fatal durante execução: {str(e)}")
    finally:
        print("\n" + "="*50)
        print("PROCESSO CONCLUÍDO")
        print("="*50)
        print("O navegador permanecerá aberto para verificação.")

    pdf_path = pdf_path or os.path.join(Config.CONTRATOS_DIR, f"{contrato}.pdf")
    return pdf_path


if __name__ == "__main__":
    py.PAUSE = 0.1
    py.FAILSAFE = True
    contrato = input("Número do contrato: ")
    main(contrato)