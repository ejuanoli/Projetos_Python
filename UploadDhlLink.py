from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


def fazer_upload_invisivel():
    # 1. Configurar o navegador para rodar oculto (Headless)
    opcoes = Options()
    # opcoes.add_argument("--headless=new")  # Mantido desativado para você ver o login acontecendo
    opcoes.add_argument("--window-size=1920,1080")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--no-sandbox")

    opcoes.add_argument("--disable-dev-shm-usage") # Resolve travamentos de memória
    opcoes.add_argument("--remote-debugging-port=9222") # Força uma porta de comunicação livre
    opcoes.add_argument("--ignore-certificate-errors") # Ignora erros de segurança da rede

    # Inicia o Chrome
    driver = webdriver.Chrome(options=opcoes)

    # Aumentei um pouco o tempo caso você precise digitar código de celular (MFA)
    wait = WebDriverWait(driver, 300)

    try:
        # 2. Tentar acessar o site de upload direto
        print("Acessando o site...")
        driver.get("https://link-cc.dhl.com/upload")

        # ====================================================================
        # LÓGICA DE REDIRECIONAMENTO PÓS-LOGIN
        # ====================================================================
        print("Aguardando autenticação e carregamento do painel DHL...")

        # Espera até que o botão "UPLOAD" do menu superior apareça.
        # Se isso apareceu, significa que o login terminou com sucesso.
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[text()='UPLOAD']")))

        # Dá 2 segundinhos para o site terminar qualquer redirecionamento maluco que ele faça
        time.sleep(2)

        # Verifica se o site nos jogou para a Home em vez da página de Upload
        if "/upload" not in driver.current_url:
            print("Site redirecionou para a Home. Forçando a volta para a tela de Upload...")
            driver.get("https://link-cc.dhl.com/upload")

            # Esperar a página de upload carregar novamente
            time.sleep(3)
        else:
            print("Já estamos na tela de Upload!")

        # ====================================================================
        # ETAPAS DO UPLOAD
        # ====================================================================
        print("Aguardando os elementos da página de upload carregarem...")

        # NOTA: Se os campos "Partner Locations" e "Flow" precisarem ser preenchidos
        # antes do site aceitar o arquivo, o robô terá que clicar neles aqui!

        # 3. Encontrar o campo de upload oculto
        campo_arquivo = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))

        caminho_do_arquivo = r"C:\Users\ejuanoli\Downloads\31260140154884000315550040001066131001066142_procNFe.xml"

        print("Anexando o arquivo silenciosamente...")
        campo_arquivo.send_keys(caminho_do_arquivo)

        # Pequena pausa para o React processar o arquivo anexado
        time.sleep(1.5)

        # 4. Encontrar e clicar no botão "Upload"
        print("Buscando o botão de Upload...")
        botao_upload = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Upload')]")))

        print("Clicando no botão...")
        driver.execute_script("arguments[0].click();", botao_upload)

        # 5. Esperar terminar
        time.sleep(5)
        print("Upload concluído com sucesso sem interromper o usuário!")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    finally:
        # Fechar o navegador no final
        print("Fechando navegador...")
        driver.quit()


if __name__ == "__main__":
    fazer_upload_invisivel()