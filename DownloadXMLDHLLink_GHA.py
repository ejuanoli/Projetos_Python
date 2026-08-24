"""
Cópia de teste do DownloadXMLDHLLink para GitHub Actions.
Abre https://link-cc.dhl.com, faz login Microsoft com a conta robótica
e verifica se chega na Home. Não baixa XML/PDF.

Nomenclatura no mesmo padrão do RPA_Apple_GHA.
"""

from __future__ import annotations

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from datetime import datetime
from pathlib import Path
import os
import sys
import time

URL_DHL_LINK = "https://link-cc.dhl.com"
XPATH_HOME = "//a[.//*[@data-testid='LocationOnOutlinedIcon']]"
MS_USER_PADRAO = "PETASL.ProcessAssistant@dhl.com"

LARGURA_VIEWPORT = 1980
ALTURA_VIEWPORT = 1080

_BASE = Path(os.environ.get("GITHUB_WORKSPACE") or os.getcwd()).resolve()
PASTA_SAIDA = Path(os.environ.get("RPA_PASTA_DESTINO") or (_BASE / "saida_dhl_link"))


def _log(mensagem: str, nivel: str = "info") -> None:
    hora = datetime.now().strftime("%H:%M:%S")
    icones = {"etapa": "▶", "info": "●", "ok": "✓", "aviso": "!", "erro": "✗"}
    print(f"[{hora}] {icones.get(nivel, '●')} {mensagem}", flush=True)


def _credencial(*nomes: str) -> str:
    for nome in nomes:
        valor = os.environ.get(nome) or ""
        if valor.strip():
            return valor
    raise RuntimeError(
        "Nenhuma das variáveis foi definida: " + ", ".join(nomes)
    )


def _salvar_diagnostico(driver, etapa: str) -> None:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        _log(f"URL: {driver.current_url}", "info")
        _log(f"Título: {driver.title}", "info")
    except Exception:
        pass
    png = PASTA_SAIDA / f"diagnostico_{etapa}_{stamp}.png"
    html = PASTA_SAIDA / f"diagnostico_{etapa}_{stamp}.html"
    try:
        driver.save_screenshot(str(png))
        _log(f"Screenshot: {png.name}", "info")
    except Exception as e:
        _log(f"Não foi possível salvar screenshot: {e}", "aviso")
    try:
        html.write_text(driver.page_source or "", encoding="utf-8")
        _log(f"HTML: {html.name}", "info")
    except Exception as e:
        _log(f"Não foi possível salvar HTML: {e}", "aviso")


def _criar_opcoes_chrome() -> Options:
    opcoes = Options()
    if os.environ.get("RPA_HEADLESS", "").strip() == "1":
        opcoes.add_argument("--headless=new")
    opcoes.add_argument(f"--window-size={LARGURA_VIEWPORT},{ALTURA_VIEWPORT}")
    opcoes.add_argument("--force-device-scale-factor=1")
    opcoes.add_argument("--high-dpi-support=1")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument("--disable-dev-shm-usage")
    opcoes.add_argument("--ignore-certificate-errors")
    opcoes.add_argument("--disable-blink-features=AutomationControlled")
    opcoes.add_experimental_option("excludeSwitches", ["enable-automation"])
    opcoes.add_experimental_option("useAutomationExtension", False)
    return opcoes


def _forcar_resolucao(driver) -> None:
    try:
        driver.set_window_size(LARGURA_VIEWPORT, ALTURA_VIEWPORT)
    except Exception:
        pass
    try:
        driver.execute_cdp_cmd(
            "Emulation.setDeviceMetricsOverride",
            {
                "width": LARGURA_VIEWPORT,
                "height": ALTURA_VIEWPORT,
                "deviceScaleFactor": 1,
                "mobile": False,
            },
        )
    except Exception:
        pass


def criar_driver():
    try:
        driver = webdriver.Chrome(options=_criar_opcoes_chrome())
    except Exception as e:
        raise RuntimeError(f"Não foi possível abrir o Google Chrome: {e}") from e
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": (
                    "Object.defineProperty(navigator, 'webdriver', "
                    "{get: () => undefined});"
                )
            },
        )
    except Exception:
        pass
    _forcar_resolucao(driver)
    return driver


def _pagina_carregou(driver) -> bool:
    try:
        url = (driver.current_url or "").strip()
        ready = driver.execute_script("return document.readyState") or ""
        return bool(url) and url != "data:," and ready in ("interactive", "complete")
    except Exception:
        return False


def _clicar_primeiro(driver, wait, seletores, timeout=8) -> bool:
    wait_local = WebDriverWait(driver, timeout)
    for seletor in seletores:
        try:
            if seletor.startswith("//") or seletor.startswith("("):
                el = wait_local.until(EC.element_to_be_clickable((By.XPATH, seletor)))
            else:
                el = wait_local.until(EC.element_to_be_clickable((By.CSS_SELECTOR, seletor)))
            driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            continue
    return False


def _preencher_primeiro(driver, wait, seletores, valor: str, timeout=20):
    wait_local = WebDriverWait(driver, timeout)
    ultimo_erro = None
    for seletor in seletores:
        try:
            if seletor.startswith("//"):
                loc = (By.XPATH, seletor)
            else:
                loc = (By.CSS_SELECTOR, seletor)
            el = wait_local.until(EC.presence_of_element_located(loc))
            el.click()
            time.sleep(0.2)
            el.clear()
            el.send_keys(valor)
            return el
        except Exception as e:
            ultimo_erro = e
            continue
    raise TimeoutException(
        f"Não achei campo de login Microsoft ({seletores}). Último erro: {ultimo_erro}"
    )


def _login_microsoft(driver, usuario: str, senha: str) -> None:
    """Preenche e-mail e senha na tela login.microsoftonline.com."""
    wait = WebDriverWait(driver, 25)
    _log(f"Login Microsoft: preenchendo usuário {usuario}...", "etapa")

    # Lista "escolher conta" / usar outra conta
    _clicar_primeiro(
        driver,
        wait,
        [
            "#otherTile",
            "//div[contains(@data-test-id,'otherTile')]",
            f"//div[contains(., '{usuario}')]",
        ],
        timeout=5,
    )
    time.sleep(1)

    _preencher_primeiro(
        driver,
        wait,
        ["input[name='loginfmt']", "#i0116", "input[type='email']"],
        usuario,
    )
    _clicar_primeiro(
        driver,
        wait,
        ["#idSIButtonNext", "input[type='submit']", "input[value='Next']"],
        timeout=8,
    )
    time.sleep(2)

    # Tela de passkey / "usar senha"
    _clicar_primeiro(
        driver,
        wait,
        [
            "#idA_PWD_SwitchToPassword",
            "//a[contains(., 'password') or contains(., 'senha')]",
            "//span[contains(., 'Use your password')]",
            "//button[contains(., 'Use your password')]",
        ],
        timeout=5,
    )
    time.sleep(1)

    _log("Login Microsoft: preenchendo senha...", "etapa")
    campo_senha = _preencher_primeiro(
        driver,
        wait,
        ["input[name='passwd']", "#i0118", "input[type='password']"],
        senha,
    )
    campo_senha.send_keys(Keys.RETURN)
    time.sleep(1)
    _clicar_primeiro(
        driver,
        wait,
        ["#idSIButtonNext", "input[type='submit']", "input[value='Sign in']"],
        timeout=5,
    )
    time.sleep(3)

    # Permanecer conectado? Não.
    if _clicar_primeiro(
        driver,
        wait,
        ["#idBtn_Back", "#idSIButton9"],
        timeout=8,
    ):
        _log("Tela 'Stay signed in' tratada.", "info")
        time.sleep(2)

    for seletor_erro in ("#passwordError", "#usernameError", "#errorText"):
        try:
            el = driver.find_element(By.CSS_SELECTOR, seletor_erro)
            if el.is_displayed() and el.text.strip():
                raise RuntimeError(f"Microsoft recusou o login: {el.text.strip()}")
        except RuntimeError:
            raise
        except Exception:
            continue


def testar_acesso() -> int:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    usuario = os.environ.get("MS_USER", MS_USER_PADRAO).strip() or MS_USER_PADRAO
    senha = _credencial("MS_PASSWORD", "OUTLOOK_PASSWORD", "DHL_LINK_PASSWORD")

    _log("Teste GitHub Actions: acesso ao DHL Link com login Microsoft.", "etapa")
    _log(f"Abrindo {URL_DHL_LINK}...", "info")

    driver = None
    try:
        driver = criar_driver()
        wait_curto = WebDriverWait(driver, 45)
        driver.get(URL_DHL_LINK)
        _forcar_resolucao(driver)

        wait_curto.until(lambda d: _pagina_carregou(d))
        time.sleep(3)
        _salvar_diagnostico(driver, "apos_get")

        url = driver.current_url or ""
        _log(f"Página carregou. URL={url} | título={driver.title}", "ok")

        ja_na_home = bool(driver.find_elements(By.XPATH, XPATH_HOME))
        if not ja_na_home:
            _login_microsoft(driver, usuario, senha)
            _salvar_diagnostico(driver, "apos_login_ms")

        wait_home = WebDriverWait(driver, 60)
        wait_home.until(EC.presence_of_element_located((By.XPATH, XPATH_HOME)))
        _log("Home do DHL Link detectada (card Partner/Flow).", "ok")
        _salvar_diagnostico(driver, "home")
        print("RESULTADO: ACESSOU A HOME (logado).")
        return 0
    except Exception as e:
        _log(f"Falha ao acessar o DHL Link: {e}", "erro")
        if driver:
            _salvar_diagnostico(driver, "erro")
        print(f"RESULTADO: FALHOU — {e}")
        return 1
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == "__main__":
    sys.exit(testar_acesso())
