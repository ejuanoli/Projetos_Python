"""
Cópia de teste do DownloadXMLDHLLink para GitHub Actions.
Só abre https://link-cc.dhl.com e verifica se a página carrega
(Home logada ou tela de login/SSO). Não baixa XML/PDF.

Nomenclatura no mesmo padrão do RPA_Apple_GHA.
"""

from __future__ import annotations

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from pathlib import Path
import os
import sys
import time

URL_DHL_LINK = "https://link-cc.dhl.com"
XPATH_HOME = "//a[.//*[@data-testid='LocationOnOutlinedIcon']]"

LARGURA_VIEWPORT = 1980
ALTURA_VIEWPORT = 1080

_BASE = Path(os.environ.get("GITHUB_WORKSPACE") or os.getcwd()).resolve()
PASTA_SAIDA = Path(os.environ.get("RPA_PASTA_DESTINO") or (_BASE / "saida_dhl_link"))


def _log(mensagem: str, nivel: str = "info") -> None:
    hora = datetime.now().strftime("%H:%M:%S")
    icones = {"etapa": "▶", "info": "●", "ok": "✓", "aviso": "!", "erro": "✗"}
    print(f"[{hora}] {icones.get(nivel, '●')} {mensagem}", flush=True)


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


def testar_acesso() -> int:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    _log("Teste GitHub Actions: acesso ao DHL Link.", "etapa")
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
        titulo = driver.title or ""
        _log(f"Página carregou. URL={url} | título={titulo}", "ok")

        try:
            wait_home = WebDriverWait(driver, 20)
            wait_home.until(EC.presence_of_element_located((By.XPATH, XPATH_HOME)))
            _log("Home do DHL Link detectada (card Partner/Flow).", "ok")
            _salvar_diagnostico(driver, "home")
            print("RESULTADO: ACESSOU A HOME (logado).")
            return 0
        except Exception:
            _log(
                "Home logada não apareceu (esperado se o SSO/MFA pedir interação).",
                "aviso",
            )

        url_l = url.lower()
        tit_l = titulo.lower()
        sinais_login = (
            "login",
            "signin",
            "sign-in",
            "sso",
            "microsoftonline",
            "okta",
            "auth",
            "sts",
            "adfs",
            "ping",
            "dhl",
            "link-cc",
        )
        if any(s in url_l or s in tit_l for s in sinais_login) or len(driver.page_source or "") > 500:
            _log("O site respondeu (login/SSO ou página DHL). Rede do Actions alcança o DHL Link.", "ok")
            print("RESULTADO: SITE ACESSÍVEL (parou no login/SSO).")
            return 0

        _log("Página vazia ou inesperada.", "erro")
        print("RESULTADO: FALHOU — página não parece o DHL Link.")
        return 1
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
