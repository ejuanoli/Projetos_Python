"""
Download XML/PDF do DHL Link para GitHub Actions.
Sem interface gráfica e sem SharePoint.

Fluxo: login Microsoft (e-mail e senha) → Partner/Flow → filtro de datas →
relatório XML → confronto com a tabela xmls_dhl_link no Supabase →
download dos faltantes (IN Files) → e-mail Gmail (XML + PDF) →
grava o nome de cada XML enviado no Supabase.
"""

from __future__ import annotations

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    InvalidSessionIdException,
    StaleElementReferenceException,
    TimeoutException,
)
from datetime import date, datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
import json
import os
import re
import shutil
import smtplib
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile

URL_DHL_LINK = "https://link-cc.dhl.com"
XPATH_HOME = "//a[.//*[@data-testid='LocationOnOutlinedIcon']]"
MS_USER_PADRAO = "PETASL.ProcessAssistant@dhl.com"

FLOW_BUSCA_XML = "com.dhl.tech.apple.br_0283.tms_danfe_mf"
FLOW_BUSCA_PDF = "com.dhl.tech.apple.br_0283.pdf_danfe_mf"
FLOWS_BUSCA = (FLOW_BUSCA_XML, FLOW_BUSCA_PDF)

XPATH_INPUT_BUSCA_FLOW = (
    "//input[contains(@id,'flow_selector') and contains(@id,'flow-search-text-input')]"
)
XPATH_CAMPOS_DATA_FILTRO = "//input[@placeholder='YYYY/MM/DD']"
XPATH_CAMPO_BUSCA_MENSAGENS = "//input[@aria-label='message-search-textbox']"
XPATH_SELECT_TAMANHO_PAGINA = (
    "//div[contains(normalize-space(text()),'Page size')]"
    "//div[@role='combobox' and contains(@class,'MuiSelect-select')]"
)

DIAS_PARA_TRAS_NO_FILTRO = 3
DIAS_PARA_FRENTE_NO_FILTRO = 1
FORMATO_DATA_FILTRO = "%Y/%m/%d"
TAMANHO_PAGINA_DESEJADO = "1000"
LIMITE_PEDIDOS_POR_BUSCA = 30
LIMITE_EMAIL_BYTES = 20 * 1024 * 1024

PADRAO_PEDIDO = re.compile(r"QBQ\d{7,}", re.IGNORECASE)
PADRAO_NFE_ARQUIVO = re.compile(r"NFE-(\d{44})", re.IGNORECASE)
PADRAO_CHAVE_AVULSA = re.compile(r"(?<!\d)(\d{44})(?!\d)")
PADRAO_NNF_XML = re.compile(r"<nNF>\s*(\d+)\s*</nNF>", re.IGNORECASE)
PADRAO_CHNFE_XML = re.compile(r"<chNFe>\s*(\d{44})\s*</chNFe>", re.IGNORECASE)
PADRAO_INBOUND_FILENAME = re.compile(
    r"<(?:[\w.-]+:)?Inbound_Filename(?:\s[^>]*)?>\s*"
    r"(?:<!\[CDATA\[(.*?)\]\]>|([^<]+?))"
    r"\s*</(?:[\w.-]+:)?Inbound_Filename>",
    re.IGNORECASE | re.DOTALL,
)

LARGURA_VIEWPORT = 1980
ALTURA_VIEWPORT = 1080
TENTATIVAS_ELEMENTO = 3
PAUSA_RETRY_ELEMENTO = 1.5
LIMITE_SEND_KEYS = 80

EMAIL_DESTINO = os.environ.get("EMAIL_DESTINO", "eduardojuan@dhl.com").strip()
GMAIL_REMETENTE = (
    os.environ.get("GMAIL_USER", "ejuan36616@gmail.com").strip()
    or "ejuan36616@gmail.com"
)
TABELA_XMLS = os.environ.get("SUPABASE_TABELA_XMLS", "xmls_dhl_link").strip() or "xmls_dhl_link"

_BASE = Path(os.environ.get("GITHUB_WORKSPACE") or os.getcwd()).resolve()
PASTA_SAIDA = Path(os.environ.get("RPA_PASTA_DESTINO") or (_BASE / "saida_dhl_link"))
PASTA_DOWNLOADS = PASTA_SAIDA / "_downloads_chrome"
PASTA_XML_LOCAL = PASTA_SAIDA / "xml"
PASTA_PDF_LOCAL = PASTA_SAIDA / "pdf"


def _log(mensagem: str, nivel: str = "info") -> None:
    hora = datetime.now().strftime("%H:%M:%S")
    icones = {"etapa": "▶", "info": "●", "ok": "✓", "aviso": "!", "erro": "✗"}
    print(f"[{hora}] {icones.get(nivel, '●')} {mensagem}", flush=True)


def status(mensagem: str) -> None:
    _log(mensagem, "info")


def status_lote(mensagem: str) -> None:
    _log(mensagem, "info")


def _credencial(*nomes: str) -> str:
    for nome in nomes:
        valor = (os.environ.get(nome) or "").strip()
        if valor:
            return valor
    raise RuntimeError("Nenhuma das variáveis foi definida: " + ", ".join(nomes))


# ---------------------------------------------------------------------------
# Chrome
# ---------------------------------------------------------------------------

def forcar_resolucao(driver) -> None:
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


def _criar_opcoes_chrome() -> Options:
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
    destino = str(PASTA_DOWNLOADS)
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
    opcoes.add_experimental_option(
        "prefs",
        {
            "download.default_directory": destino,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0,
        },
    )
    return opcoes


def criar_driver():
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
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
    pasta = str(PASTA_DOWNLOADS)
    for comando in ("Browser.setDownloadBehavior", "Page.setDownloadBehavior"):
        try:
            driver.execute_cdp_cmd(
                comando,
                {"behavior": "allow", "downloadPath": pasta, "eventsEnabled": True},
            )
            break
        except Exception:
            continue
    forcar_resolucao(driver)
    return driver


def encerrar_driver(driver) -> None:
    if driver is None:
        return
    try:
        driver.quit()
    except Exception:
        pass


def _pagina_carregou(driver) -> bool:
    try:
        url = (driver.current_url or "").strip()
        ready = driver.execute_script("return document.readyState") or ""
        return bool(url) and url != "data:," and ready in ("interactive", "complete")
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Login Microsoft (sempre e-mail + senha; não é erro)
# ---------------------------------------------------------------------------

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
            loc = (By.XPATH, seletor) if seletor.startswith("//") else (By.CSS_SELECTOR, seletor)
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
    wait = WebDriverWait(driver, 25)
    _log(f"Login Microsoft com e-mail e senha ({usuario})...", "etapa")

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

    if _clicar_primeiro(driver, wait, ["#idBtn_Back", "#idSIButton9"], timeout=8):
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


def acessar_home(driver, wait_login, wait) -> None:
    usuario = os.environ.get("MS_USER", MS_USER_PADRAO).strip() or MS_USER_PADRAO
    senha = _credencial("MS_PASSWORD", "OUTLOOK_PASSWORD", "DHL_LINK_PASSWORD")

    _log(f"Abrindo {URL_DHL_LINK}...", "etapa")
    driver.get(URL_DHL_LINK)
    forcar_resolucao(driver)
    wait.until(lambda d: _pagina_carregou(d))
    time.sleep(2)

    ja_na_home = bool(driver.find_elements(By.XPATH, XPATH_HOME))
    if not ja_na_home:
        _login_microsoft(driver, usuario, senha)

    wait_login.until(EC.presence_of_element_located((By.XPATH, XPATH_HOME)))
    forcar_resolucao(driver)
    time.sleep(2)
    _log("Home do DHL Link detectada.", "ok")


# ---------------------------------------------------------------------------
# Helpers Selenium (fluxo do DHL Link)
# ---------------------------------------------------------------------------

def clicar(driver, elemento) -> None:
    try:
        elemento.click()
    except ElementClickInterceptedException:
        driver.execute_script(
            "arguments[0].dispatchEvent(new MouseEvent('click', {bubbles: true}));",
            elemento,
        )
    time.sleep(1)


def com_retry_stale(acao, tentativas=TENTATIVAS_ELEMENTO, pausa=PAUSA_RETRY_ELEMENTO, descricao="ação"):
    ultimo = None
    for i in range(1, tentativas + 1):
        try:
            return acao()
        except (TimeoutException, StaleElementReferenceException) as e:
            ultimo = e
            if i < tentativas:
                motivo = "não encontrado" if isinstance(e, TimeoutException) else "stale"
                _log(
                    f"Elemento '{descricao}' {motivo} "
                    f"(tentativa {i}/{tentativas}); tentando de novo...",
                    "aviso",
                )
                time.sleep(pausa)
    tipo = type(ultimo).__name__ if ultimo else "erro"
    raise TimeoutException(
        f"Falha ao localizar/usar '{descricao}' após {tentativas} tentativas ({tipo})."
    ) from ultimo


def esperar_xpath(wait, xpath, *, clicavel=False, todos=False, descricao="elemento"):
    def _acao():
        if todos:
            return wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
        cond = EC.element_to_be_clickable if clicavel else EC.presence_of_element_located
        return wait.until(cond((By.XPATH, xpath)))

    return com_retry_stale(_acao, descricao=descricao)


def clicar_xpath(wait, driver, xpath, *, presence=False, scroll=False, descricao="clique"):
    def _acao():
        cond = EC.presence_of_element_located if presence else EC.element_to_be_clickable
        el = wait.until(cond((By.XPATH, xpath)))
        if scroll:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
            time.sleep(0.3)
        clicar(driver, el)

    com_retry_stale(_acao, descricao=descricao)


def limpar_campo(driver, elemento) -> None:
    clicar(driver, elemento)
    elemento.send_keys(Keys.CONTROL, "a")
    elemento.send_keys(Keys.DELETE)


def _valor_campo(driver, elemento) -> str:
    return driver.execute_script("return arguments[0].value || '';", elemento) or ""


def inserir_texto_rapido(driver, elemento, texto: str) -> None:
    if not texto:
        return
    driver.execute_script("arguments[0].focus(); arguments[0].select();", elemento)
    driver.execute_script(
        """
        const el = arguments[0];
        const valor = arguments[1];
        el.focus();
        el.select();
        document.execCommand('insertText', false, valor);
        """,
        elemento,
        texto,
    )
    if _valor_campo(driver, elemento) == texto:
        return
    try:
        driver.execute_cdp_cmd(
            "Browser.grantPermissions",
            {
                "origin": URL_DHL_LINK,
                "permissions": ["clipboardReadWrite", "clipboardSanitizedWrite"],
            },
        )
        driver.execute_script("return navigator.clipboard.writeText(arguments[0]);", texto)
        elemento.send_keys(Keys.CONTROL, "v")
    except Exception:
        pass
    if _valor_campo(driver, elemento) == texto:
        return
    raise TimeoutException(
        f"Não consegui colar o texto longo no campo de busca ({len(texto)} caracteres)."
    )


def limpar_e_digitar(wait, driver, xpath, texto, descricao="digitar"):
    def _acao():
        el = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        limpar_campo(driver, el)
        el = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        if not texto:
            return
        if len(texto) > LIMITE_SEND_KEYS:
            inserir_texto_rapido(driver, el, texto)
        else:
            el.send_keys(texto)

    com_retry_stale(_acao, descricao=descricao)


def dispensar_aviso_dismiss(driver, timeout=2) -> None:
    xpath = "//button[normalize-space(text())='Dismiss']"
    try:
        WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    except TimeoutException:
        return
    _log("Aviso 'Dismiss' encontrado; fechando...", "info")
    try:
        clicar_xpath(WebDriverWait(driver, timeout), driver, xpath, descricao="Dismiss")
    except Exception:
        pass


def clicar_botao(wait, driver, texto: str) -> None:
    if texto != "Dismiss":
        dispensar_aviso_dismiss(driver)
    xpath = f"//button[normalize-space(text())='{texto}']"

    def _acao():
        try:
            botao = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        except TimeoutException:
            botoes_visiveis = [
                b.text.strip()
                for b in driver.find_elements(By.TAG_NAME, "button")
                if b.text.strip()
            ]
            raise TimeoutException(
                f"Não encontrei o botão '{texto}'. "
                f"Botões visíveis: {botoes_visiveis or 'nenhum'}"
            )
        clicar(driver, botao)

    com_retry_stale(_acao, descricao=f"botão '{texto}'")


def remover_selecionados(driver) -> None:
    for _ in range(30):
        icones = driver.find_elements(By.XPATH, "//*[@data-testid='CancelIcon']")
        if not icones:
            break
        try:
            icones[0].click()
        except StaleElementReferenceException:
            continue
        except Exception:
            try:
                driver.execute_script(
                    "arguments[0].dispatchEvent(new MouseEvent('click', {bubbles: true}));",
                    icones[0],
                )
            except StaleElementReferenceException:
                continue
        time.sleep(0.3)
    else:
        _log("Aviso: ainda havia chips após 30 tentativas de remoção.", "aviso")
        return
    _log("Seleção limpa (nenhum chip restante).", "ok")


def buscar_e_selecionar_flow(wait, driver, valor_busca: str) -> None:
    xpath_item_lista = (
        "//div[contains(@class,'flowSelectRowContent')]"
        f"[contains(normalize-space(.), '{valor_busca}')]"
    )
    for tentativa in range(1, TENTATIVAS_ELEMENTO + 1):
        try:
            _log(f"Buscando Flow '{valor_busca}'...", "info")
            limpar_e_digitar(
                wait,
                driver,
                XPATH_INPUT_BUSCA_FLOW,
                valor_busca,
                descricao=f"busca Flow '{valor_busca}'",
            )
            time.sleep(1.5)
            try:
                item = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_item_lista)))
            except TimeoutException:
                raise TimeoutException(
                    f"Não encontrei o Flow '{valor_busca}' na lista."
                )
            clicar(driver, item)
            time.sleep(0.5)
            return
        except (TimeoutException, StaleElementReferenceException):
            if tentativa < TENTATIVAS_ELEMENTO:
                _log(
                    f"Não consegui selecionar o Flow '{valor_busca}' "
                    f"(tentativa {tentativa}/{TENTATIVAS_ELEMENTO}).",
                    "aviso",
                )
                time.sleep(PAUSA_RETRY_ELEMENTO)
    raise TimeoutException(
        f"Não consegui buscar/selecionar '{valor_busca}' após {TENTATIVAS_ELEMENTO} tentativas."
    )


def selecionar_partner_e_flow(driver, wait) -> None:
    _log("Abrindo o seletor de Partner Location / Flow...", "etapa")
    xpath_card = "//a[.//*[@data-testid='LocationOnOutlinedIcon']]"
    try:
        clicar_xpath(wait, driver, xpath_card, presence=True, descricao="card Partner/Flow")
    except TimeoutException:
        raise TimeoutException("Não encontrei o card de Partner Location / Flow na tela.")
    remover_selecionados(driver)
    for valor_busca in FLOWS_BUSCA:
        buscar_e_selecionar_flow(wait, driver, valor_busca)
        clicar_botao(wait, driver, "Select")
    clicar_botao(wait, driver, "Confirm")
    _log("Flows (XML + PDF) confirmados.", "ok")


def filtrar_periodo(driver, wait) -> None:
    _log("Abrindo o filtro de datas...", "etapa")
    xpath = "//*[@data-testid='DateRangeTwoToneIcon']"
    try:
        clicar_xpath(wait, driver, xpath, presence=True, descricao="ícone calendário")
    except TimeoutException:
        raise TimeoutException("Não encontrei o ícone de calendário na tela.")

    hoje = date.today()
    data_from = (hoje - timedelta(days=DIAS_PARA_TRAS_NO_FILTRO)).strftime(FORMATO_DATA_FILTRO)
    data_to = (hoje + timedelta(days=DIAS_PARA_FRENTE_NO_FILTRO)).strftime(FORMATO_DATA_FILTRO)
    _log(f"Preenchendo período: {data_from} até {data_to}...", "info")

    def _acao():
        campos = wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_CAMPOS_DATA_FILTRO)))
        if len(campos) < 2:
            raise TimeoutException(
                f"Esperava 2 campos de data (From/To), encontrei {len(campos)}."
            )
        limpar_campo(driver, campos[0])
        campos = wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_CAMPOS_DATA_FILTRO)))
        campos[0].send_keys(data_from)
        limpar_campo(driver, campos[1])
        campos = wait.until(EC.presence_of_all_elements_located((By.XPATH, XPATH_CAMPOS_DATA_FILTRO)))
        if len(campos) < 2:
            raise StaleElementReferenceException("Campos de data recriados durante o preenchimento.")
        campos[1].send_keys(data_to)

    com_retry_stale(_acao, descricao="preencher From/To")
    time.sleep(0.5)
    clicar_botao(wait, driver, "Apply")
    _log("Período aplicado.", "ok")


def clicar_lupa(wait, driver) -> None:
    xpath = "//*[@data-testid='SearchIcon']"
    try:
        clicar_xpath(wait, driver, xpath, descricao="lupa SearchIcon")
    except TimeoutException:
        raise TimeoutException("Não encontrei o ícone de busca (lupa) na tela.")


def pesquisar_mensagens(wait, driver) -> None:
    _log("Limpando o campo de busca de mensagens...", "info")
    limpar_e_digitar(
        wait, driver, XPATH_CAMPO_BUSCA_MENSAGENS, "", descricao="limpar busca mensagens"
    )
    clicar_lupa(wait, driver)
    _log("Pesquisa disparada.", "ok")


def selecionar_tamanho_pagina(wait, driver, valor: str) -> None:
    _log(f"Selecionando tamanho de página '{valor}'...", "info")
    xpath_opcao = f"//li[@role='option' and @data-value='{valor}']"
    esperar_xpath(
        wait, XPATH_SELECT_TAMANHO_PAGINA, clicavel=True, descricao="seletor de tamanho de página"
    )
    for tentativa, estrategia in enumerate(["clique", "action_chains", "teclado"], start=1):
        try:
            select = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_SELECT_TAMANHO_PAGINA)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select)
            time.sleep(0.2)
            if estrategia == "clique":
                clicar(driver, select)
            elif estrategia == "action_chains":
                ActionChains(driver).move_to_element(select).pause(0.2).click().perform()
                time.sleep(1)
            else:
                select.send_keys(Keys.ENTER)
                time.sleep(1)
            opcao = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xpath_opcao)))
            clicar(driver, opcao)
            return
        except StaleElementReferenceException:
            _log(f"Seletor stale na estratégia '{estrategia}' ({tentativa}/3).", "aviso")
            time.sleep(PAUSA_RETRY_ELEMENTO)
            continue
        except TimeoutException:
            _log(f"Menu não abriu com '{estrategia}' ({tentativa}/3).", "aviso")
    opcoes_visiveis = [
        li.text.strip()
        for li in driver.find_elements(By.XPATH, "//li[@role='option']")
        if li.text.strip()
    ]
    raise TimeoutException(
        "Não consegui abrir o menu de tamanho de página. "
        f"Opções visíveis: {opcoes_visiveis or 'nenhuma'}"
    )


def marcar_selecionar_tudo(wait, driver) -> None:
    _log("Marcando 'selecionar tudo' na lista...", "info")
    xpath = "//span[contains(@class,'MuiCheckbox-root')][.//input[@aria-label='parent-checkbox']]"
    try:
        clicar_xpath(wait, driver, xpath, descricao="selecionar tudo")
    except TimeoutException:
        raise TimeoutException("Não encontrei o checkbox de 'selecionar tudo'.")


def clicar_menu_mais_opcoes(wait, driver) -> None:
    xpath = "//button[.//*[@data-testid='MoreHorizIcon']]"
    try:
        clicar_xpath(wait, driver, xpath, presence=True, scroll=True, descricao="mais opções FAB")
    except TimeoutException:
        raise TimeoutException("Não encontrei o botão de 'mais opções' (FAB).")


def clicar_exportar_download(wait, driver) -> None:
    xpath = "//button[.//*[@data-testid='GetAppIcon']]"
    try:
        clicar_xpath(wait, driver, xpath, descricao="Export to File")
    except TimeoutException:
        raise TimeoutException("Não encontrei o botão 'Export to File'.")


def selecionar_formato_xml(wait, driver) -> None:
    xpath = "//li[@role='menuitem'][.//p[normalize-space(text())='XML']]"
    try:
        clicar_xpath(wait, driver, xpath, descricao="opção XML")
    except TimeoutException:
        raise TimeoutException("Não encontrei a opção 'XML' na lista de exportação.")


def selecionar_opcao_in_files(wait, driver) -> None:
    xpath = "//li[@role='menuitem'][.//p[normalize-space(text())='IN Files']]"
    try:
        clicar_xpath(wait, driver, xpath, descricao="opção IN Files")
    except TimeoutException:
        raise TimeoutException("Não encontrei a opção 'IN Files'.")


# ---------------------------------------------------------------------------
# Downloads / arquivos
# ---------------------------------------------------------------------------

def _snapshot_pasta(pasta) -> dict:
    snap = {}
    try:
        nomes = os.listdir(pasta)
    except OSError:
        return snap
    for nome in nomes:
        caminho = os.path.join(str(pasta), nome)
        try:
            st = os.stat(caminho)
            snap[nome] = (st.st_mtime_ns, st.st_size)
        except OSError:
            continue
    return snap


def _arquivo_download_candidato(nome, extensoes_permitidas=None) -> bool:
    nome_lower = nome.lower()
    if nome_lower.endswith((".crdownload", ".tmp", ".htm", ".html", ".part")):
        return False
    if nome.startswith("~$"):
        return False
    if extensoes_permitidas:
        return nome_lower.endswith(tuple(e.lower() for e in extensoes_permitidas))
    return True


def _ainda_baixando(pasta_downloads) -> bool:
    try:
        nomes = os.listdir(pasta_downloads)
    except OSError:
        return False
    return any(n.lower().endswith((".crdownload", ".tmp", ".part")) for n in nomes)


def _zip_valido(caminho) -> bool:
    try:
        if not caminho or not os.path.isfile(caminho):
            return False
        if os.path.getsize(caminho) <= 0:
            return False
        return zipfile.is_zipfile(caminho)
    except OSError:
        return False


def esperar_download_concluir(
    pasta_downloads,
    snapshot_antes,
    timeout=60,
    extensoes_permitidas=(".xml",),
    validar_zip=False,
):
    _log("Aguardando o download terminar...", "info")
    pasta = str(pasta_downloads)
    fim = time.time() + timeout
    while time.time() < fim:
        if _ainda_baixando(pasta):
            time.sleep(0.5)
            continue
        atuais = _snapshot_pasta(pasta)
        candidatos = []
        for nome, meta in atuais.items():
            if not _arquivo_download_candidato(nome, extensoes_permitidas):
                continue
            if nome not in snapshot_antes or snapshot_antes[nome] != meta:
                candidatos.append(nome)
        if candidatos:
            time.sleep(1)
            estaveis = []
            for nome in candidatos:
                caminho = os.path.join(pasta, nome)
                try:
                    tamanho1 = os.path.getsize(caminho)
                    time.sleep(0.5)
                    tamanho2 = os.path.getsize(caminho)
                except OSError:
                    continue
                if tamanho1 == tamanho2 and tamanho2 > 0:
                    if validar_zip and not _zip_valido(caminho):
                        continue
                    estaveis.append(nome)
            if estaveis:
                estaveis.sort(
                    key=lambda n: os.path.getmtime(os.path.join(pasta, n)),
                    reverse=True,
                )
                caminho = os.path.join(pasta, estaveis[0])
                _log(f"Download concluído: {os.path.basename(caminho)}", "ok")
                return caminho
        time.sleep(0.5)
    raise TimeoutException(
        f"Nenhum arquivo novo {extensoes_permitidas} apareceu em "
        f"'{pasta}' após {timeout}s."
    )


def _mover_substituindo(origem, destino):
    pasta = os.path.dirname(destino)
    if pasta:
        os.makedirs(pasta, exist_ok=True)
    if os.path.exists(destino):
        try:
            os.remove(destino)
        except OSError:
            base, ext = os.path.splitext(destino)
            i = 1
            while True:
                alt = f"{base}_{i}{ext}"
                if not os.path.exists(alt):
                    destino = alt
                    break
                i += 1
    shutil.move(origem, destino)
    return destino


def limpar_extensoes_invalidas(pasta, extensao_permitida) -> int:
    if not os.path.isdir(pasta):
        return 0
    extensao = extensao_permitida.lower()
    if not extensao.startswith("."):
        extensao = f".{extensao}"
    removidos = 0
    for nome in os.listdir(pasta):
        caminho = os.path.join(str(pasta), nome)
        if not os.path.isfile(caminho):
            continue
        if nome.lower().endswith(extensao):
            continue
        try:
            os.remove(caminho)
            removidos += 1
        except OSError as e:
            _log(f"Aviso: não consegui apagar '{caminho}': {e}", "aviso")
    return removidos


def extrair_e_mover_zip(caminho_zip, pasta_xml=None, pasta_pdf=None):
    pasta_xml = str(pasta_xml or PASTA_XML_LOCAL)
    pasta_pdf = str(pasta_pdf or PASTA_PDF_LOCAL)
    if not _zip_valido(caminho_zip):
        raise ValueError(f"Arquivo baixado não é um ZIP válido: '{caminho_zip}'.")

    _log(f"Descompactando {os.path.basename(caminho_zip)}...", "etapa")
    pasta_extracao = (
        caminho_zip[:-4] if caminho_zip.lower().endswith(".zip") else f"{caminho_zip}_extraido"
    )
    if os.path.isdir(pasta_extracao):
        shutil.rmtree(pasta_extracao, ignore_errors=True)

    try:
        with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
            ruim = zip_ref.testzip()
            if ruim is not None:
                raise zipfile.BadZipFile(
                    f"ZIP corrompido (arquivo interno: {ruim}): {caminho_zip}"
                )
            zip_ref.extractall(pasta_extracao)
    except zipfile.BadZipFile as e:
        raise ValueError(f"Falha ao abrir o ZIP '{caminho_zip}': {e}") from e

    os.makedirs(pasta_xml, exist_ok=True)
    os.makedirs(pasta_pdf, exist_ok=True)
    movidos_xml = []
    movidos_pdf = []
    ignorados = 0
    for raiz, _dirs, arquivos in os.walk(pasta_extracao):
        for nome_arquivo in arquivos:
            origem = os.path.join(raiz, nome_arquivo)
            nome_lower = nome_arquivo.lower()
            if nome_lower.endswith(".xml"):
                movidos_xml.append(
                    _mover_substituindo(origem, os.path.join(pasta_xml, nome_arquivo))
                )
            elif nome_lower.endswith(".pdf"):
                movidos_pdf.append(
                    _mover_substituindo(origem, os.path.join(pasta_pdf, nome_arquivo))
                )
            else:
                ignorados += 1

    shutil.rmtree(pasta_extracao, ignore_errors=True)
    try:
        os.remove(caminho_zip)
    except OSError as e:
        _log(f"Aviso: não consegui apagar o ZIP '{caminho_zip}': {e}", "aviso")

    _log(
        f"{len(movidos_xml)} XML(s) | {len(movidos_pdf)} PDF(s)"
        + (f" | {ignorados} ignorado(s)" if ignorados else ""),
        "ok",
    )
    limpar_extensoes_invalidas(pasta_xml, ".xml")
    limpar_extensoes_invalidas(pasta_pdf, ".pdf")
    return movidos_xml, movidos_pdf


# ---------------------------------------------------------------------------
# Identificadores / confronto
# ---------------------------------------------------------------------------

def normalizar_nf(valor):
    digitos = re.sub(r"\D", "", str(valor or ""))
    if not digitos:
        return None
    return digitos.zfill(9)[-9:]


def nf_da_chave_acesso(chave):
    digitos = re.sub(r"\D", "", str(chave or ""))
    if len(digitos) != 44:
        return None
    return digitos[25:34]


def chave_nfe_do_nome(nome_arquivo):
    nome = os.path.basename(str(nome_arquivo or "")).strip()
    m = PADRAO_NFE_ARQUIVO.search(nome)
    if m:
        return m.group(1)
    m = PADRAO_CHAVE_AVULSA.search(nome)
    if m:
        return m.group(1)
    return None


def termo_busca_do_nome(nome_arquivo):
    nome = os.path.basename(str(nome_arquivo or "")).strip()
    pedidos = [m.upper() for m in PADRAO_PEDIDO.findall(nome)]
    if pedidos:
        return pedidos[0]
    chave = chave_nfe_do_nome(nome)
    if chave:
        return f"NFE-{chave}"
    return None


def chaves_presenca_do_nome(nome_arquivo, conteudo=None):
    chaves = set()
    nome = os.path.basename(str(nome_arquivo or "")).strip()
    if nome:
        chaves.add(f"A:{nome.lower()}")
        stem = os.path.splitext(nome)[0].lower()
        if stem:
            chaves.add(f"S:{stem}")
    for m in PADRAO_PEDIDO.findall(nome):
        chaves.add(f"P:{m.upper()}")
    chave = chave_nfe_do_nome(nome)
    if chave:
        chaves.add(f"C:{chave}")
        nf = nf_da_chave_acesso(chave)
        if nf:
            chaves.add(f"N:{nf}")
    if conteudo:
        for m in PADRAO_PEDIDO.findall(conteudo):
            chaves.add(f"P:{m.upper()}")
        for nf_bruto in PADRAO_NNF_XML.findall(conteudo):
            nf = normalizar_nf(nf_bruto)
            if nf:
                chaves.add(f"N:{nf}")
        for chave_xml in PADRAO_CHNFE_XML.findall(conteudo):
            chaves.add(f"C:{chave_xml}")
            nf = nf_da_chave_acesso(chave_xml)
            if nf:
                chaves.add(f"N:{nf}")
    return chaves


def _texto_tag_xml(conteudo, tag):
    m = re.search(
        rf"<(?:[\w.-]+:)?{re.escape(tag)}(?:\s[^>]*)?>\s*([^<]*?)\s*</(?:[\w.-]+:)?{re.escape(tag)}>",
        conteudo or "",
        flags=re.IGNORECASE,
    )
    return (m.group(1).strip() if m else "") or ""


def _parse_data_emissao(dh_emi):
    if not dh_emi:
        return None
    texto = str(dh_emi).strip()
    try:
        return datetime.fromisoformat(texto).astimezone(timezone.utc).isoformat()
    except ValueError:
        pass
    try:
        base = texto[:19]
        dt = datetime.strptime(base, "%Y-%m-%dT%H:%M:%S")
        return dt.replace(tzinfo=timezone.utc).isoformat()
    except ValueError:
        return None


def _parse_valor(valor):
    try:
        return round(float(str(valor).replace(",", ".")), 2)
    except (TypeError, ValueError):
        return None


def extrair_dados_xml_arquivo(caminho: str) -> dict:
    nome = os.path.basename(caminho)
    conteudo = ""
    try:
        with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()
    except OSError as e:
        return {"arquivo": nome, "erro": str(e)}

    pedido = ""
    m_ped = PADRAO_PEDIDO.search(nome) or PADRAO_PEDIDO.search(conteudo)
    if m_ped:
        pedido = m_ped.group(0).upper()

    chave = ""
    m_nfe = PADRAO_NFE_ARQUIVO.search(nome)
    if m_nfe:
        chave = m_nfe.group(1)
    if not chave:
        m_ch = PADRAO_CHNFE_XML.search(conteudo)
        if m_ch:
            chave = m_ch.group(1)

    nf = normalizar_nf(_texto_tag_xml(conteudo, "nNF")) or ""
    if not nf and chave:
        nf = nf_da_chave_acesso(chave) or ""

    uf = _texto_tag_xml(conteudo, "UF")
    m_uf_dest = re.search(
        r"<(?:[\w.-]+:)?enderDest\b[\s\S]*?<(?:[\w.-]+:)?UF(?:\s[^>]*)?>\s*([^<]+)\s*</(?:[\w.-]+:)?UF>",
        conteudo,
        flags=re.IGNORECASE,
    )
    if m_uf_dest:
        uf = m_uf_dest.group(1).strip()

    destinatario = _texto_tag_xml(conteudo, "xNome")
    m_dest = re.search(
        r"<(?:[\w.-]+:)?dest\b[\s\S]*?<(?:[\w.-]+:)?xNome(?:\s[^>]*)?>\s*([^<]+)\s*</(?:[\w.-]+:)?xNome>",
        conteudo,
        flags=re.IGNORECASE,
    )
    if m_dest:
        destinatario = m_dest.group(1).strip()

    return {
        "arquivo": nome,
        "pedido": pedido or None,
        "nf": nf or None,
        "chave_acesso": chave or None,
        "uf": uf or None,
        "destinatario": destinatario or None,
        "valor": _parse_valor(_texto_tag_xml(conteudo, "vNF")),
        "data_emissao": _parse_data_emissao(_texto_tag_xml(conteudo, "dhEmi")),
    }


def extrair_entradas_do_relatorio(caminho_arquivo: str):
    with open(caminho_arquivo, "r", encoding="utf-8", errors="ignore") as f:
        conteudo = f.read()

    try:
        PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
        shutil.copy2(caminho_arquivo, str(PASTA_SAIDA / "ultimo_relatorio_messages.xml"))
    except OSError as e:
        _log(f"Aviso: não consegui guardar cópia do relatório: {e}", "aviso")

    nomes = []
    for m in PADRAO_INBOUND_FILENAME.finditer(conteudo):
        nome = (m.group(1) or m.group(2) or "").strip()
        if nome:
            nomes.append(nome)
    if not nomes:
        nomes = [f"NFE-{c}.xml" for c in PADRAO_NFE_ARQUIVO.findall(conteudo)]
        nomes += PADRAO_PEDIDO.findall(conteudo)

    qtd_xml = sum(1 for n in nomes if os.path.basename(str(n)).lower().endswith(".xml"))
    qtd_pdf = sum(1 for n in nomes if os.path.basename(str(n)).lower().endswith(".pdf"))
    _log(
        f"Inbound no relatório: {len(nomes)} (xml: {qtd_xml}, pdf: {qtd_pdf})",
        "info",
    )

    entradas = []
    vistos = set()
    sem_termo = 0
    for nome in nomes:
        termo = termo_busca_do_nome(nome)
        chaves = chaves_presenca_do_nome(nome)
        if not termo or not chaves:
            sem_termo += 1
            continue
        if termo in vistos:
            continue
        vistos.add(termo)
        entradas.append((termo, chaves))
    if sem_termo:
        _log(f"Aviso: {sem_termo} inbound(s) sem QBQ/chave no nome.", "aviso")
    return entradas


def montar_pedidos_faltantes(caminho_relatorio: str, nomes_ja_baixados: set[str]):
    _log("Confrontando relatório com a tabela xmls_dhl_link no Supabase...", "etapa")
    entradas = extrair_entradas_do_relatorio(caminho_relatorio)
    existentes = set()
    for nome in nomes_ja_baixados:
        existentes |= chaves_presenca_do_nome(nome)

    faltantes = []
    for termo, chaves in entradas:
        if chaves & existentes:
            continue
        faltantes.append(termo)

    qtd_pedido = sum(1 for t, _ in entradas if t.upper().startswith("QBQ"))
    _log(
        f"Itens únicos no relatório: {len(entradas)} "
        f"(pedido: {qtd_pedido}, só NF: {len(entradas) - qtd_pedido}) | "
        f"já no Supabase: {len(nomes_ja_baixados)} | faltantes: {len(faltantes)}",
        "info",
    )
    return faltantes


def dividir_em_lotes(itens, tamanho):
    return [itens[i : i + tamanho] for i in range(0, len(itens), tamanho)]


# ---------------------------------------------------------------------------
# Supabase
# ---------------------------------------------------------------------------

def _supabase_config() -> tuple[str, str, str]:
    url = (os.environ.get("SUPABASE_URL") or "").strip().rstrip("/")
    key = _credencial("SUPABASE_SERVICE_ROLE_KEY", "SUPABASE_KEY")
    if not url:
        raise RuntimeError("SUPABASE_URL não definida.")
    return url, key, TABELA_XMLS


def _supabase_headers(key: str, prefer: str | None = None) -> dict:
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


def _http_json(method: str, url: str, headers: dict, payload=None) -> str:
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        detalhe = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Supabase HTTP {e.code} em {url}: {detalhe}") from e


def listar_arquivos_supabase() -> set[str]:
    url, key, tabela = _supabase_config()
    arquivos: set[str] = set()
    offset = 0
    limite = 1000
    while True:
        qs = urllib.parse.urlencode(
            {"select": "arquivo", "order": "arquivo", "offset": offset, "limit": limite}
        )
        body = _http_json(
            "GET",
            f"{url}/rest/v1/{tabela}?{qs}",
            _supabase_headers(key),
        )
        rows = json.loads(body or "[]")
        for row in rows:
            nome = (row.get("arquivo") or "").strip()
            if nome:
                arquivos.add(nome)
        if len(rows) < limite:
            break
        offset += limite
    _log(f"Supabase: {len(arquivos)} XML(s) já registrados em {tabela}.", "ok")
    return arquivos


def registrar_xmls_supabase(caminhos_xml: list[str], pdfs_por_stem: dict[str, str]) -> int:
    if not caminhos_xml:
        return 0
    url, key, tabela = _supabase_config()
    agora = datetime.now(timezone.utc).isoformat()
    payload = []
    for caminho in caminhos_xml:
        dados = extrair_dados_xml_arquivo(caminho)
        nome = dados.get("arquivo") or os.path.basename(caminho)
        stem = os.path.splitext(nome)[0].lower()
        pdf = pdfs_por_stem.get(stem)
        payload.append(
            {
                "arquivo": nome,
                "arquivo_pdf": os.path.basename(pdf) if pdf else None,
                "pedido": dados.get("pedido"),
                "nf": dados.get("nf"),
                "chave_acesso": dados.get("chave_acesso"),
                "uf": dados.get("uf"),
                "destinatario": dados.get("destinatario"),
                "valor": dados.get("valor"),
                "data_emissao": dados.get("data_emissao"),
                "data_download": agora,
                "enviado_em": agora,
                "email_destino": EMAIL_DESTINO,
            }
        )

    _http_json(
        "POST",
        f"{url}/rest/v1/{tabela}?on_conflict=arquivo",
        _supabase_headers(key, prefer="resolution=merge-duplicates,return=minimal"),
        payload,
    )
    _log(f"Supabase: {len(payload)} XML(s) gravado(s) em {tabela}.", "ok")
    return len(payload)


# ---------------------------------------------------------------------------
# E-mail
# ---------------------------------------------------------------------------

def _mime_anexo(path: Path) -> tuple[str, str]:
    ext = path.suffix.lower()
    if ext == ".xml":
        return "application", "xml"
    if ext == ".pdf":
        return "application", "pdf"
    return "application", "octet-stream"


def _enviar_email_smtp(anexos: list[Path], assunto: str, corpo: str) -> None:
    remetente = GMAIL_REMETENTE or _credencial("GMAIL_USER")
    senha = _credencial("GMAIL_APP_PASSWORD", "GMAIL_PASSWORD").replace(" ", "")
    servidor = os.environ.get("SMTP_SERVER", "smtp.gmail.com").strip()
    porta = int(os.environ.get("SMTP_PORT", "587"))

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = remetente
    msg["To"] = EMAIL_DESTINO
    msg.set_content(corpo)
    for path in anexos:
        maintype, subtype = _mime_anexo(path)
        msg.add_attachment(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )

    contexto = ssl.create_default_context()
    try:
        with smtplib.SMTP(servidor, porta, timeout=60) as smtp:
            smtp.starttls(context=contexto)
            smtp.login(remetente, senha)
            smtp.send_message(msg)
    except smtplib.SMTPAuthenticationError as e:
        raise RuntimeError(
            "Gmail recusou o login SMTP. Confira o secret GMAIL_APP_PASSWORD."
        ) from e
    _log(f"E-mail enviado para {EMAIL_DESTINO} ({len(anexos)} anexo(s)).", "ok")


def _lotes_email(xmls: list[str], pdfs_por_stem: dict[str, str]) -> list[list[tuple[str, str | None]]]:
    lotes: list[list[tuple[str, str | None]]] = []
    atual: list[tuple[str, str | None]] = []
    tamanho = 0
    for xml in xmls:
        pdf = pdfs_por_stem.get(os.path.splitext(os.path.basename(xml))[0].lower())
        tam = os.path.getsize(xml)
        if pdf and os.path.isfile(pdf):
            tam += os.path.getsize(pdf)
        if atual and tamanho + tam > LIMITE_EMAIL_BYTES:
            lotes.append(atual)
            atual = []
            tamanho = 0
        atual.append((xml, pdf if pdf and os.path.isfile(pdf) else None))
        tamanho += tam
    if atual:
        lotes.append(atual)
    return lotes


def enviar_e_registrar(movidos_xml: list[str], movidos_pdf: list[str]) -> int:
    if not movidos_xml and not movidos_pdf:
        return 0
    pdfs_por_stem = {
        os.path.splitext(os.path.basename(p))[0].lower(): p for p in movidos_pdf
    }
    xmls = [p for p in movidos_xml if os.path.isfile(p)]
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    registrados = 0
    lotes = _lotes_email(xmls, pdfs_por_stem)
    pdfs_com_xml = set()
    for lote in lotes:
        for _xml, pdf in lote:
            if pdf:
                pdfs_com_xml.add(pdf)
    pdfs_soltos = [p for p in movidos_pdf if p not in pdfs_com_xml and os.path.isfile(p)]

    for i, lote in enumerate(lotes, start=1):
        anexos = []
        for xml, pdf in lote:
            anexos.append(Path(xml))
            if pdf:
                anexos.append(Path(pdf))
        if i == len(lotes):
            anexos.extend(Path(p) for p in pdfs_soltos)
            pdfs_soltos = []
        linhas = [
            f"DHL Link (GitHub Actions) — {agora}",
            "",
            f"Lote {i}/{max(len(lotes), 1)}: {len(lote)} XML(s) + PDFs correspondentes.",
            "",
            "Arquivos:",
        ]
        linhas.extend(f"  - {p.name}" for p in anexos)
        _enviar_email_smtp(
            anexos,
            f"DHL Link — XML/PDF {datetime.now().strftime('%d/%m/%Y %H:%M')} ({i}/{len(lotes)})",
            "\n".join(linhas),
        )
        registrados += registrar_xmls_supabase([xml for xml, _pdf in lote], pdfs_por_stem)

    if pdfs_soltos:
        anexos = [Path(p) for p in pdfs_soltos]
        _enviar_email_smtp(
            anexos,
            f"DHL Link — PDF {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            "PDFs sem XML correspondente neste lote:\n"
            + "\n".join(f"  - {p.name}" for p in anexos),
        )
    return registrados


# ---------------------------------------------------------------------------
# Lotes de faltantes
# ---------------------------------------------------------------------------

def pesquisar_pedidos_faltantes(wait, driver, pedidos_faltantes) -> int:
    if not pedidos_faltantes:
        _log("Nenhum item faltante — todos os XMLs do relatório já estão no Supabase.", "ok")
        return 0

    total_registrados = 0
    lotes = dividir_em_lotes(pedidos_faltantes, LIMITE_PEDIDOS_POR_BUSCA)
    for numero_lote, lote in enumerate(lotes, start=1):
        query = " OR ".join(lote)
        status_lote(f"Lote {numero_lote}/{len(lotes)}: pesquisando {len(lote)} item(ns)...")
        _log(f"Lote {numero_lote}/{len(lotes)} ({len(lote)} itens): {query}", "info")
        limpar_e_digitar(
            wait,
            driver,
            XPATH_CAMPO_BUSCA_MENSAGENS,
            query,
            descricao=f"busca lote {numero_lote}",
        )
        clicar_lupa(wait, driver)
        time.sleep(10)

        status_lote(f"Lote {numero_lote}/{len(lotes)}: selecionando e exportando...")
        marcar_selecionar_tudo(wait, driver)
        time.sleep(10)
        clicar_menu_mais_opcoes(wait, driver)

        snapshot_antes_download = _snapshot_pasta(PASTA_DOWNLOADS)
        clicar_exportar_download(wait, driver)
        selecionar_opcao_in_files(wait, driver)

        status_lote(f"Lote {numero_lote}/{len(lotes)}: aguardando ZIP...")
        dispensar_aviso_dismiss(driver, timeout=180)
        caminho_zip = esperar_download_concluir(
            PASTA_DOWNLOADS,
            snapshot_antes_download,
            timeout=180,
            extensoes_permitidas=(".zip",),
            validar_zip=True,
        )
        status_lote(f"Lote {numero_lote}/{len(lotes)}: extraindo, enviando e registrando...")
        movidos_xml, movidos_pdf = extrair_e_mover_zip(caminho_zip)
        total_registrados += enviar_e_registrar(movidos_xml, movidos_pdf)
        status_lote(f"Lote {numero_lote}/{len(lotes)}: concluído.")

    _log(f"Todos os {len(lotes)} lote(s) processados. XMLs gravados: {total_registrados}.", "ok")
    return total_registrados


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------

def baixar_xml_dhl_link() -> int:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
    PASTA_XML_LOCAL.mkdir(parents=True, exist_ok=True)
    PASTA_PDF_LOCAL.mkdir(parents=True, exist_ok=True)

    _log("Download XML DHL Link (GitHub Actions).", "etapa")
    nomes_ja_baixados = listar_arquivos_supabase()

    driver = None
    try:
        driver = criar_driver()
        wait_login = WebDriverWait(driver, 60)
        wait = WebDriverWait(driver, 15)

        acessar_home(driver, wait_login, wait)

        limpar_extensoes_invalidas(PASTA_XML_LOCAL, ".xml")
        limpar_extensoes_invalidas(PASTA_PDF_LOCAL, ".pdf")

        _log("Selecionando Partner/Flow (XML + PDF)...", "etapa")
        selecionar_partner_e_flow(driver, wait)

        _log("Aplicando filtro de período e pesquisando mensagens...", "etapa")
        filtrar_periodo(driver, wait)
        pesquisar_mensagens(wait, driver)
        selecionar_tamanho_pagina(wait, driver, TAMANHO_PAGINA_DESEJADO)
        clicar_lupa(wait, driver)
        time.sleep(10)
        marcar_selecionar_tudo(wait, driver)
        time.sleep(10)
        clicar_menu_mais_opcoes(wait, driver)

        _log("Exportando relatório para descobrir os pedidos...", "etapa")
        snapshot_antes_download = _snapshot_pasta(PASTA_DOWNLOADS)
        clicar_exportar_download(wait, driver)
        selecionar_formato_xml(wait, driver)
        caminho_relatorio = esperar_download_concluir(
            PASTA_DOWNLOADS,
            snapshot_antes_download,
            timeout=90,
        )

        pedidos_faltantes = montar_pedidos_faltantes(caminho_relatorio, nomes_ja_baixados)
        if pedidos_faltantes:
            _log("Itens faltantes: " + ", ".join(pedidos_faltantes), "info")

        _log(
            f"Processando {len(pedidos_faltantes)} item(ns) faltante(s) "
            f"em lotes de {LIMITE_PEDIDOS_POR_BUSCA}...",
            "etapa",
        )
        pesquisar_pedidos_faltantes(wait, driver, pedidos_faltantes)
        _log("Execução concluída com sucesso.", "ok")
        return 0
    except InvalidSessionIdException:
        _log("Navegador fechado antes do script terminar.", "erro")
        return 1
    except Exception as e:
        linhas = str(e).strip().splitlines()
        primeira = linhas[0] if linhas else type(e).__name__
        _log(f"Ocorreu um erro: {primeira or type(e).__name__}", "erro")
        print(f"RESULTADO: FALHOU — {e}")
        return 1
    finally:
        encerrar_driver(driver)


if __name__ == "__main__":
    sys.exit(baixar_xml_dhl_link())
