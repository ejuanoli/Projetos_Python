"""
RPA Apple para GitHub Actions.
Cópia do RPA_Apple.py sem interface gráfica, sem SharePoint, sem
consolidação/tratamento de Excel. Baixa os relatórios Brudam (embarque +
detalhado) e TSMS (xls) e envia os arquivos por e-mail Gmail para
eduardojuan@dhl.com.

Credenciais vêm de variáveis de ambiente (GitHub Secrets), nunca do código.
"""

from __future__ import annotations

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from datetime import date, timedelta, datetime
from email.message import EmailMessage
from pathlib import Path
import glob
import os
import shutil
import smtplib
import ssl
import sys
import time

URL_LOGIN = "https://dhl.brudam.com.br/index.php?acao=2"
URL_EMBARQUE = "https://dhl.brudam.com.br/gerencia/relatorio_embarque.php"
URL_LISTA_NOTAS = "https://dhl.brudam.com.br/sac/lista_notas.php"
URL_TSMS = "https://brasillogic.com.br/tsms/logtrack/index.php#"

EMAIL_DESTINO = os.environ.get("EMAIL_DESTINO", "eduardojuan@dhl.com").strip()
GMAIL_REMETENTE = os.environ.get("GMAIL_USER", "ejuan36616@gmail.com").strip() or "ejuan36616@gmail.com"

_BASE = Path(os.environ.get("GITHUB_WORKSPACE") or os.getcwd()).resolve()
PASTA_DESTINO = Path(os.environ.get("RPA_PASTA_DESTINO") or (_BASE / "saida_rpa_apple"))
PASTA_DOWNLOADS = PASTA_DESTINO / "_downloads_chrome"


def _log(mensagem: str, nivel: str = "info") -> None:
    hora = datetime.now().strftime("%H:%M:%S")
    icones = {"etapa": "▶", "info": "●", "ok": "✓", "aviso": "!", "erro": "✗"}
    print(f"[{hora}] {icones.get(nivel, '●')} {mensagem}", flush=True)


def _credencial(*nomes: str) -> str:
    for nome in nomes:
        valor = (os.environ.get(nome) or "").strip()
        if valor:
            return valor
    raise RuntimeError(
        "Nenhuma das variáveis foi definida: " + ", ".join(nomes)
    )


def _navegador_fechado(exc: BaseException) -> bool:
    texto = str(exc).lower()
    if isinstance(exc, WebDriverException) and "gethandleverifier" in texto:
        return True
    termos = (
        "gethandleverifier",
        "chrome not reachable",
        "disconnected",
        "invalid session id",
        "no such window",
        "target window already closed",
        "session deleted",
        "not connected to devtools",
        "browser has closed",
    )
    return any(t in texto for t in termos)


def _encerrar_driver(driver) -> None:
    if driver is None:
        return
    try:
        driver.quit()
    except Exception:
        try:
            driver.service.stop()
        except Exception:
            pass


def _criar_opcoes_chrome() -> Options:
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
    destino = str(PASTA_DOWNLOADS)
    opcoes = Options()
    opcoes.add_argument("--headless=new")
    opcoes.add_argument("--window-size=1920,1080")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument("--disable-dev-shm-usage")
    opcoes.add_argument("--ignore-certificate-errors")
    opcoes.add_argument("--disable-blink-features=AutomationControlled")
    opcoes.add_argument(
        "--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    )
    opcoes.add_experimental_option("prefs", {
        "download.default_directory": destino,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
    })
    return opcoes


def _aplicar_pasta_download_chrome(driver) -> None:
    try:
        PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(PASTA_DOWNLOADS)},
        )
        _log(f"Pasta de download do Chrome: {PASTA_DOWNLOADS}", "info")
    except Exception as exc:
        _log(f"Não foi possível forçar pasta de download via CDP: {exc}", "aviso")


def _calcular_datas() -> tuple[str, str]:
    hoje = date.today()
    dias_atras = 3 if hoje.weekday() == 0 else 1
    ontem = hoje - timedelta(days=dias_atras)
    return ontem.strftime("%d/%m/%Y"), hoje.strftime("%d/%m/%Y")


def _preencher_easyui(campo, valor: str) -> None:
    campo.click()
    time.sleep(0.3)
    campo.clear()
    campo.send_keys(valor)


def _clicar(driver, elemento) -> None:
    driver.execute_script("arguments[0].click();", elemento)


def _abrir_combo_por_input(driver, wait: WebDriverWait, input_id: str) -> None:
    campo = wait.until(EC.element_to_be_clickable((By.ID, input_id)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo)
    time.sleep(0.3)
    try:
        seta = campo.find_element(
            By.XPATH,
            "./following-sibling::span[contains(@class,'textbox-addon')]"
            "//a[contains(@class,'combo-arrow')]",
        )
        ActionChains(driver).move_to_element(seta).pause(0.2).click().perform()
    except Exception:
        ActionChains(driver).move_to_element(campo).pause(0.2).click().perform()
    time.sleep(0.5)


def _selecionar_item_combo(driver, wait: WebDriverWait, texto: str) -> None:
    item = wait.until(EC.visibility_of_element_located((
        By.XPATH,
        f"//div[contains(@class,'combobox-item') and normalize-space()='{texto}']",
    )))
    driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'});", item)
    ActionChains(driver).move_to_element(item).pause(0.2).click().perform()
    time.sleep(0.5)


def _selecionar_combo_por_input(
    driver, wait: WebDriverWait, input_id: str, texto: str, combo_name: str = ""
) -> None:
    _abrir_combo_por_input(driver, wait, input_id)
    try:
        _selecionar_item_combo(driver, wait, texto)
    except Exception:
        if not combo_name:
            raise
        _log(f"Tentando selecionar '{texto}' via EasyUI ({combo_name})...", "info")
        driver.execute_script("""
            const comboName = arguments[0];
            const valor = arguments[1];
            if (!window.jQuery || !jQuery('#' + comboName).length) return;
            const combo = jQuery('#' + comboName);
            combo.combobox('showPanel');
            const dados = combo.combobox('getData') || [];
            const item = dados.find(i => i.text === valor || i.value === valor);
            if (item) {
                combo.combobox('select', item.value);
            } else {
                combo.combobox('setText', valor);
            }
        """, combo_name, texto)
        time.sleep(0.5)


def aguardar_download_novo(
    extensoes: tuple,
    arquivos_antes: set | None = None,
    timeout: int = 60,
    nome_contem: str | None = None,
    desde: float | None = None,
) -> str:
    if arquivos_antes is None:
        arquivos_antes = set()
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
    filtro = nome_contem.lower() if nome_contem else None
    _log(f"Monitorando downloads em: {PASTA_DOWNLOADS}", "info")

    for _ in range(timeout):
        time.sleep(1)
        if glob.glob(str(PASTA_DOWNLOADS / "*.crdownload")):
            continue
        if glob.glob(str(PASTA_DOWNLOADS / "*.tmp")):
            continue
        if glob.glob(str(PASTA_DOWNLOADS / "*.part")):
            continue

        candidatos = []
        for ext in extensoes:
            for path in glob.glob(str(PASTA_DOWNLOADS / f"*.{ext}")):
                if os.path.getsize(path) == 0:
                    continue
                if filtro and filtro not in os.path.basename(path).lower():
                    continue
                if path in arquivos_antes:
                    if desde is None or os.path.getmtime(path) < desde:
                        continue
                candidatos.append(path)

        if candidatos:
            arquivo = max(candidatos, key=os.path.getmtime)
            tamanho = os.path.getsize(arquivo)
            time.sleep(0.5)
            if os.path.getsize(arquivo) == tamanho:
                _log(f"Download encontrado: {os.path.basename(arquivo)}", "ok")
                return arquivo

    exts = ", ".join(f".{e}" for e in extensoes)
    filtro_msg = f" contendo '{nome_contem}'" if nome_contem else ""
    existentes = sorted(os.listdir(PASTA_DOWNLOADS))[:20] if PASTA_DOWNLOADS.is_dir() else []
    raise Exception(
        f"Arquivo ({exts}){filtro_msg} não encontrado em '{PASTA_DOWNLOADS}' "
        f"após {timeout}s. Conteúdo recente: {existentes or '(vazio)'}"
    )


def _guardar_download(origem: str, prefixo: str) -> Path:
    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)
    ext = Path(origem).suffix.lower() or ".dat"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = PASTA_DESTINO / f"{prefixo}_{stamp}{ext}"
    shutil.move(origem, destino)
    _log(f"Arquivo guardado: {destino.name}", "ok")
    return destino


def _salvar_diagnostico(driver, etapa: str) -> None:
    """Grava screenshot + HTML para o artifact do Actions."""
    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        url = driver.current_url
        titulo = driver.title
    except Exception:
        url, titulo = "?", "?"
    _log(f"Diagnóstico {etapa}: URL={url} | título={titulo}", "aviso")
    png = PASTA_DESTINO / f"diagnostico_{etapa}_{stamp}.png"
    html = PASTA_DESTINO / f"diagnostico_{etapa}_{stamp}.html"
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


def _focar_login_brudam(driver, wait: WebDriverWait) -> None:
    """Aguarda o formulário de login (#user) no documento ou em iframe."""
    driver.switch_to.default_content()
    try:
        _log(f"URL de login: {driver.current_url}", "info")
    except Exception:
        pass

    if driver.find_elements(By.CSS_SELECTOR, "#formularioAcesso #user, input#user"):
        _log("Login Brudam no documento principal.", "info")
        return

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    _log(
        f"Campo #user não está no documento principal. {len(iframes)} iframe(s).",
        "info",
    )
    for i, iframe in enumerate(iframes):
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(iframe)
        except Exception:
            continue
        if driver.find_elements(By.CSS_SELECTOR, "#formularioAcesso #user, input#user"):
            _log(f"Login Brudam encontrado no iframe {i}.", "ok")
            return

    driver.switch_to.default_content()
    try:
        _log(f"URL atual: {driver.current_url} | título: {driver.title}", "aviso")
    except Exception:
        pass
    wait.until(EC.visibility_of_element_located((By.ID, "user")))


def _preencher_campo_data(driver, campo, valor: str) -> None:
    driver.execute_script("arguments[0].removeAttribute('readonly')", campo)
    driver.execute_script("arguments[0].value = '';", campo)
    campo.click()
    campo.send_keys(valor)
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', {bubbles: true}))", campo
    )


def _exportar_brudam(driver, wait: WebDriverWait) -> list[Path]:
    arquivos: list[Path] = []
    usuario = _credencial("BRUDAM_USUARIO", "USUARIO_BRUDAM")
    senha = _credencial("BRUDAM_SENHA", "SENHA_BRUDAM")

    _log("Etapa 1/2: Brudam DHL", "etapa")
    _log(f"Abrindo {URL_LOGIN}...", "info")
    driver.get(URL_LOGIN)
    wait.until(EC.presence_of_element_located((By.ID, "formularioAcesso")))
    _focar_login_brudam(driver, wait)

    campo_usuario = wait.until(EC.visibility_of_element_located((By.ID, "user")))
    campo_usuario.clear()
    campo_usuario.send_keys(usuario)
    time.sleep(0.5)

    campo_senha = wait.until(EC.presence_of_element_located((By.ID, "password")))
    campo_senha.clear()
    campo_senha.send_keys(senha)
    time.sleep(0.5)

    botao_login = wait.until(EC.element_to_be_clickable((By.ID, "acessar")))
    driver.execute_script("arguments[0].click();", botao_login)
    driver.switch_to.default_content()

    try:
        botao_fechar_modal = wait.until(EC.element_to_be_clickable((
            By.CSS_SELECTOR, "button.btn_secondary_964a3b119be.close_modal_964a3b119be"
        )))
        _log("Modal de aviso detectado — fechando...", "info")
        driver.execute_script("arguments[0].click();", botao_fechar_modal)
        time.sleep(1)
    except Exception:
        _log("Modal não apareceu, continuando...", "aviso")

    data_inicio, data_fim = _calcular_datas()
    _log(f"Filtro de datas: {data_inicio} até {data_fim}", "etapa")

    _log("Navegando para o relatório de embarque...", "etapa")
    driver.get(URL_EMBARQUE)
    time.sleep(3)

    campo_data_inicio = wait.until(EC.presence_of_element_located((By.NAME, "data_1")))
    _preencher_campo_data(driver, campo_data_inicio, data_inicio)
    time.sleep(0.5)
    campo_data_fim = wait.until(EC.presence_of_element_located((By.NAME, "data_2")))
    _preencher_campo_data(driver, campo_data_fim, data_fim)
    time.sleep(0.5)

    botao_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "PESQUISAR")))
    driver.execute_script("arguments[0].click();", botao_pesquisar)
    wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "table#contents tbody tr")) > 0)

    botao_excel = wait.until(EC.element_to_be_clickable((By.ID, "excel")))
    driver.execute_script("arguments[0].click();", botao_excel)
    botao_csv = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(., 'Formato CSV')]"
    )))
    arquivos_antes_embarque = set(glob.glob(str(PASTA_DOWNLOADS / "*")))
    driver.execute_script("arguments[0].click();", botao_csv)
    arquivo_csv = aguardar_download_novo(("csv",), arquivos_antes_embarque)
    arquivos.append(_guardar_download(arquivo_csv, "embarque"))

    _log("Iniciando relatório detalhado (Lista de Notas / relatório 668)...", "etapa")
    driver.get(URL_LISTA_NOTAS)
    time.sleep(3)

    campo_data_inicio2 = wait.until(EC.presence_of_element_located((By.ID, "data_1")))
    _preencher_campo_data(driver, campo_data_inicio2, data_inicio)
    time.sleep(0.3)
    campo_data_fim2 = wait.until(EC.presence_of_element_located((By.ID, "data_2")))
    _preencher_campo_data(driver, campo_data_fim2, data_fim)
    time.sleep(0.3)

    botao_pesquisar2 = wait.until(EC.element_to_be_clickable((By.ID, "PESQUISAR")))
    driver.execute_script("arguments[0].click();", botao_pesquisar2)
    wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "table#gridMinuta tbody tr")) > 0)

    botao_excel_pers = wait.until(EC.element_to_be_clickable((By.ID, "excel_pers")))
    driver.execute_script("arguments[0].click();", botao_excel_pers)
    botao_meus_rel = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[contains(., 'Meus relatórios')]"
    )))
    driver.execute_script("arguments[0].click();", botao_meus_rel)
    time.sleep(1)

    radio_668 = wait.until(EC.presence_of_element_located((
        By.XPATH,
        "//input[@type='radio' and contains(@value, "
        "'eyJjb2x1bmFzIjoiMSwyLDMsNiwxNywyNiw0NiIsImNhYmVjYWxobyI6IkNURTtORjtDTElFTlRFO05VTUVSTyBBV0I7T0JTIERPIEFXQjtRVUFOVElEQURFIERFIFZPTFVNRVM7VUxUSU1PIFNUQVRVUyJ9')]",
    )))
    driver.execute_script("arguments[0].click();", radio_668)
    time.sleep(1)

    xpath_gerar = (
        "(//div[contains(@class,'ui-dialog') and @role='dialog'"
        " and not(contains(@style,'display: none'))]"
        "//div[contains(@class,'ui-dialog-buttonset')]"
        "//button[normalize-space()='Gerar'])[last()]"
    )
    botao_gerar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_gerar)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_gerar)
    time.sleep(0.3)

    arquivos_antes_detalhado = set(glob.glob(str(PASTA_DOWNLOADS / "*")))
    momento_download = time.time()
    try:
        ActionChains(driver).move_to_element(botao_gerar).pause(0.2).click().perform()
    except Exception:
        driver.execute_script("arguments[0].click();", botao_gerar)

    arquivo_baixado = aguardar_download_novo(
        ("csv",),
        arquivos_antes_detalhado,
        nome_contem="relatorio_notas",
        desde=momento_download,
        timeout=120,
    )
    arquivos.append(_guardar_download(arquivo_baixado, "detalhado"))
    _log("Brudam concluído (arquivos brutos, sem consolidação).", "ok")
    return arquivos


def _menu_relatorios_visivel(driver) -> bool:
    menus = driver.find_elements(By.ID, "mmRelatorios")
    return bool(menus) and menus[0].is_displayed()


def _abrir_menu_relatorios(driver, wait: WebDriverWait) -> None:
    menu_btn = wait.until(EC.visibility_of_element_located((By.ID, "mbRelatorios")))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu_btn)
    time.sleep(0.5)
    ActionChains(driver).move_to_element(menu_btn).pause(1).perform()
    if not _menu_relatorios_visivel(driver):
        texto_relatorios = menu_btn.find_element(By.CSS_SELECTOR, "span.l-btn-text")
        ActionChains(driver).move_to_element(texto_relatorios).pause(0.5).click().perform()
    if not _menu_relatorios_visivel(driver):
        driver.execute_script("""
            const btn = document.getElementById('mbRelatorios');
            if (window.jQuery && btn) {
                jQuery('#mbRelatorios').menubutton('showMenu');
            } else if (btn) {
                btn.click();
            }
        """)
    wait.until(lambda d: _menu_relatorios_visivel(d))


def _exportar_tsms(driver, wait: WebDriverWait) -> list[Path]:
    usuario = _credencial("TSMS_USUARIO", "USUARIO_TSMS")
    senha = _credencial("TSMS_SENHA", "SENHA_TSMS")

    _log("Etapa 2/2: TSMS", "etapa")
    driver.get(URL_TSMS)
    time.sleep(3)

    campo_usuario = wait.until(EC.element_to_be_clickable((By.ID, "_easyui_textbox_input1")))
    _preencher_easyui(campo_usuario, usuario)
    campo_senha = wait.until(EC.element_to_be_clickable((By.ID, "_easyui_textbox_input2")))
    _preencher_easyui(campo_senha, senha)
    botao_entrar = wait.until(EC.element_to_be_clickable((By.ID, "btnEntrar")))
    _clicar(driver, botao_entrar)
    time.sleep(2)

    _log("Atualizando página após login no TSMS (bypass do popup)...", "info")
    driver.refresh()
    time.sleep(3)
    driver.switch_to.default_content()

    wait.until(EC.presence_of_element_located((By.ID, "mbRelatorios")))
    time.sleep(2)
    _abrir_menu_relatorios(driver, wait)
    item_viagens = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        "//div[@id='mmRelatorios']//div[contains(@class,'menu-text') and normalize-space()='Viagens']",
    )))
    ActionChains(driver).move_to_element(item_viagens).pause(0.3).click().perform()
    time.sleep(2)

    data_inicio, data_fim = _calcular_datas()
    _log(f"Preenchendo datas TSMS: {data_inicio} até {data_fim}", "etapa")
    campo_data_inicio = wait.until(EC.element_to_be_clickable((By.ID, "_easyui_textbox_input15")))
    _preencher_easyui(campo_data_inicio, data_inicio)
    campo_data_fim = wait.until(EC.element_to_be_clickable((By.ID, "_easyui_textbox_input17")))
    _preencher_easyui(campo_data_fim, data_fim)
    _selecionar_combo_por_input(driver, wait, "_easyui_textbox_input12", "APPLE", "relOperacao")
    _selecionar_combo_por_input(
        driver, wait, "_easyui_textbox_input14", "Formato Planilha(xls)", "relSaida"
    )

    arquivos_antes = set(glob.glob(str(PASTA_DOWNLOADS / "*")))
    momento_download = time.time()
    botao_gerar = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        "//a[contains(@class,'easyui-linkbutton') and .//span[contains(@class,'l-btn-text') "
        "and normalize-space()='Gerar']]",
    )))
    _clicar(driver, botao_gerar)

    arquivo_baixado = aguardar_download_novo(
        ("xls", "xlsx"),
        arquivos_antes,
        desde=momento_download,
        timeout=120,
    )
    destino = _guardar_download(arquivo_baixado, "tsms")
    _log("TSMS concluído (arquivo bruto, sem consolidação).", "ok")
    return [destino]


def _enviar_email_smtp(anexos: list[Path], assunto: str, corpo: str) -> None:
    remetente = GMAIL_REMETENTE or _credencial("GMAIL_USER")
    senha = _credencial("GMAIL_APP_PASSWORD", "GMAIL_PASSWORD")
    senha = senha.replace(" ", "")
    servidor = os.environ.get("SMTP_SERVER", "smtp.gmail.com").strip()
    porta = int(os.environ.get("SMTP_PORT", "587"))

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = remetente
    msg["To"] = EMAIL_DESTINO
    msg.set_content(corpo)
    for path in anexos:
        dados = path.read_bytes()
        maintype, subtype = ("application", "octet-stream")
        if path.suffix.lower() == ".csv":
            maintype, subtype = "text", "csv"
        msg.add_attachment(
            dados,
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
            "Gmail recusou o login SMTP. Use uma senha de app (não a senha "
            "normal da conta): Google Account → Segurança → Verificação em "
            "duas etapas → Senhas de app. Grave essa senha no secret "
            "GMAIL_APP_PASSWORD (16 caracteres, pode colar com ou sem espaços)."
        ) from e
    _log(f"E-mail enviado para {EMAIL_DESTINO} (de {remetente})", "ok")

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = remetente
    msg["To"] = EMAIL_DESTINO
    msg.set_content(corpo)
    for path in anexos:
        dados = path.read_bytes()
        maintype, subtype = ("application", "octet-stream")
        if path.suffix.lower() == ".csv":
            maintype, subtype = "text", "csv"
        msg.add_attachment(
            dados,
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
            "O Microsoft 365 da DHL bloqueou SMTP com senha (535 basic auth "
            "disabled). Sem Graph, o e-mail não sai da nuvem até o TI habilitar "
            "SMTP AUTH nessa caixa. Baixe os arquivos no artifact "
            "'relatorios-rpa-apple' desta execução do Actions."
        ) from e
    _log(f"E-mail Outlook enviado para {EMAIL_DESTINO} (de {remetente})", "ok")


def enviar_email(anexos: list[Path], erros: list[str]) -> None:
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    linhas = [
        f"RPA Apple (GitHub Actions) — {agora}",
        "",
        f"Arquivos anexados ({len(anexos)}):",
    ]
    if anexos:
        linhas.extend(f"  - {p.name}" for p in anexos)
    else:
        linhas.append("  (nenhum arquivo baixado)")
    if erros:
        linhas.append("")
        linhas.append("Erros:")
        linhas.extend(f"  - {e}" for e in erros)
    corpo = "\n".join(linhas)
    assunto = f"RPA Apple — relatórios {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    if erros:
        assunto = "[FALHA] " + assunto

    _enviar_email_smtp(anexos, assunto, corpo)


def main() -> int:
    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)
    PASTA_DOWNLOADS.mkdir(parents=True, exist_ok=True)
    _log("Iniciando RPA Apple (GitHub Actions).", "etapa")

    driver = None
    arquivos: list[Path] = []
    erros: list[str] = []
    try:
        driver = webdriver.Chrome(options=_criar_opcoes_chrome())
        _aplicar_pasta_download_chrome(driver)
        wait = WebDriverWait(driver, 45)
        _log("Navegador aberto.", "ok")

        try:
            arquivos.extend(_exportar_brudam(driver, wait))
        except Exception as e:
            if _navegador_fechado(e):
                raise
            _salvar_diagnostico(driver, "brudam")
            _log(f"Falha no Brudam: {e}", "erro")
            erros.append(f"Brudam: {e}")

        try:
            arquivos.extend(_exportar_tsms(driver, wait))
        except Exception as e:
            if _navegador_fechado(e):
                raise
            _salvar_diagnostico(driver, "tsms")
            _log(f"Falha no TSMS: {e}", "erro")
            erros.append(f"TSMS: {e}")
    except Exception as e:
        _log(f"Erro: {e}", "erro")
        erros.append(str(e))
    finally:
        _encerrar_driver(driver)

    try:
        enviar_email(arquivos, erros)
    except Exception as e:
        _log(f"Falha ao enviar e-mail: {e}", "erro")
        erros.append(f"E-mail: {e}")

    if erros:
        _log("Execução concluída com erros.", "erro")
        return 1
    _log("Execução concluída com sucesso.", "ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
