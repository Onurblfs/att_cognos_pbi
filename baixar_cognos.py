# -*- coding: utf-8 -*-
"""
Automacao de download das bases do IBM Planning Analytics (Cognos)
para atualizacao do Power BI "Orcamento e Forecast Gerencial".

Fluxo (replica o processo manual do documento de atualizacao):
  1. Abre o site do Planning Analytics.
  2. Aguarda o login (manual ou automatico via SSO).
  3. Para cada exportacao do config.json:
       - pesquisa o nome da view na pasta Compartilhado;
       - abre a view de exportacao;
       - exporta para Excel;
       - move o arquivo baixado para a pasta de rede de destino
         (fazendo backup local do arquivo anterior).

Uso:
  python baixar_cognos.py                     -> executa todas as exportacoes
  python baixar_cognos.py --somente IRAT.950  -> executa so as que contem o texto no nome
  python baixar_cognos.py --sem-mover         -> baixa mas nao copia para a rede (teste)
"""

import argparse
import json
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

BASE_DIR = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Seletores da interface do Planning Analytics Workspace (Claro).
# O campo "Pesquisar" so aparece DEPOIS de abrir a pasta Compartilhado
# (sidebar esquerda). Cada entrada e uma lista de alternativas em ordem.
# ---------------------------------------------------------------------------
SELETORES = {
    # Debug Claro: botao id=com.ibm.bi.glass.common.navmenu, aria-label="Início".
    "menu_hamburguer": [
        (By.CSS_SELECTOR, "#com\\.ibm\\.bi\\.glass\\.common\\.navmenu"),
        (By.CSS_SELECTOR, "button[data-id='com.ibm.bi.glass.common.navmenu']"),
        (By.CSS_SELECTOR, "[walkme-data-id='com.ibm.bi.glass.common.navmenu'] button"),
        (By.CSS_SELECTOR, "button.ba-carbon-nav-menu"),
        (By.CSS_SELECTOR, "button[data-tid='buc-OverflowMenu']"),
        (By.XPATH, "//button[@id='com.ibm.bi.glass.common.navmenu']"),
    ],
    "link_compartilhado": [
        (By.XPATH, "//*[self::a or self::button or self::span or self::div][normalize-space(.)='Compartilhado']"),
        (By.XPATH, "//*[contains(@aria-label,'Compartilhado')]"),
        (By.XPATH, "//*[contains(@class,'create-menu-link')][contains(.,'Compartilhado')]"),
        (By.XPATH, "//*[contains(.,'Conteúdo compartilhado') or contains(.,'Conteudo compartilhado')]"),
        (By.XPATH, "//*[contains(.,'Pasta compartilhada') or contains(.,'Arquivos compartilhados')]"),
        (By.XPATH, "//*[self::a or self::button or self::span][normalize-space(.)='Shared']"),
        (By.XPATH, "//*[contains(text(),'Compartilhado') or contains(text(),'Shared')]"),
    ],
    "card_aplicativos": [
        (By.XPATH, "//*[contains(.,'Aplicativos e planos')]"),
        (By.XPATH, "//*[contains(.,'Apps and plans')]"),
    ],
    "aba_favoritos": [
        (By.CSS_SELECTOR, "#tab-favorites"),
        (By.XPATH, "//*[@role='tab' and (normalize-space(.)='Favoritos' or @title='Favoritos')]"),
        (By.XPATH, "//button[normalize-space(.)='Favoritos' or @id='tab-favorites']"),
        (By.XPATH, "//*[@title='Favoritos']"),
    ],
    "aba_recentes": [
        (By.CSS_SELECTOR, "#tab-recents, #tab-recent"),
        (By.XPATH, "//*[@role='tab' and (normalize-space(.)='Recentes' or @title='Recentes')]"),
        (By.XPATH, "//button[normalize-space(.)='Recentes']"),
        (By.XPATH, "//*[@title='Recentes']"),
    ],
    "campo_pesquisa": [
        # Campo da sidebar Compartilhado (screenshot: placeholder "Pesquisar")
        (By.XPATH, "//input[contains(@placeholder,'Pesquisar') or contains(@placeholder,'esquis')]"),
        (By.CSS_SELECTOR, "input[placeholder*='Pesquisar']"),
        (By.CSS_SELECTOR, "input[placeholder*='esquis']"),
        (By.CSS_SELECTOR, "input[placeholder*='Search']"),
        (By.CSS_SELECTOR, "input[type='search']"),
        (By.CSS_SELECTOR, "input[role='searchbox']"),
        (By.XPATH, "//aside//input | //nav//input | //*[contains(@class,'search')]//input"),
    ],
    "home_carregada": [
        (By.XPATH, "//*[contains(.,'Início rápido') or contains(.,'Inicio rapido') or contains(.,'Quick start')]"),
        (By.XPATH, "//*[contains(.,'Meus aplicativos') or contains(.,'My applications')]"),
        (By.XPATH, "//*[contains(.,'IBM Planning Analytics')]"),
        (By.CSS_SELECTOR, "#com\\.ibm\\.bi\\.glass\\.common\\.navmenu"),
    ],
    # Botao da toolbar on-demand do cubo (tooltip Claro: "Exportar para planilha").
    # So aparece apos selecionar o widget e com Editar DESLIGADO.
    "botao_exportar_planilha": [
        (By.XPATH, "//*[@aria-label='Exportar para planilha' or @title='Exportar para planilha']"),
        (By.XPATH, "//button[contains(@aria-label,'Exportar para planilha') or contains(@title,'Exportar para planilha')]"),
        (By.XPATH, "//*[contains(@aria-label,'Export to spreadsheet') or contains(@title,'Export to spreadsheet')]"),
        (By.XPATH, "//*[contains(@aria-label,'planilha') or contains(@title,'planilha')]"),
        (By.XPATH, "//button[.//span[contains(@class,'assistive') and contains(.,'planilha')]]"),
        (By.XPATH, "//button[.//span[contains(@class,'assistive') and contains(.,'Exportar')]]"),
        (By.XPATH, "//*[contains(@class,'OnDemandToolbar') or contains(@class,'toolbarDock')]//button"),
        (By.XPATH, "//button[.//use[contains(@*,'document_export') or contains(@*,'common-download') or contains(@*,'metricsExport')]]"),
    ],
    "toggle_editar": [
        (By.CSS_SELECTOR, "input[id*='editToggleButton']"),
        (By.XPATH, "//input[contains(@id,'editToggleButton')]"),
    ],
    "widget_cubo": [
        (By.XPATH, "//*[contains(@aria-label,'Visualização do cubo') or contains(@aria-label,'Visualizacao do cubo')]"),
        (By.XPATH, "//*[contains(@aria-label,'Cube view') or contains(@aria-label,'cube view')]"),
        (By.CSS_SELECTOR, "[aria-label*='IRAT.'], [aria-label*='FIS.'], [aria-label*='REV.'], [aria-label*='CTS.']"),
    ],
    "menu_exportar": [
        (By.XPATH, "//*[self::button or self::span or self::a or self::li or self::div][normalize-space(.)='Exportar' or normalize-space(.)='Export']"),
        (By.XPATH, "//*[contains(@aria-label,'Exportar') or contains(@aria-label,'Export')]"),
        (By.XPATH, "//*[contains(@title,'Exportar') or contains(@title,'Export')]"),
        (By.CSS_SELECTOR, "button[aria-label*='xport']"),
    ],
    "opcao_excel": [
        (By.XPATH, "//*[self::button or self::span or self::a or self::li or self::div][contains(.,'Excel') or contains(.,'xlsx')]"),
        (By.XPATH, "//*[contains(@aria-label,'Excel') or contains(@title,'Excel')]"),
        (By.XPATH, "//*[contains(text(), 'Excel')]"),
        (By.XPATH, "//*[contains(text(), 'xlsx')]"),
    ],
    "confirmar_exportacao": [
        (By.XPATH, "//div[contains(@id,'ExportExcelDialog') or contains(@class,'ExportView')]//button[contains(.,'Exportar') or contains(.,'Export') or contains(.,'OK')]"),
        (By.XPATH, "//button[contains(., 'Exportar')]"),
        (By.XPATH, "//button[contains(., 'OK')]"),
        (By.XPATH, "//button[contains(., 'Export')]"),
    ],
    "grade_cubo": [
        (By.CSS_SELECTOR, "[class*='TM1MDV']"),
        (By.CSS_SELECTOR, "[class*='cubeViewer'], [class*='CubeViewer']"),
        (By.CSS_SELECTOR, "[class*='exploration'], [class*='Exploration']"),
        (By.XPATH, "//*[contains(@class,'grid') or contains(@class,'tigre')]"),
        (By.CSS_SELECTOR, "canvas"),
    ],
}

EXTENSOES_VALIDAS = {".xlsx", ".xls", ".csv"}


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def carregar_config(caminho: Path) -> dict:
    with open(caminho, encoding="utf-8") as f:
        return json.load(f)


def criar_driver(pasta_downloads: Path) -> webdriver.Edge:
    opcoes = webdriver.EdgeOptions()
    opcoes.set_capability("acceptInsecureCerts", True)  # certificado corporativo autoassinado
    opcoes.add_experimental_option("prefs", {
        "download.default_directory": str(pasta_downloads),
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
    })
    opcoes.add_argument("--start-maximized")
    return webdriver.Edge(options=opcoes)


def achar_elemento(driver, chave_seletor: str, timeout: int = 20, clicavel: bool = True):
    """Tenta as alternativas de seletor em ordem ate encontrar um elemento."""
    condicao = EC.element_to_be_clickable if clicavel else EC.presence_of_element_located
    fim = time.time() + timeout
    while time.time() < fim:
        for by, seletor in SELETORES[chave_seletor]:
            try:
                el = WebDriverWait(driver, 2).until(condicao((by, seletor)))
                return el
            except TimeoutException:
                continue
    raise TimeoutException(
        f"Nao encontrei o elemento '{chave_seletor}'. "
        "Os seletores em SELETORES provavelmente precisam de ajuste para esta versao do PA."
    )


def clicar(driver, elemento) -> None:
    """Clica com scroll + fallback via JavaScript (PA às vezes bloqueia click nativo)."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
        time.sleep(0.3)
        elemento.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elemento)


def listar_inputs_visiveis(driver) -> list:
    """Lista inputs na pagina (inclui shadow DOM). Util para debug."""
    script = """
    const out = [];
    function walk(root, path) {
      root.querySelectorAll('input, textarea').forEach((el, i) => {
        out.push({
          path: path + '/' + el.tagName.toLowerCase() + '[' + i + ']',
          type: el.type || '',
          placeholder: el.placeholder || '',
          aria: el.getAttribute('aria-label') || '',
          name: el.name || '',
          id: el.id || '',
          visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length)
        });
      });
      root.querySelectorAll('*').forEach((el) => {
        if (el.shadowRoot) walk(el.shadowRoot, path + '/' + (el.tagName || 'shadow'));
      });
    }
    walk(document, 'doc');
    return out;
    """
    try:
        return driver.execute_script(script) or []
    except Exception:
        return []


def achar_input_por_js(driver, textos=("pesquisar", "search", "esquis")):
    """Encontra input pelo placeholder/aria, inclusive dentro de shadow DOM."""
    script = """
    const textos = arguments[0].map(t => t.toLowerCase());
    function match(el) {
      const p = (el.placeholder || '').toLowerCase();
      const a = (el.getAttribute('aria-label') || '').toLowerCase();
      const t = (el.type || '').toLowerCase();
      if (t === 'hidden') return false;
      return textos.some(x => p.includes(x) || a.includes(x)) || t === 'search';
    }
    function walk(root) {
      const inputs = root.querySelectorAll('input, textarea');
      for (const el of inputs) {
        if (match(el)) return el;
      }
      for (const el of root.querySelectorAll('*')) {
        if (el.shadowRoot) {
          const found = walk(el.shadowRoot);
          if (found) return found;
        }
      }
      return null;
    }
    return walk(document);
    """
    return driver.execute_script(script, list(textos))


def com_iframes(driver, fn):
    """
    Executa fn(driver) no documento principal e em cada iframe.
    Retorna o primeiro resultado nao-nulo. Volta sempre ao default_content.
    """
    driver.switch_to.default_content()
    try:
        resultado = fn(driver)
        if resultado is not None:
            return resultado
    except Exception:
        pass

    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for idx, frame in enumerate(iframes):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)
            resultado = fn(driver)
            if resultado is not None:
                log(f"Elemento encontrado no iframe[{idx}].")
                return resultado
        except Exception:
            continue

    driver.switch_to.default_content()
    return None


def achar_campo_pesquisa(driver, timeout: int = 20):
    """Localiza o campo Pesquisar no PA (documento, iframes e shadow DOM)."""
    fim = time.time() + timeout
    while time.time() < fim:
        def tentar(_drv):
            try:
                return achar_elemento(_drv, "campo_pesquisa", timeout=2, clicavel=False)
            except TimeoutException:
                return achar_input_por_js(_drv)

        el = com_iframes(driver, tentar)
        if el is not None:
            return el
        time.sleep(1)
    raise TimeoutException(
        "Nao encontrei o elemento 'campo_pesquisa'. "
        "Os seletores em SELETORES provavelmente precisam de ajuste para esta versao do PA."
    )


def salvar_debug(driver, pasta: Path, rotulo: str, erro: str | None = None) -> Path:
    """Salva screenshot + lista de inputs + HTML para diagnostico."""
    pasta.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Evita caracteres invalidos no Windows (ex.: parenteses no nome da exportacao).
    seguro = "".join(c if c.isalnum() or c in "-_." else "_" for c in rotulo)[:60]
    base = pasta / f"debug_{seguro}_{ts}"
    try:
        driver.switch_to.default_content()
        driver.save_screenshot(str(base.with_suffix(".png")))
    except Exception as e:
        log(f"Falha ao salvar screenshot: {e}")
    try:
        inputs = listar_inputs_visiveis(driver)
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        relatorio = {
            "url": driver.current_url,
            "title": driver.title,
            "iframes": len(iframes),
            "erro": erro,
            "inputs": inputs,
        }
        base.with_suffix(".json").write_text(
            json.dumps(relatorio, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        log(f"Debug salvo em: {base.with_suffix('.json')} ({len(inputs)} inputs, {len(iframes)} iframes)")
    except Exception as e:
        log(f"Falha ao salvar debug JSON: {e}")
    try:
        base.with_suffix(".html").write_text(driver.page_source, encoding="utf-8")
    except Exception:
        pass
    return base


def aguardar_login(driver, timeout: int) -> None:
    """Espera ate a home do PA carregar (apos login manual/SSO)."""
    log("Aguardando login... Se aparecer a tela de login, entre com seu usuario.")
    fim = time.time() + timeout
    while time.time() < fim:
        try:
            for by, seletor in SELETORES["home_carregada"]:
                if driver.find_elements(by, seletor):
                    log("Login concluido, pagina inicial carregada.")
                    time.sleep(3)
                    return
        except Exception:
            pass
        time.sleep(2)
    raise TimeoutException("Tempo esgotado aguardando o login no Planning Analytics.")


def abrir_por_tile(driver, nome_busca: str) -> bool:
    """
    Tenta abrir a view pelo tile ja presente na home (Favoritos/Recentes).
    No HTML de debug da Claro esses tiles existem com:
      <div title="Receita DRE PowerBI V2 (irat950)" class="pa-tile-header ...">
    """
    literais = xpath_literal(nome_busca)
    xpaths = [
        f"//div[contains(@class,'pa-tile-header') and @title={literais}]",
        f"//*[@title={literais} and contains(@class,'pa-tile')]",
        f"//a[contains(@class,'pa-tile')][.//div[@title={literais}]]",
        f"//div[contains(@class,'click-area')][.//*[@title={literais}]]",
    ]

    # Tenta nas abas onde os tiles costumam aparecer.
    for aba in ("aba_favoritos", "aba_recentes"):
        try:
            el_aba = achar_elemento(driver, aba, timeout=3)
            clicar(driver, el_aba)
            time.sleep(2)
        except TimeoutException:
            pass

        for xp in xpaths:
            els = driver.find_elements(By.XPATH, xp)
            for el in els:
                try:
                    if not el.is_displayed():
                        continue
                    log(f"Tile encontrado na home: {nome_busca}")
                    clicar(driver, el)
                    time.sleep(10)
                    return True
                except Exception:
                    continue
    return False


def abrir_compartilhado(driver, pasta_debug: Path | None = None) -> None:
    """Abre o menu de navegacao e entra em Compartilhado (campo Pesquisar)."""
    try:
        achar_campo_pesquisa(driver, timeout=3)
        log("Sidebar Compartilhado ja aberta.")
        return
    except TimeoutException:
        pass

    log("Abrindo menu de navegacao (navmenu)...")
    try:
        menu = achar_elemento(driver, "menu_hamburguer", timeout=15)
        clicar(driver, menu)
        time.sleep(2)
    except TimeoutException:
        if pasta_debug is not None:
            salvar_debug(driver, pasta_debug, "sem_navmenu")
        raise TimeoutException(
            "Nao encontrei o botao do menu (#com.ibm.bi.glass.common.navmenu)."
        )

    try:
        link = achar_elemento(driver, "link_compartilhado", timeout=20)
        clicar(driver, link)
        time.sleep(4)
    except TimeoutException:
        if pasta_debug is not None:
            salvar_debug(driver, pasta_debug, "sem_compartilhado")
        raise

    try:
        achar_campo_pesquisa(driver, timeout=20)
        log("Compartilhado aberto.")
    except TimeoutException:
        if pasta_debug is not None:
            salvar_debug(driver, pasta_debug, "sem_campo_pesquisa")
        raise


def abrir_por_dashboard_id(driver, base_url: str, dashboard_id: str) -> bool:
    """Abre a view direto pela URL do dashboard (mais estavel que tile/pesquisa)."""
    if not dashboard_id:
        return False
    # Ex.: https://host:9443/?perspective=pa-home -> .../?perspective=dashboard&id=...
    root = base_url.split("?")[0].rstrip("/")
    url = f"{root}/?perspective=dashboard&id={dashboard_id}"
    log(f"Abrindo dashboard direto: {url}")
    driver.get(url)
    time.sleep(12)
    if "perspective=dashboard" in driver.current_url:
        return True
    return False


def pesquisar_e_abrir(
    driver,
    nome_busca: str,
    pasta_debug: Path | None = None,
    dashboard_id: str | None = None,
    base_url: str | None = None,
) -> None:
    """Abre a view: URL direta > tile na home > pesquisa no Compartilhado."""
    log(f"Abrindo view: {nome_busca}")

    if dashboard_id and base_url and abrir_por_dashboard_id(driver, base_url, dashboard_id):
        log("View aberta via dashboard_id.")
        return

    if abrir_por_tile(driver, nome_busca):
        log("View aberta via tile da home.")
        return

    log("Tile nao encontrado na home; indo para Compartilhado...")
    # Atalho: card "Aplicativos e planos" às vezes leva ao browser de conteudo.
    try:
        card = achar_elemento(driver, "card_aplicativos", timeout=5)
        clicar(driver, card)
        time.sleep(4)
    except TimeoutException:
        pass

    abrir_compartilhado(driver, pasta_debug=pasta_debug)

    log(f"Pesquisando: {nome_busca}")
    campo = achar_campo_pesquisa(driver, timeout=20)
    clicar(driver, campo)
    try:
        campo.send_keys(Keys.CONTROL, "a")
        campo.send_keys(Keys.BACKSPACE)
    except Exception:
        driver.execute_script(
            "arguments[0].value='';"
            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));",
            campo,
        )
    time.sleep(0.3)
    campo.send_keys(nome_busca)
    time.sleep(1)
    campo.send_keys(Keys.ENTER)
    time.sleep(4)

    resultado = achar_resultado(driver, nome_busca, timeout=30)
    clicar(driver, resultado)
    log("View aberta via pesquisa no Compartilhado. Aguardando carregar...")
    time.sleep(10)


def achar_resultado(driver, nome_busca: str, timeout: int = 30):
    """
    Localiza o resultado da pesquisa.
    A UI do PA trunca o texto (ex.: 'Receita DRE P... V2 (irat950)'),
    entao tenta match exato, title/aria-label e contains parcial.
    """
    literais = xpath_literal(nome_busca)
    # Trechos distintivos do nome (evita depender do texto truncado).
    # NAO usar sufixos genericos como "(Power BI)" — batem em varias views.
    trechos = [nome_busca]
    if "(" in nome_busca and nome_busca.endswith(")"):
        sufixo = nome_busca[nome_busca.rfind("(") :]
        if sufixo.lower() not in {"(power bi)", "(tableau)", "(cognos)"}:
            trechos.append(sufixo)  # ex.: (irat950)
    # Codigo do cubo no inicio do nome, se houver (CTS.100, REV.420, etc.).
    primeiro = nome_busca.split()[0]
    if "." in primeiro and len(primeiro) >= 5:
        trechos.append(primeiro)

    candidatos_xpath = [
        f"//*[@title={literais} or @aria-label={literais}]",
        f"//*[normalize-space(.)={literais}]",
    ]
    for trecho in trechos:
        lit = xpath_literal(trecho)
        candidatos_xpath.append(f"//*[contains(normalize-space(.), {lit})]")
        candidatos_xpath.append(f"//*[@title[contains(., {lit})] or @aria-label[contains(., {lit})]]")

    fim = time.time() + timeout
    while time.time() < fim:
        for xp in candidatos_xpath:
            els = driver.find_elements(By.XPATH, xp)
            for el in els:
                try:
                    texto = (el.text or "").strip()
                    title = (el.get_attribute("title") or el.get_attribute("aria-label") or "").strip()
                    # Ignora o proprio campo de busca / labels genericos.
                    if el.tag_name.lower() in {"input", "textarea", "html", "body"}:
                        continue
                    if nome_busca.lower() in texto.lower() or nome_busca.lower() in title.lower():
                        if el.is_displayed():
                            log(f"Resultado encontrado: {texto or title}")
                            return el
                    # Match por trecho curto distintivo (ex.: (irat950)) quando truncado.
                    for trecho in trechos[1:]:
                        if trecho.lower() in texto.lower() or trecho.lower() in title.lower():
                            if el.is_displayed() and len(texto) > 3:
                                log(f"Resultado encontrado (parcial): {texto or title}")
                                return el
                except Exception:
                    continue
        time.sleep(1)
    raise TimeoutException(f"Nao encontrei o resultado da pesquisa para: {nome_busca}")


def xpath_literal(texto: str) -> str:
    """Gera literal XPath seguro para textos com aspas."""
    if "'" not in texto:
        return f"'{texto}'"
    if '"' not in texto:
        return f'"{texto}"'
    partes = texto.split("'")
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in partes) + ")"


def _clicar_item_menu_por_texto(driver, textos, timeout: int = 10) -> bool:
    """Clica no primeiro item de menu/contexto cujo texto contenha um dos textos."""
    fim = time.time() + timeout
    while time.time() < fim:
        for texto in textos:
            xp = (
                "//*[self::button or self::span or self::a or self::li or self::div]"
                f"[contains(translate(normalize-space(.),"
                f"'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), '{texto.lower()}')]"
            )
            for el in driver.find_elements(By.XPATH, xp):
                try:
                    if not el.is_displayed():
                        continue
                    t = (el.text or el.get_attribute("aria-label") or "").strip()
                    if not t:
                        continue
                    clicar(driver, el)
                    log(f"Clique em menu: {t}")
                    return True
                except Exception:
                    continue
        time.sleep(0.5)
    return False


def desligar_modo_editar(driver) -> None:
    """O botao Exportar para planilha so aparece com Editar desligado."""
    try:
        toggle = achar_elemento(driver, "toggle_editar", timeout=5, clicavel=False)
    except TimeoutException:
        return
    try:
        if toggle.is_selected():
            log("Desligando modo Editar...")
            # Clica no label/toggle via JS (input checkbox as vezes nao e clicavel).
            driver.execute_script(
                "arguments[0].click();"
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                toggle,
            )
            time.sleep(2)
    except Exception as e:
        log(f"Nao foi possivel desligar Editar: {e}")


def focar_widget_cubo(driver) -> None:
    """Seleciona o widget do cubo para exibir a OnDemandToolbar."""
    try:
        widget = achar_elemento(driver, "widget_cubo", timeout=10, clicavel=False)
        clicar(driver, widget)
        time.sleep(1.5)
        return
    except TimeoutException:
        pass
    try:
        grade = achar_elemento(driver, "grade_cubo", timeout=10, clicavel=False)
        clicar(driver, grade)
        time.sleep(1.5)
    except TimeoutException:
        log("Widget/grade do cubo nao localizados.")


def achar_botao_exportar_planilha_js(driver):
    """Procura o botao por aria/title/texto, inclusive shadow DOM."""
    script = """
    const needles = ['exportar para planilha','export to spreadsheet','planilha','spreadsheet'];
    function texts(el) {
      return [
        el.getAttribute('aria-label') || '',
        el.getAttribute('title') || '',
        el.innerText || '',
        ...(Array.from(el.querySelectorAll('span')).map(s => s.textContent || ''))
      ].join(' | ').toLowerCase();
    }
    function walk(root) {
      const nodes = root.querySelectorAll('button, [role="button"], a');
      for (const el of nodes) {
        const t = texts(el);
        if (needles.some(n => t.includes(n))) {
          const r = el.getBoundingClientRect();
          if (r.width > 0 && r.height > 0) return el;
        }
      }
      for (const el of root.querySelectorAll('*')) {
        if (el.shadowRoot) {
          const found = walk(el.shadowRoot);
          if (found) return found;
        }
      }
      return null;
    }
    return walk(document);
    """
    return driver.execute_script(script)


def exportar_para_excel(driver) -> None:
    """
    Exporta a exploration aberta para Excel.
    Na Claro: desliga Editar -> seleciona o cubo -> clica 'Exportar para planilha'.
    """
    log("Acionando exportacao para Excel...")
    driver.switch_to.default_content()

    desligar_modo_editar(driver)
    focar_widget_cubo(driver)

    botao = None
    fim = time.time() + 25
    while time.time() < fim and botao is None:
        try:
            botao = achar_elemento(driver, "botao_exportar_planilha", timeout=2)
        except TimeoutException:
            botao = achar_botao_exportar_planilha_js(driver)
        if botao is None:
            # Re-foca o widget; a toolbar on-demand às vezes some.
            focar_widget_cubo(driver)
            time.sleep(1)

    if botao is None:
        log("Botao nao apareceu; tentando menu de contexto na grade...")
        try:
            grade = achar_elemento(driver, "grade_cubo", timeout=5, clicavel=False)
            ActionChains(driver).move_to_element(grade).context_click().perform()
            time.sleep(1.5)
        except Exception as e:
            log(f"Falha no clique direito: {e}")
        if not _clicar_item_menu_por_texto(
            driver, ["exportar para planilha", "export to spreadsheet", "exportar", "export"],
            timeout=8,
        ):
            raise TimeoutException(
                "Nao encontrei o botao/menu 'Exportar para planilha' na view aberta. "
                "Confirme que o modo Editar esta desligado e o cubo esta selecionado."
            )
    else:
        log("Clicando em 'Exportar para planilha'...")
        try:
            ActionChains(driver).move_to_element(botao).pause(0.3).click().perform()
        except Exception:
            clicar(driver, botao)
        time.sleep(3)

    # Se abrir submenu/dialog, escolhe Excel e confirma.
    if _clicar_item_menu_por_texto(driver, ["excel", "xlsx", "planilha"], timeout=5):
        time.sleep(2)
    try:
        confirmar = achar_elemento(driver, "confirmar_exportacao", timeout=6)
        clicar(driver, confirmar)
        time.sleep(2)
    except TimeoutException:
        pass
    log("Exportacao acionada; aguardando arquivo baixar.")


def aguardar_download(pasta: Path, arquivos_antes: set, timeout: int) -> Path:
    """Espera um arquivo novo terminar de baixar na pasta de downloads."""
    log(f"Aguardando download terminar (ate {timeout}s) em: {pasta}")
    fim = time.time() + timeout
    ultimo_ping = 0
    while time.time() < fim:
        atuais = {p for p in pasta.iterdir() if p.is_file()}
        novos = [
            p for p in atuais - arquivos_antes
            if p.suffix.lower() in EXTENSOES_VALIDAS
        ]
        baixando = [p for p in atuais if p.suffix in (".crdownload", ".tmp", ".partial")]
        if novos and not baixando:
            arquivo = max(novos, key=lambda p: p.stat().st_mtime)
            time.sleep(2)  # margem para o SO liberar o arquivo
            log(f"Download concluido: {arquivo.name}")
            return arquivo
        decorrido = int(timeout - (fim - time.time()))
        if decorrido - ultimo_ping >= 15:
            ultimo_ping = decorrido
            log(f"... ainda aguardando ({decorrido}s) — baixando={len(baixando)} novos={len(novos)}")
        time.sleep(2)
    raise TimeoutException(
        f"Tempo esgotado aguardando o download do arquivo em {pasta}. "
        "Verifique se o botao 'Exportar para planilha' realmente iniciou o download."
    )


def mover_para_destino(arquivo: Path, job: dict, pasta_backup: Path) -> Path:
    """Move o arquivo baixado para a pasta de rede, com backup do anterior."""
    destino_dir = Path(job["pasta_destino"])
    if not destino_dir.exists():
        raise FileNotFoundError(
            f"Pasta de destino inacessivel: {destino_dir}\n"
            "Verifique a conexao com a rede corporativa (VPN)."
        )

    nome_final = job.get("nome_arquivo_destino") or arquivo.name
    destino = destino_dir / nome_final

    if destino.exists():
        backup_dir = pasta_backup / job["nome"].replace("/", "-")
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup = backup_dir / f"{datetime.now():%Y%m%d_%H%M%S}_{nome_final}"
        shutil.copy2(destino, backup)
        log(f"Backup do arquivo anterior salvo em: {backup}")

    shutil.copy2(arquivo, destino)
    arquivo.unlink(missing_ok=True)
    log(f"Arquivo atualizado em: {destino}")
    return destino


def main() -> int:
    parser = argparse.ArgumentParser(description="Baixa as bases do Cognos/Planning Analytics.")
    parser.add_argument("--config", default=str(BASE_DIR / "config.json"))
    parser.add_argument("--somente", default=None,
                        help="Executa apenas exportacoes cujo nome contenha este texto.")
    parser.add_argument("--sem-mover", action="store_true",
                        help="Baixa os arquivos mas nao copia para a pasta de rede.")
    parser.add_argument("--debug", action="store_true",
                        help="Salva screenshot/HTML/inputs quando falhar e para no 1o erro.")
    args = parser.parse_args()

    cfg = carregar_config(Path(args.config))
    pasta_downloads = BASE_DIR / cfg["pasta_downloads"]
    pasta_backup = BASE_DIR / cfg["pasta_backup"]
    pasta_debug = BASE_DIR / "debug"
    pasta_downloads.mkdir(parents=True, exist_ok=True)

    jobs = cfg["exportacoes"]
    if args.somente:
        filtro = args.somente.lower()
        jobs = [j for j in jobs if filtro in j["nome"].lower() or filtro in j["nome_busca"].lower()]
        if not jobs:
            log(f"Nenhuma exportacao corresponde ao filtro '{args.somente}'.")
            return 1

    log(f"Iniciando automacao: {len(jobs)} exportacao(oes).")
    driver = criar_driver(pasta_downloads)
    resultados = []
    try:
        driver.get(cfg["url"])
        aguardar_login(driver, cfg["timeout_login_segundos"])

        for job in jobs:
            log("=" * 60)
            log(f"Exportacao: {job['nome']}  (servidor {job['servidor']})")
            try:
                arquivos_antes = {p for p in pasta_downloads.iterdir() if p.is_file()}
                pesquisar_e_abrir(
                    driver,
                    job["nome_busca"],
                    pasta_debug=pasta_debug,
                    dashboard_id=job.get("dashboard_id"),
                    base_url=cfg["url"],
                )
                exportar_para_excel(driver)
                arquivo = aguardar_download(
                    pasta_downloads, arquivos_antes, cfg["timeout_download_segundos"]
                )
                if args.sem_mover:
                    log(f"(--sem-mover) Arquivo mantido em: {arquivo}")
                else:
                    mover_para_destino(arquivo, job, pasta_backup)
                resultados.append((job["nome"], "OK"))
                # Volta para a home para a proxima exportacao.
                driver.get(cfg["url"])
                time.sleep(6)
            except Exception as e:
                log(f"ERRO em '{job['nome']}': {e}")
                resultados.append((job["nome"], f"ERRO: {e}"))
                salvar_debug(
                    driver,
                    pasta_debug,
                    job["nome"].replace("/", "-")[:40],
                    erro=str(e),
                )
                if args.debug:
                    log("Modo --debug: navegador permanece aberto. Pressione ENTER neste terminal para encerrar.")
                    try:
                        input()
                    except EOFError:
                        time.sleep(60)
                    break
                # Volta para a home para tentar a proxima exportacao.
                driver.get(cfg["url"])
                time.sleep(8)
    finally:
        driver.quit()

    log("=" * 60)
    log("Resumo:")
    houve_erro = False
    for nome, status in resultados:
        log(f"  {nome}: {status}")
        if status != "OK":
            houve_erro = True
    return 1 if houve_erro else 0


if __name__ == "__main__":
    sys.exit(main())
