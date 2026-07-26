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
    "menu_hamburguer": [
        (By.CSS_SELECTOR, "button[aria-label*='Navigation']"),
        (By.CSS_SELECTOR, "button[aria-label*='avegação']"),
        (By.CSS_SELECTOR, "button[aria-label*='avegacao']"),
        (By.CSS_SELECTOR, "button[aria-label*='Main menu']"),
        (By.CSS_SELECTOR, "button[aria-label*='Menu']"),
        (By.XPATH, "//button[contains(@aria-label,'menu') or contains(@aria-label,'Menu')]"),
        (By.XPATH, "//header//button[1]"),
    ],
    "link_compartilhado": [
        (By.XPATH, "//*[self::a or self::button or self::span or self::div][normalize-space(.)='Compartilhado']"),
        (By.XPATH, "//*[contains(@aria-label,'Compartilhado')]"),
        (By.XPATH, "//*[self::a or self::button or self::span][normalize-space(.)='Shared']"),
        (By.XPATH, "//*[contains(text(),'Compartilhado')]"),
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
    ],
    "menu_exportar": [
        (By.XPATH, "//*[self::button or self::span or self::a][contains(., 'Exportar')]"),
        (By.XPATH, "//*[self::button or self::span or self::a][contains(., 'Export')]"),
        (By.CSS_SELECTOR, "button[aria-label*='xport']"),
    ],
    "opcao_excel": [
        (By.XPATH, "//*[contains(text(), 'Excel')]"),
        (By.XPATH, "//*[contains(text(), 'xlsx')]"),
    ],
    "confirmar_exportacao": [
        (By.XPATH, "//button[contains(., 'Exportar')]"),
        (By.XPATH, "//button[contains(., 'OK')]"),
        (By.XPATH, "//button[contains(., 'Export')]"),
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


def abrir_compartilhado(driver) -> None:
    """Garante que a sidebar Compartilhado (com o campo Pesquisar) esteja aberta."""
    # Se o campo Pesquisar ja esta visivel, estamos no lugar certo.
    try:
        achar_elemento(driver, "campo_pesquisa", timeout=3, clicavel=False)
        log("Sidebar Compartilhado ja aberta.")
        return
    except TimeoutException:
        pass

    log("Abrindo pasta Compartilhado...")
    try:
        menu = achar_elemento(driver, "menu_hamburguer", timeout=10)
        clicar(driver, menu)
        time.sleep(1.5)
    except TimeoutException:
        log("Menu hamburguer nao encontrado; tentando link Compartilhado direto.")

    link = achar_elemento(driver, "link_compartilhado", timeout=20)
    clicar(driver, link)
    time.sleep(3)

    # Confirma que o campo Pesquisar apareceu.
    achar_elemento(driver, "campo_pesquisa", timeout=20, clicavel=False)
    log("Compartilhado aberto.")


def pesquisar_e_abrir(driver, nome_busca: str) -> None:
    """Abre Compartilhado, pesquisa a view e clica no resultado."""
    abrir_compartilhado(driver)

    log(f"Pesquisando: {nome_busca}")
    campo = achar_elemento(driver, "campo_pesquisa", timeout=20)
    clicar(driver, campo)
    # Limpa o campo de forma robusta (clear() às vezes falha em inputs React).
    campo.send_keys(Keys.CONTROL, "a")
    campo.send_keys(Keys.BACKSPACE)
    time.sleep(0.3)
    campo.send_keys(nome_busca)
    time.sleep(1)
    campo.send_keys(Keys.ENTER)
    time.sleep(4)

    resultado = achar_resultado(driver, nome_busca, timeout=30)
    clicar(driver, resultado)
    log("View aberta. Aguardando carregar...")
    time.sleep(10)


def achar_resultado(driver, nome_busca: str, timeout: int = 30):
    """
    Localiza o resultado da pesquisa.
    A UI do PA trunca o texto (ex.: 'Receita DRE P... V2 (irat950)'),
    entao tenta match exato, title/aria-label e contains parcial.
    """
    literais = xpath_literal(nome_busca)
    # Trechos distintivos do nome (evita depender do texto truncado).
    trechos = [nome_busca]
    if "(" in nome_busca and nome_busca.endswith(")"):
        trechos.append(nome_busca[nome_busca.rfind("(") :])  # ex.: (irat950)
    if " - " in nome_busca:
        trechos.append(nome_busca.split(" - ")[0].strip())

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


def exportar_para_excel(driver) -> None:
    """Aciona a exportacao para Excel na view aberta."""
    log("Acionando exportacao para Excel...")
    menu = achar_elemento(driver, "menu_exportar", timeout=30)
    menu.click()
    time.sleep(2)
    try:
        opcao = achar_elemento(driver, "opcao_excel", timeout=10)
        opcao.click()
        time.sleep(2)
    except TimeoutException:
        log("Opcao 'Excel' nao apareceu; seguindo (a exportacao pode ja ter iniciado).")
    try:
        confirmar = achar_elemento(driver, "confirmar_exportacao", timeout=5)
        confirmar.click()
    except TimeoutException:
        pass


def aguardar_download(pasta: Path, arquivos_antes: set, timeout: int) -> Path:
    """Espera um arquivo novo terminar de baixar na pasta de downloads."""
    log("Aguardando download terminar...")
    fim = time.time() + timeout
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
        time.sleep(2)
    raise TimeoutException("Tempo esgotado aguardando o download do arquivo.")


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
    args = parser.parse_args()

    cfg = carregar_config(Path(args.config))
    pasta_downloads = BASE_DIR / cfg["pasta_downloads"]
    pasta_backup = BASE_DIR / cfg["pasta_backup"]
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
                pesquisar_e_abrir(driver, job["nome_busca"])
                exportar_para_excel(driver)
                arquivo = aguardar_download(
                    pasta_downloads, arquivos_antes, cfg["timeout_download_segundos"]
                )
                if args.sem_mover:
                    log(f"(--sem-mover) Arquivo mantido em: {arquivo}")
                else:
                    mover_para_destino(arquivo, job, pasta_backup)
                resultados.append((job["nome"], "OK"))
            except Exception as e:
                log(f"ERRO em '{job['nome']}': {e}")
                resultados.append((job["nome"], f"ERRO: {e}"))
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
