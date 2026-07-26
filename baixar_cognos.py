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
# Seletores da interface do Planning Analytics Workspace.
# ATENCAO: podem precisar de ajuste conforme a versao instalada na Claro.
# Cada entrada e uma lista de alternativas tentadas em ordem.
# ---------------------------------------------------------------------------
SELETORES = {
    "botao_pesquisa": [
        (By.CSS_SELECTOR, "button[aria-label*='esquis']"),   # Pesquisa / Pesquisar
        (By.CSS_SELECTOR, "button[aria-label*='Search']"),
        (By.CSS_SELECTOR, "[data-testid*='search'] button"),
    ],
    "campo_pesquisa": [
        (By.CSS_SELECTOR, "input[placeholder*='esquis']"),
        (By.CSS_SELECTOR, "input[placeholder*='Search']"),
        (By.CSS_SELECTOR, "input[type='search']"),
        (By.CSS_SELECTOR, "input[role='searchbox']"),
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


def achar_elemento(driver, chave_seletor: str, timeout: int = 20):
    """Tenta as alternativas de seletor em ordem ate encontrar um elemento clicavel."""
    fim = time.time() + timeout
    while time.time() < fim:
        for by, seletor in SELETORES[chave_seletor]:
            try:
                el = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((by, seletor))
                )
                return el
            except TimeoutException:
                continue
    raise TimeoutException(
        f"Nao encontrei o elemento '{chave_seletor}'. "
        "Os seletores em SELETORES provavelmente precisam de ajuste para esta versao do PA."
    )


def aguardar_login(driver, timeout: int) -> None:
    """Espera ate a home do PA carregar (apos login manual/SSO)."""
    log("Aguardando login... Se aparecer a tela de login, entre com seu usuario.")
    fim = time.time() + timeout
    while time.time() < fim:
        try:
            # A home carregada tem a barra superior com o botao de pesquisa.
            for by, seletor in SELETORES["botao_pesquisa"] + SELETORES["campo_pesquisa"]:
                if driver.find_elements(by, seletor):
                    log("Login concluido, pagina inicial carregada.")
                    time.sleep(3)
                    return
        except Exception:
            pass
        time.sleep(2)
    raise TimeoutException("Tempo esgotado aguardando o login no Planning Analytics.")


def pesquisar_e_abrir(driver, nome_busca: str) -> None:
    """Pesquisa o nome da view e abre o primeiro resultado com o nome exato."""
    log(f"Pesquisando: {nome_busca}")
    try:
        botao = achar_elemento(driver, "botao_pesquisa", timeout=10)
        botao.click()
        time.sleep(1)
    except TimeoutException:
        pass  # em algumas versoes o campo ja fica visivel sem clicar no icone

    campo = achar_elemento(driver, "campo_pesquisa", timeout=15)
    campo.clear()
    campo.send_keys(nome_busca)
    campo.send_keys(Keys.ENTER)
    time.sleep(4)

    # Clica no resultado cujo texto bate com o nome pesquisado.
    xpath_resultado = f"//*[normalize-space(text())={xpath_literal(nome_busca)}]"
    resultado = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, xpath_resultado))
    )
    resultado.click()
    log("View aberta. Aguardando carregar...")
    time.sleep(10)


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
