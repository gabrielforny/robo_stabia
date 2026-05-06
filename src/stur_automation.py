from decimal import Decimal, InvalidOperation
from logging import Logger
from pathlib import Path
import re
from datetime import datetime

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

from config import AppConfig
from models import ResultadoVenda, Transacao


SELECTORS = {
    # Login
    "login_usuario": "#c0_PH1_EdtUsuario",
    "login_senha": "#c0_PH1_EdtSenha",
    "login_botao": "#c0_PH1_BtnLogin",
    "usuario_ativo_msg": "#c0_PH1_Label1",
    "usuario_ativo_desbloquear": "#c0_PH1_LinkButton1",

    # Menu - TROCAR conforme a tela real
    "menu_operacional": "#UsrMenu1_Menu1n4",
    "menu_vendas": "#waM61",

    # Busca em vendas - TROCAR conforme a tela real
    "campo_busca_vendas": "input[type='search'], input[name*='search'], input[id*='search']",
    "botao_buscar_vendas": "button:has-text('Buscar'), input[value='Buscar']",

    # Grid/resultado - TROCAR conforme a tela real
    "linha_resultado": "table tbody tr",
    "mensagem_sem_resultado": "text=Nenhum registro encontrado",

    # Detalhe venda - TROCAR conforme a tela real
    "primeiro_icone_item": "table tbody tr:first-child a, table tbody tr:first-child img",
    "aba_pagamento_fornecedor": "text=Pagamento do Fornecedor",
    "botao_editar": "text=Editar",
    "select_opcao_pagamento": "select[name*='pagamento'], select[id*='pagamento']",
    "select_fornecedor": "select[name*='fornecedor'], select[id*='fornecedor']",
    "campo_data": "input[name*='data'], input[id*='data']",
    "botao_ok": "text=OK",
    "botao_gravar": "text=GRAVAR",
}


class SturAutomation:
    def __init__(self, config: AppConfig, logger: Logger, headless: bool = False):
        self.config = config
        self.logger = logger
        self.headless = headless
        self._playwright = None
        self._browser = None
        self._context = None
        self.page: Page | None = None

    def __enter__(self) -> "SturAutomation":
        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.launch(headless=self.headless)
        self._context = self._browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 768},
        )
        self.page = self._context.new_page()
        self.page.set_default_timeout(15000)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._context:
            self._context.close()
        if self._browser:
            self._browser.close()
        if self._playwright:
            self._playwright.stop()

    def login(self) -> None:
        page = self._page()

        self.logger.info("Acessando STUR: %s", self.config.stur_url)
        page.goto(self.config.stur_url, wait_until="domcontentloaded")

        page.locator(SELECTORS["login_usuario"]).fill(self.config.stur_user)
        page.locator(SELECTORS["login_senha"]).fill(self.config.stur_password)
        page.locator(SELECTORS["login_botao"]).click()

        page.wait_for_load_state("networkidle")

        if self._existe(SELECTORS["usuario_ativo_msg"], timeout=3000):
            texto_msg = page.locator(SELECTORS["usuario_ativo_msg"]).inner_text().strip()

            if "Usuário ativo em outra sessão" in texto_msg:
                self.logger.warning("Usuário ativo em outra sessão. Desbloqueando usuário...")

                page.locator(SELECTORS["usuario_ativo_desbloquear"]).click()
                page.wait_for_load_state("networkidle")

                self.logger.info("Usuário desbloqueado e login realizado.")

        self.logger.info("Login finalizado com sucesso.")

    def acessar_vendas(self) -> None:
        page = self._page()

        self.logger.info("Acessando menu Operacional -> Vendas")

        page.locator(SELECTORS["menu_operacional"]).hover()
        page.wait_for_timeout(500)

        page.locator(SELECTORS["menu_vendas"]).click()
        page.wait_for_load_state("networkidle")

        self.logger.info("Tela de vendas acessada.")

    def buscar_venda(self, codigo_companhia: str) -> ResultadoVenda:
        page = self._page()

        self.logger.info("Buscando venda pelo código/localizador: %s", codigo_companhia)

        campo_busca = page.locator(SELECTORS["campo_busca_vendas"]).first
        campo_busca.fill("")
        campo_busca.fill(codigo_companhia)

        try:
            page.locator(SELECTORS["botao_buscar_vendas"]).first.click()
        except PlaywrightTimeoutError:
            campo_busca.press("Enter")

        page.wait_for_load_state("networkidle")

        if self._existe(SELECTORS["mensagem_sem_resultado"], timeout=3000):
            return ResultadoVenda(
                encontrada=False,
                mensagem=f"Venda não encontrada para código {codigo_companhia}",
            )

        linhas = page.locator(SELECTORS["linha_resultado"])

        if linhas.count() == 0:
            return ResultadoVenda(
                encontrada=False,
                mensagem=f"Nenhuma linha retornada para código {codigo_companhia}",
            )

        primeira_linha = linhas.first
        texto_linha = primeira_linha.inner_text(timeout=5000)

        codigo_venda = self._extrair_codigo_venda(texto_linha)
        total_fornecedor = self._extrair_total_fornecedor(texto_linha)

        return ResultadoVenda(
            encontrada=True,
            codigo_venda=codigo_venda,
            total_fornecedor=total_fornecedor,
            mensagem="Venda encontrada.",
        )

    def processar_pagamento_fornecedor(self, transacao: Transacao) -> None:
        """
        Fluxo base.
        Troque os selectors e complemente campos conforme o vídeo/tela real.
        """
        page = self._page()

        self.logger.info(
            "Processando pagamento do fornecedor para código %s",
            transacao.codigo_companhia,
        )

        page.locator(SELECTORS["primeiro_icone_item"]).first.click()
        page.wait_for_load_state("networkidle")

        page.locator(SELECTORS["aba_pagamento_fornecedor"]).first.click()
        page.wait_for_load_state("networkidle")

        page.locator(SELECTORS["botao_editar"]).first.click()

        # Exemplo: selecionar primeira opção válida.
        # Depois você pode trocar por valor específico.
        if self._existe(SELECTORS["select_opcao_pagamento"], timeout=3000):
            page.locator(SELECTORS["select_opcao_pagamento"]).first.select_option(index=1)

        if self._existe(SELECTORS["select_fornecedor"], timeout=3000):
            page.locator(SELECTORS["select_fornecedor"]).first.select_option(index=1)

        # Exemplo: alterar data somente se precisar.
        # page.locator(SELECTORS["campo_data"]).first.fill("15/04/2026")

        page.locator(SELECTORS["botao_ok"]).first.click()
        page.wait_for_load_state("networkidle")

        page.locator(SELECTORS["botao_gravar"]).first.click()
        page.wait_for_load_state("networkidle")

        self.logger.info("Pagamento do fornecedor gravado para %s", transacao.codigo_companhia)

    def salvar_screenshot_erro(self, codigo: str) -> Path | None:
        if not self.config.salvar_screenshot_erro:
            return None

        page = self._page()
        screenshots_dir = self.config.logs_dir / "screenshots"
        screenshots_dir.mkdir(exist_ok=True)

        arquivo = screenshots_dir / f"erro_{codigo}_{datetime.now():%Y%m%d_%H%M%S}.png"
        page.screenshot(path=str(arquivo), full_page=True)

        return arquivo

    def _page(self) -> Page:
        if self.page is None:
            raise RuntimeError("Browser não inicializado.")
        return self.page

    def _existe(self, selector: str, timeout: int = 1000) -> bool:
        try:
            self._page().locator(selector).first.wait_for(state="visible", timeout=timeout)
            return True
        except PlaywrightTimeoutError:
            return False

    def _extrair_codigo_venda(self, texto_linha: str) -> str | None:
        """
        Ajuste quando souber a coluna exata.
        Por enquanto pega o primeiro número grande da linha.
        """
        match = re.search(r"\b\d{4,}\b", texto_linha)
        return match.group(0) if match else None

    def _extrair_total_fornecedor(self, texto_linha: str) -> Decimal | None:
        """
        Ajuste ideal:
        - pegar célula da coluna "Total do Fornecedor" diretamente no grid.

        Por enquanto:
        - procura valores no formato brasileiro na linha
        - retorna o último valor monetário encontrado
        """
        matches = re.findall(r"(?:R\$\s*)?[-+]?\d{1,3}(?:\.\d{3})*,\d{2}", texto_linha)

        if not matches:
            return None

        return self._parse_decimal(matches[-1])

    def _parse_decimal(self, value: str) -> Decimal | None:
        texto = str(value).replace("R$", "").replace(" ", "").strip()

        if "," in texto and "." in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")

        texto = re.sub(r"[^0-9.-]", "", texto)

        try:
            return Decimal(texto)
        except InvalidOperation:
            return None