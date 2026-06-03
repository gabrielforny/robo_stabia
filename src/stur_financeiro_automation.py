from __future__ import annotations

import logging
import re
import time
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Optional

from playwright.sync_api import FrameLocator, Locator, Page


@dataclass(slots=True)
class ResultadoConferencia:
    encontrada: bool
    valor_bateu: bool = False
    descricao: str | None = None
    data_conferencia: str | None = None
    valor_tabela: Decimal | None = None
    mensagem: str = ""


class SturFinanceiroAutomation:
    """
    Novo fluxo do robô:
    Financeiro -> Conferências e Baixas -> Conferências e Baixas por Conferência.

    Nesta versão o fluxo vai até:
    - buscar por Descrição;
    - se houver resultado, refinar por Dt.Conferência com Manter Pesquisa;
    - validar a coluna Valor ignorando sinal;
    - clicar no primeiro ícone Editar da linha correta.
    """

    IFRAME_SELECTOR = "#sturweb"

    MENU_FINANCEIRO = "#UsrMenu1_Menu1n8"
    MENU_CONFERENCIAS_E_BAIXAS = "a[href*=\"FinanceirobarraConfer\"]"
    MENU_CONFERENCIAS_POR_CONFERENCIA = "a[href*=\"ListaConferencias.aspx?fcf\"]"

    CAMPO_BUSCA = "#c0_PH1_UsrPesquisaRapidaLista1_EdtPesquisa"
    CHECK_MANTER_PESQUISA = "#c0_PH1_UsrPesquisaRapidaLista1_ChkManterPedquisa"
    BOTAO_LIMPAR_FILTROS = "#c0_PH1_UsrCabecLista1_ImgCancelarFiltro, input[value='Limpar'], input[title*='Limpar']"
    GRID = "#c0_PH1_GridView1"

    def __init__(self, page: Page, logger: logging.Logger, espera_padrao_segundos: int = 3):
        self.page = page
        self.logger = logger
        self.espera_padrao_segundos = espera_padrao_segundos

    def _frame(self) -> FrameLocator:
        return self.page.frame_locator(self.IFRAME_SELECTOR)

    def esperar(self, motivo: str = "") -> None:
        if motivo:
            self.logger.info("Aguardando %ss | %s", self.espera_padrao_segundos, motivo)
        time.sleep(self.espera_padrao_segundos)

    # ==========================================================
    # NAVEGAÇÃO
    # ==========================================================
    def acessar_tela_conferencias_baixas(self) -> None:
        self.logger.info("Acessando Financeiro -> Conferências e Baixas -> Conferências e Baixas por Conferência")

        try:
            self.page.locator(self.MENU_FINANCEIRO).hover()
            self.esperar("hover no menu Financeiro")

            self.page.locator(self.MENU_CONFERENCIAS_E_BAIXAS).hover()
            self.esperar("hover no submenu Conferências e Baixas")

            self.page.locator(self.MENU_CONFERENCIAS_POR_CONFERENCIA).click()
            self.esperar("clique em Conferências e Baixas por Conferência")
        except Exception:
            self.logger.warning("Menu por hover falhou. Tentando navegar direto via JavaScript.", exc_info=True)
            self.page.evaluate("Redirecionar('ListaConferencias.aspx?fcf')")
            self.esperar("redirecionamento direto para ListaConferencias")

        self.aguardar_tela_conferencias()
        self.logger.info("Tela de Conferências e Baixas carregada.")

    def aguardar_tela_conferencias(self) -> None:
        frame = self._frame()
        frame.locator(self.CAMPO_BUSCA).wait_for(state="visible", timeout=30000)
        frame.locator(self.GRID).wait_for(state="visible", timeout=30000)

    # ==========================================================
    # FILTROS / BUSCA
    # ==========================================================
    def limpar_filtros_com_calma(self) -> None:
        frame = self._frame()
        self.logger.info("Limpando filtros da tela de conferências")

        try:
            manter = frame.locator(self.CHECK_MANTER_PESQUISA)
            if manter.count() > 0 and manter.first.is_checked():
                manter.first.uncheck(force=True)
                self.esperar("desmarcar Manter Pesquisa")
        except Exception:
            self.logger.warning("Falha ao desmarcar Manter Pesquisa", exc_info=True)

        try:
            botao_limpar = frame.locator(self.BOTAO_LIMPAR_FILTROS).first
            if botao_limpar.count() > 0:
                botao_limpar.click()
                self.esperar("clicar em limpar filtros")
                return
        except Exception:
            self.logger.warning("Falha ao clicar em limpar filtros", exc_info=True)

        try:
            campo = frame.locator(self.CAMPO_BUSCA)
            campo.fill("")
            campo.press("Enter")
            self.esperar("limpar search manualmente")
        except Exception:
            self.logger.warning("Falha ao limpar search manualmente", exc_info=True)

    def clicar_coluna(self, nome_coluna: str) -> None:
        frame = self._frame()
        self.logger.info("Clicando na coluna: %s", nome_coluna)

        coluna = frame.locator(
            f"{self.GRID} th a",
            has_text=re.compile(rf"^\s*{re.escape(nome_coluna)}\s*", re.I),
        ).first
        coluna.wait_for(state="visible", timeout=15000)
        coluna.click()
        self.esperar(f"coluna {nome_coluna} selecionada")

    def preencher_search(self, valor: str) -> None:
        frame = self._frame()
        valor = str(valor or "").strip()

        self.logger.info("Preenchendo search: %s", valor)
        campo = frame.locator(self.CAMPO_BUSCA)
        campo.wait_for(state="visible", timeout=15000)
        campo.fill("")
        self.esperar("search limpo")
        campo.fill(valor)
        self.esperar("search preenchido")
        campo.press("Enter")
        self.esperar("pesquisa executada")

    def clicar_manter_pesquisa(self) -> None:
        frame = self._frame()
        manter = frame.locator(self.CHECK_MANTER_PESQUISA)

        if manter.count() == 0:
            self.logger.warning("Checkbox Manter Pesquisa não encontrado")
            return

        if not manter.first.is_checked():
            self.logger.info("Marcando Manter Pesquisa")
            manter.first.check(force=True)
            self.esperar("Manter Pesquisa marcado")
        else:
            self.logger.info("Manter Pesquisa já estava marcado")

    # ==========================================================
    # FLUXO PRINCIPAL
    # ==========================================================
    def buscar_conferencia_por_descricao_e_data(
        self,
        descricao: str,
        data_conferencia: str | None,
        valor_excel,
        clicar_editar_quando_bater: bool = True,
    ) -> ResultadoConferencia:
        self.limpar_filtros_com_calma()

        self.clicar_coluna("Descrição")
        self.preencher_search(descricao)

        resultados_iniciais = self.coletar_resultados_da_tabela()
        if not resultados_iniciais:
            return ResultadoConferencia(
                encontrada=False,
                mensagem=f"NÃO LOCALIZADO | Nenhuma conferência para descrição: {descricao}",
            )

        if data_conferencia:
            self.clicar_manter_pesquisa()
            self.clicar_coluna("Dt.Conferência")
            self.preencher_search(data_conferencia)

        resultados = self.coletar_resultados_da_tabela()
        if not resultados:
            return ResultadoConferencia(
                encontrada=False,
                mensagem=f"NÃO LOCALIZADO | Nenhuma conferência para descrição/data: {descricao} / {data_conferencia}",
            )

        valor_excel_decimal = self._parse_valor_decimal(valor_excel)
        if valor_excel_decimal is None:
            return ResultadoConferencia(
                encontrada=True,
                valor_bateu=False,
                mensagem=f"ERRO | Valor do Excel inválido para comparação: {valor_excel}",
            )

        for item in resultados:
            valor_tabela = item.get("Valor Decimal")
            self.logger.info(
                "Validando candidato | Descrição=%s | Dt.Conferência=%s | ValorTabela=%s | ValorExcel=%s",
                item.get("Descrição"),
                item.get("Dt.Conferência"),
                valor_tabela,
                valor_excel_decimal,
            )

            if valor_tabela is not None and abs(valor_tabela) == abs(valor_excel_decimal):
                self.logger.info("Valor bateu ignorando sinal. Conferência encontrada.")

                if clicar_editar_quando_bater:
                    self.clicar_editar_linha(item["__linha_locator"])

                return ResultadoConferencia(
                    encontrada=True,
                    valor_bateu=True,
                    descricao=item.get("Descrição"),
                    data_conferencia=item.get("Dt.Conferência"),
                    valor_tabela=valor_tabela,
                    mensagem=(
                        f"OK | Conferência encontrada | Descrição={item.get('Descrição')} | "
                        f"Dt.Conferência={item.get('Dt.Conferência')} | Valor={item.get('Valor')}"
                    ),
                )

        return ResultadoConferencia(
            encontrada=True,
            valor_bateu=False,
            mensagem="ENCONTRADO SEM MATCH | Conferências encontradas, mas nenhum valor bateu com o Excel",
        )

    # ==========================================================
    # TABELA
    # ==========================================================
    def coletar_resultados_da_tabela(self) -> list[dict]:
        frame = self._frame()
        grid = frame.locator(self.GRID)
        grid.wait_for(state="visible", timeout=15000)

        headers = self._obter_headers_grid()
        self.logger.info("Headers detectados: %s", headers)

        linhas = grid.locator("tr.g")
        total_linhas = linhas.count()
        self.logger.info("Quantidade de linhas retornadas: %s", total_linhas)

        resultados: list[dict] = []
        for index in range(total_linhas):
            linha = linhas.nth(index)
            valores = self._obter_valores_linha(linha)
            dados = self._mapear_linha_por_headers(headers, valores)
            dados["__linha_locator"] = linha
            dados["Valor Decimal"] = self._parse_valor_decimal(dados.get("Valor"))

            self.logger.info(
                "Linha coletada | Descrição=%s | Dt.Conferência=%s | Valor=%s | ValorDecimal=%s",
                dados.get("Descrição"),
                dados.get("Dt.Conferência"),
                dados.get("Valor"),
                dados.get("Valor Decimal"),
            )
            resultados.append(dados)

        return resultados

    def _obter_headers_grid(self) -> list[str]:
        frame = self._frame()
        ths = frame.locator(f"{self.GRID} tr").first.locator("th")
        headers: list[str] = []

        for i in range(ths.count()):
            th = ths.nth(i)
            texto = self._normalizar_texto(th.inner_text())
            colspan = th.get_attribute("colspan")
            qtd_colunas = int(colspan) if colspan and colspan.isdigit() else 1

            if not texto:
                texto = f"__COLUNA_VAZIA_{i}"

            for indice in range(qtd_colunas):
                headers.append(f"{texto}_{indice + 1}" if qtd_colunas > 1 else texto)

        return headers

    def _obter_valores_linha(self, linha: Locator) -> list[str]:
        tds = linha.locator("td")
        valores: list[str] = []
        for i in range(tds.count()):
            valores.append(self._normalizar_texto(tds.nth(i).inner_text()))
        return valores

    def _mapear_linha_por_headers(self, headers: list[str], valores: list[str]) -> dict:
        dados: dict = {}
        for index, valor in enumerate(valores):
            chave = headers[index] if index < len(headers) else f"__EXTRA_{index}"
            dados[chave] = valor
        return dados

    def clicar_editar_linha(self, linha: Locator) -> None:
        self.logger.info("Clicando no primeiro ícone Editar da linha encontrada")
        editar = linha.locator("input[type='image'][id*='ImgEditar']").first
        editar.wait_for(state="visible", timeout=15000)
        editar.click()
        self.esperar("abrir tela de edição")

    # ==========================================================
    # UTILITÁRIOS
    # ==========================================================
    def _normalizar_texto(self, texto: str | None) -> str:
        if texto is None:
            return ""
        texto = texto.replace("\xa0", " ")
        texto = re.sub(r"\s+", " ", texto)
        return texto.strip()

    def _parse_valor_decimal(self, valor) -> Optional[Decimal]:
        if valor is None:
            return None

        texto = str(valor).strip()
        if not texto:
            return None

        texto = texto.replace("R$", "").replace(" ", "")
        texto = texto.replace(".", "").replace(",", ".")
        texto = re.sub(r"[^0-9\.\-]", "", texto)

        if texto in {"", "-", ".", "-."}:
            return None

        try:
            return Decimal(texto)
        except InvalidOperation:
            return None
