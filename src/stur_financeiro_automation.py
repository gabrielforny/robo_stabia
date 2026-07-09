from __future__ import annotations

import logging
import re
import time
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Optional

from playwright.sync_api import FrameLocator, Locator, Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError


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
        self.logger.info("Navegando para Conferências e Baixas por Conferência")
        self.page.locator(self.MENU_FINANCEIRO).wait_for(state="visible", timeout=30000)

        # Tenta clicar pelo menu real; se falhar usa Redirecionar com retry
        navegou = False
        try:
            self.page.locator(self.MENU_FINANCEIRO).hover()
            self.esperar("hover menu Financeiro")
            self.page.locator(self.MENU_CONFERENCIAS_E_BAIXAS).click()
            self.esperar("submenu Conferências aberto")
            self.page.locator(self.MENU_CONFERENCIAS_POR_CONFERENCIA).click()
            navegou = True
            self.logger.info("Navegação por menu concluída.")
        except Exception:
            self.logger.warning("Hover/click no menu falhou. Tentando via Redirecionar.")

        if not navegou:
            for tentativa in range(1, 4):
                try:
                    self.page.wait_for_load_state("domcontentloaded", timeout=10000)
                    self.page.evaluate("Redirecionar('ListaConferencias.aspx?fcf')")
                    navegou = True
                    break
                except Exception as exc:
                    self.logger.warning(
                        "Redirecionar falhou na tentativa %d: %s", tentativa, exc
                    )
                    self.esperar("aguardar antes de nova tentativa de navegação")

        if not navegou:
            raise RuntimeError("Não foi possível navegar para a tela de Conferências após 3 tentativas.")

        self.esperar("redirecionamento para ListaConferencias")
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
    # FLUXO LATAM — CONFERÊNCIAS
    # ==========================================================
    def buscar_ou_criar_conferencia_latam(
        self,
        descricao_busca: str,
        descricao_criar: str,
        data_fatura: str,
    ) -> None:
        """Busca conferência pelo termo. Se existir, abre em edição; se não, cria nova."""
        self.limpar_filtros_com_calma()
        self.clicar_coluna("Descrição")
        self.preencher_search(descricao_busca)

        resultados = self.coletar_resultados_da_tabela()

        if resultados:
            self.logger.info("Conferência '%s' encontrada. Abrindo em edição.", descricao_busca)
            self.clicar_editar_linha(resultados[0]["__linha_locator"])
        else:
            self.logger.info("Conferência '%s' não encontrada. Criando nova.", descricao_busca)
            self._criar_nova_conferencia(descricao_criar, data_fatura)

        self.esperar("conferência aberta")

    def _criar_nova_conferencia(self, descricao: str, data_fatura: str) -> None:
        frame = self._frame()
        self.logger.info("Criando nova conferência | Descrição=%s | Fatura=%s", descricao, data_fatura)

        img_novo = frame.locator("#c0_PH1_UsrCabecLista1_ImgNovo")
        img_novo.wait_for(state="visible", timeout=15000)
        img_novo.click()
        self.esperar("clique em ImgNovo")

        edt_desc = frame.locator("#c0_PH1_EdtDescricao")
        edt_desc.wait_for(state="visible", timeout=15000)
        edt_desc.fill(descricao)

        if data_fatura:
            edt_fatura = frame.locator("#c0_PH1_EdtFatura")
            edt_fatura.wait_for(state="visible", timeout=10000)
            edt_fatura.fill(data_fatura)

        self.esperar("campos conferência preenchidos")

    def abrir_adicionar_titulos(self) -> None:
        frame = self._frame()
        self.logger.info("Clicando em Adicionar Títulos")

        btn_funcao = frame.locator("#c0_PH1_UsrRodapeEdicao1_BtnFuncao")
        btn_funcao.wait_for(state="visible", timeout=15000)
        btn_funcao.click()
        self.esperar("Adicionar Títulos clicado")

        # aguarda sub-tela carregar e seleciona Clientes
        frame2 = self._frame()
        rad_cliente = frame2.locator("#c0_PH1_RadCliente")
        rad_cliente.wait_for(state="visible", timeout=20000)
        rad_cliente.click()
        self.esperar("Clientes selecionado")

        btn_filtrar = frame2.locator("#c0_PH1_BtnFiltroTitulosOk")
        btn_filtrar.wait_for(state="visible", timeout=15000)
        btn_filtrar.click()
        self.esperar("Filtrar clicado")
        self.esperar("aguardando lista de títulos")

    def garantir_coluna_localizador_visivel(self) -> None:
        frame = self._frame()
        self.logger.info("Verificando coluna Localizador na sub-tela de títulos")

        col = frame.locator(
            f"{self.GRID} th a",
            has_text=re.compile(r"^\s*Localizador\s*$", re.I),
        )

        if col.count() > 0 and col.first.is_visible():
            self.logger.info("Coluna Localizador já visível")
            return

        self.logger.info("Habilitando coluna Localizador via ícone de colunas visíveis")
        img_eye = frame.locator("#c0_PH1_ImgColunasVisiveis")
        img_eye.wait_for(state="visible", timeout=15000)
        img_eye.click()
        self.esperar("painel colunas aberto")

        chk_loc = frame.locator("#c0_PH1_ChkLocalizador")
        chk_loc.wait_for(state="visible", timeout=10000)
        if not chk_loc.is_checked():
            chk_loc.check(force=True)
            self.esperar("Localizador marcado")

        img_eye.click()
        self.esperar("painel colunas fechado")

    def buscar_e_selecionar_localizador(
        self,
        localizador: str,
        valor_excel: Decimal | None = None,
    ) -> tuple[bool, str]:
        """
        Busca o localizador na sub-tela de títulos, valida o 'Valor Oficial' contra
        valor_excel (ignorando sinal) e, se bater, marca o ChkSelecionado.

        Retorna (sucesso, mensagem).
        """
        self.logger.info("Buscando localizador nos títulos: %s", localizador)

        self.clicar_coluna("Localizador")
        self.preencher_search(localizador)

        frame = self._frame()
        grid = frame.locator(self.GRID)
        grid.wait_for(state="visible", timeout=15000)

        # Debug: loga headers e total de linhas visíveis antes de filtrar por ChkSelecionado
        headers_debug = self._obter_headers_grid()
        self.logger.info("[DEBUG] Headers da grid de títulos: %s", headers_debug)
        todas_linhas = grid.locator("tr")
        self.logger.info("[DEBUG] Total de <tr> na grid após busca: %d", todas_linhas.count())
        for i in range(min(todas_linhas.count(), 5)):
            self.logger.info("[DEBUG] Linha %d texto: %s", i, self._normalizar_texto(todas_linhas.nth(i).inner_text()))

        # Títulos usam tr.d (ou tr.g); buscamos qualquer linha que tenha ChkSelecionado
        linhas = grid.locator("tr:has([id*='ChkSelecionado'])")

        if linhas.count() == 0:
            self.logger.warning("Localizador %s não encontrado nos títulos", localizador)
            self.limpar_filtros_com_calma()
            return False, f"localizador {localizador} não encontrado nos títulos disponíveis"

        linha = linhas.first

        # Valida Valor Oficial antes de marcar
        if valor_excel is not None:
            valor_tabela = self._obter_valor_oficial_da_linha(linha)
            if valor_tabela is not None:
                if abs(valor_tabela) != abs(valor_excel):
                    self.logger.warning(
                        "Valor não bate | localizador=%s | tabela=%s | excel=%s",
                        localizador, valor_tabela, valor_excel,
                    )
                    self.limpar_filtros_com_calma()
                    return False, (
                        f"valor não bate | Valor Oficial={valor_tabela} | Excel={valor_excel}"
                    )
                self.logger.info(
                    "Valor conferido | localizador=%s | %s = %s", localizador, valor_tabela, valor_excel
                )
            else:
                self.logger.warning("Não foi possível ler Valor Oficial da linha; prosseguindo sem validação")

        chk = linha.locator("[id*='ChkSelecionado']")
        if chk.count() > 0:
            chk.first.check(force=True)
            self.esperar("localizador selecionado")
            self.logger.info("Localizador %s marcado com sucesso", localizador)
            self.limpar_filtros_com_calma()
            return True, "OK"

        self.logger.warning("ChkSelecionado não encontrado na linha do localizador %s", localizador)
        self.limpar_filtros_com_calma()
        return False, f"ChkSelecionado não encontrado na linha do localizador {localizador}"

    def _obter_valor_oficial_da_linha(self, linha: Locator) -> Decimal | None:
        """Lê a coluna 'Valor Oficial' da linha usando os headers do grid."""
        headers = self._obter_headers_grid()
        valor_idx = next(
            (i for i, h in enumerate(headers) if "valor oficial" in h.lower()),
            None,
        )

        if valor_idx is None:
            self.logger.warning("Coluna 'Valor Oficial' não encontrada nos headers: %s", headers)
            return None

        tds = linha.locator("td")
        if valor_idx >= tds.count():
            return None

        texto = self._normalizar_texto(tds.nth(valor_idx).inner_text())
        return self._parse_valor_decimal(texto)

    def gravar_titulos(self) -> None:
        frame = self._frame()
        self.logger.info("Gravando títulos selecionados")

        # type="button" para a sub-tela de títulos
        btn = frame.locator("#c0_PH1_UsrRodapeEdicao1_BtnGravar[type='button']")
        if btn.count() == 0:
            btn = frame.locator("#c0_PH1_UsrRodapeEdicao1_BtnGravar").first

        btn.wait_for(state="visible", timeout=15000)
        btn.click()
        self.esperar("títulos gravados")
        self.esperar("aguardando retorno à conferência")

    def gravar_conferencia(self) -> None:
        frame = self._frame()
        self.logger.info("Gravando conferência")

        # type="submit" na tela de edição da conferência
        btn = frame.locator("#c0_PH1_UsrRodapeEdicao1_BtnGravar[type='submit']")
        if btn.count() == 0:
            btn = frame.locator("#c0_PH1_UsrRodapeEdicao1_BtnGravar").first

        btn.wait_for(state="visible", timeout=15000)
        btn.click()
        self.esperar("conferência gravada")
        self.esperar("aguardando retorno à lista")

    # ==========================================================
    # FLUXO HOTEL — CONFERÊNCIAS
    # ==========================================================

    def buscar_ou_criar_conferencia_hotel(
        self,
        descricao_busca: str,
        descricao_criar: str,
        data_fatura: str,
    ) -> None:
        """Busca conferência de hotelaria. Se existir, abre em edição; se não, cria nova."""
        self.limpar_filtros_com_calma()
        self.clicar_coluna("Descrição")
        self.preencher_search(descricao_busca)

        resultados = self.coletar_resultados_da_tabela()

        if resultados:
            self.logger.info("Conferência '%s' encontrada. Abrindo em edição.", descricao_busca)
            self.clicar_editar_linha(resultados[0]["__linha_locator"])
        else:
            self.logger.info("Conferência '%s' não encontrada. Criando nova.", descricao_busca)
            self._criar_nova_conferencia(descricao_criar, data_fatura)

        self.esperar("conferência hotel aberta")

    def habilitar_coluna_dados_integracao(self) -> None:
        """Garante que a coluna 'Dados Integração' está visível na sub-tela de títulos."""
        frame = self._frame()
        self.logger.info("Verificando coluna Dados Integração na sub-tela de títulos")

        col = frame.locator(
            f"{self.GRID} th a",
            has_text=re.compile(r"dados\s+integra", re.I),
        )
        if col.count() > 0 and col.first.is_visible():
            self.logger.info("Coluna Dados Integração já visível")
            return

        self.logger.info("Habilitando coluna Dados Integração via ícone de colunas visíveis")
        img_eye = frame.locator("#c0_PH1_ImgColunasVisiveis")
        img_eye.wait_for(state="visible", timeout=15000)
        img_eye.click()
        self.esperar("painel colunas aberto")

        chk = frame.locator("#c0_PH1_ChkDadosIntegracao")
        if chk.count() == 0:
            self.logger.warning(
                "Checkbox #c0_PH1_ChkDadosIntegracao não encontrado — "
                "tentando localizar por label 'Dados Integração'"
            )
            label = frame.locator("label", has_text=re.compile(r"dados.integra", re.I)).first
            if label.count() > 0:
                chk = label.locator("xpath=preceding-sibling::input[@type='checkbox']")
                if chk.count() == 0:
                    chk = label.locator("xpath=../input[@type='checkbox']")

        chk.first.wait_for(state="visible", timeout=10000)
        if not chk.first.is_checked():
            chk.first.check(force=True)
            self.esperar("Dados Integração marcado")

        img_eye.click()
        self.esperar("painel colunas fechado")
        self.logger.info("Coluna Dados Integração habilitada com sucesso")

    def buscar_e_selecionar_dados_integracao(
        self,
        observacao: str,
        valor_excel=None,
    ) -> tuple[bool, str]:
        """
        Busca pelo campo 'Dados Integração' na sub-tela de títulos e marca ChkSelecionado.
        Retorna (sucesso, mensagem).
        """
        self.logger.info("Buscando Dados Integração nos títulos: %s", observacao)

        self.clicar_coluna("Dados Integração")
        self.preencher_search(observacao)

        frame = self._frame()
        grid = frame.locator(self.GRID)
        grid.wait_for(state="visible", timeout=15000)

        linhas = grid.locator("tr:has([id*='ChkSelecionado'])")

        if linhas.count() == 0:
            self.logger.warning("Dados Integração '%s' não encontrado nos títulos", observacao)
            self.limpar_filtros_com_calma()
            return False, f"Dados Integração '{observacao}' não encontrado nos títulos disponíveis"

        linha = linhas.first

        if valor_excel is not None:
            valor_tabela = self._obter_valor_oficial_da_linha(linha)
            if valor_tabela is not None:
                from decimal import Decimal
                valor_excel_dec = valor_excel if isinstance(valor_excel, Decimal) else self._parse_valor_decimal(str(valor_excel))
                if valor_excel_dec is not None and abs(valor_tabela) != abs(valor_excel_dec):
                    self.logger.warning(
                        "Valor não bate | Dados Integração=%s | tabela=%s | excel=%s",
                        observacao, valor_tabela, valor_excel,
                    )
                    self.limpar_filtros_com_calma()
                    return False, f"valor não bate | Valor Oficial={valor_tabela} | Excel={valor_excel}"
                self.logger.info(
                    "Valor conferido | Dados Integração=%s | %s = %s", observacao, valor_tabela, valor_excel
                )
            else:
                self.logger.warning(
                    "Não foi possível ler Valor Oficial para Dados Integração=%s; prosseguindo sem validação",
                    observacao,
                )

        chk = linha.locator("[id*='ChkSelecionado']")
        if chk.count() > 0:
            chk.first.check(force=True)
            self.esperar("Dados Integração selecionado")
            self.logger.info("Dados Integração '%s' marcado com sucesso", observacao)
            self.limpar_filtros_com_calma()
            return True, "OK"

        self.logger.warning("ChkSelecionado não encontrado para Dados Integração '%s'", observacao)
        self.limpar_filtros_com_calma()
        return False, f"ChkSelecionado não encontrado para Dados Integração '{observacao}'"

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
