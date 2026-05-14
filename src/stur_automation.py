import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from logging import Logger
from pathlib import Path

from playwright.sync_api import FrameLocator, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

from config import AppConfig
from models import CandidatoVenda, Transacao


SELECTORS = {
    # Login
    "login_usuario": "#c0_PH1_EdtUsuario",
    "login_senha": "#c0_PH1_EdtSenha",
    "login_botao": "#c0_PH1_BtnLogin",
    "usuario_ativo_msg": "#c0_PH1_Label1",
    "usuario_ativo_desbloquear": "#c0_PH1_LinkButton1",

    # Menu principal
    "menu_operacional": "#UsrMenu1_Menu1n4",
    "menu_vendas": "#waM61",

    # Iframe e tela de vendas
    "iframe_stur": "#sturweb",
    "campo_busca": "#c0_PH1_UsrPesquisaRapidaLista1_EdtPesquisa",
    "botao_buscar": "#c0_PH1_UsrPesquisaRapidaLista1_BtnPesquisa",
    "manter_pesquisa": "#c0_PH1_UsrPesquisaRapidaLista1_ChkManterPedquisa",
    "limpar_filtros": "#c0_PH1_UsrCabecLista1_ImgCancelarFiltro",
    "grid": "#c0_PH1_GridView1",
    "linhas_grid": "#c0_PH1_GridView1 tbody tr",

    # Tela de edição da venda / pagamento fornecedor
    "editar_venda_linha": "#c0_PH1_GridView1 tbody tr input[id*='ImgEditar']",
    "grid_pagamentos_fornecedor": "#c0_PH1_UFRPV1_GrdP",
    "editar_pagamento_fornecedor": "#c0_PH1_UFRPV1_GrdP input[id*='ImgEditar']",
    "modal_forma_pagamento": "#c0_PH1_UFRPV1_PnlFormaPag",
    "radio_cartao_credito_agencia": "#c0_PH1_UFRPV1_RblFormaPag_2",
    "select_titular_cartao_agencia": "#c0_PH1_UFRPV1_UTCCRAG_Dbl",
    "select_numero_cartao_agencia": "#c0_PH1_UFRPV1_UCCRAG_Dbl",
    "data_vencimento_cartao_agencia": "#c0_PH1_UFRPV1_UPCCDV_Edt",
    "botao_ok_pagamento": "#c0_PH1_UFRPV1_BtnOKPag",
    "botao_gravar_venda": "#c0_PH1_URE1_BtnGravar",
    "botao_voltar_venda": "#c0_PH1_URE1_BtnCancelar",
}

ESPERA_SEGUNDOS = 3


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
        self._browser = self._playwright.chromium.launch(headless=self.headless, slow_mo=300)
        self._context = self._browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 768},
        )
        self.page = self._context.new_page()
        self.page.set_default_timeout(20000)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._context:
            self._context.close()
        if self._browser:
            self._browser.close()
        if self._playwright:
            self._playwright.stop()

    # ==========================================================
    # ENTRADA NO SISTEMA
    # ==========================================================

    def login(self) -> None:
        page = self._page()
        self.logger.info("Acessando STUR: %s", self.config.stur_url)
        page.goto(self.config.stur_url, wait_until="domcontentloaded")

        page.locator(SELECTORS["login_usuario"]).fill(self.config.stur_user)
        self.esperar("usuário preenchido")

        page.locator(SELECTORS["login_senha"]).fill(self.config.stur_password)
        self.esperar("senha preenchida")

        page.locator(SELECTORS["login_botao"]).click()
        self.esperar("login enviado")

        if self._existe_page(SELECTORS["usuario_ativo_msg"], timeout=3000):
            texto = page.locator(SELECTORS["usuario_ativo_msg"]).inner_text().strip()
            if "Usuário ativo em outra sessão" in texto:
                self.logger.warning("Sessão ativa detectada. Clicando em desbloquear usuário ativo...")
                page.locator(SELECTORS["usuario_ativo_desbloquear"]).click()
                self.esperar("usuário desbloqueado")

        self.logger.info("Login realizado.")

    def acessar_tela_vendas(self) -> None:
        page = self._page()
        self.logger.info("Acessando Operacional -> Vendas")

        try:
            page.locator(SELECTORS["menu_operacional"]).hover()
            self.esperar("hover no menu Operacional")
            page.locator(SELECTORS["menu_vendas"]).click()
        except Exception:
            self.logger.warning("Hover/click no menu falhou. Navegando para ListaVendas.aspx via JavaScript.")
            page.evaluate("Redirecionar('ListaVendas.aspx?ore')")

        self.esperar("tela de Vendas carregando")
        self.aguardar_campo_busca()
        self.logger.info("Tela de Vendas pronta.")

    # ==========================================================
    # FLUXOS DE BUSCA - PASSO A PASSO
    # ==========================================================

    def buscar_vcn_por_venda(self, transacao: Transacao) -> list[CandidatoVenda]:
        self.logger.info("Estratégia VCN: buscar direto pela coluna Venda. Termo=%s", transacao.termo_busca)
        self.limpar_filtros_com_calma()
        self.clicar_coluna("Venda")
        self.preencher_search(transacao.termo_busca)
        self.clicar_botao_pesquisar()
        self.esperar("consulta por Venda")
        return self.coletar_resultados_da_tabela(origem_busca="Venda")

    def buscar_latam_por_localizador(self, transacao: Transacao) -> list[CandidatoVenda]:
        self.logger.info("Estratégia LATAM: buscar Localizador=%s", transacao.termo_busca)
        self.limpar_filtros_com_calma()
        self.clicar_coluna("Localizador")
        self.preencher_search(transacao.termo_busca)
        self.clicar_botao_pesquisar()
        self.esperar("consulta por Localizador")
        return self.coletar_resultados_da_tabela(origem_busca="Localizador")

    def buscar_generico_por_datas(self, transacao: Transacao) -> list[CandidatoVenda]:
        """
        Estratégia otimizada para hotéis/fornecedores genéricos.

        Regra importante:
        - Primeiro pesquisa SOMENTE pelo fornecedor/fornecedor serviço.
        - Se essa primeira busca não trouxer nenhuma linha, NÃO aplica data.
          Já considera que não há base para refinar e segue para o próximo item.
        - Só marca "Manter Pesquisa" e aplica Data de Emissão/Início/Término
          quando a busca inicial trouxe candidatos.

        Isso evita gastar tempo em casos como AdaptaOrg, Uber, lanchonete etc.,
        que não existem no STUR como venda/importação.
        """
        buscas_iniciais: list[tuple[str, str]] = [
            ("Fornecedor", transacao.termo_busca),
            ("Fornecedor Serviço", transacao.termo_busca),
        ]

        if transacao.tipo_busca == "GENERICO_COM_LOCALIZADOR" and transacao.localizador_extraido:
            buscas_iniciais.append(("Localizador", transacao.localizador_extraido))

        colunas_data = ["Data de Emissão", "Data de Início", "Data de Término"]
        todos: list[CandidatoVenda] = []

        for coluna_busca, termo_busca in buscas_iniciais:
            self.logger.info(
                "Estratégia %s — busca inicial: %s='%s'",
                transacao.tipo_busca,
                coluna_busca,
                termo_busca,
            )

            self.limpar_filtros_com_calma()

            self.clicar_coluna(coluna_busca)
            self.preencher_search(termo_busca)
            self.clicar_botao_pesquisar()
            self.esperar(f"consulta inicial por {coluna_busca}")

            candidatos_iniciais = self.coletar_resultados_da_tabela(
                origem_busca=f"{coluna_busca} sem data"
            )

            if not candidatos_iniciais:
                self.logger.info(
                    "Nenhum resultado inicial para %s='%s'. Não vou aplicar filtro de data para este campo.",
                    coluna_busca,
                    termo_busca,
                )
                continue

            self.logger.info(
                "Busca inicial em %s encontrou %s candidato(s). Agora vou refinar por data.",
                coluna_busca,
                len(candidatos_iniciais),
            )

            for coluna_data in colunas_data:
                self.logger.info(
                    "Estratégia GENERICO — refinando: %s='%s' + %s='%s'",
                    coluna_busca,
                    termo_busca,
                    coluna_data,
                    transacao.data_stur,
                )

                # Recomeça a combinação do zero para não empilhar Data de Emissão + Início + Término.
                self.limpar_filtros_com_calma()

                self.clicar_coluna(coluna_busca)
                self.preencher_search(termo_busca)
                self.clicar_botao_pesquisar()
                self.esperar(f"consulta por {coluna_busca} antes de manter pesquisa")

                # Confirma de novo se ainda há resultado antes de manter pesquisa/data.
                candidatos_base = self.coletar_resultados_da_tabela(
                    origem_busca=f"{coluna_busca} base antes de {coluna_data}"
                )
                if not candidatos_base:
                    self.logger.info(
                        "A busca base deixou de retornar resultados para %s. Pulando refinamento por %s.",
                        coluna_busca,
                        coluna_data,
                    )
                    continue

                self.clicar_manter_pesquisa()
                self.clicar_coluna(coluna_data)
                self.preencher_search(transacao.data_stur)
                self.clicar_botao_pesquisar()
                self.esperar(f"consulta complementar por {coluna_data}")

                candidatos = self.coletar_resultados_da_tabela(
                    origem_busca=f"{coluna_busca} + {coluna_data}"
                )

                if candidatos:
                    self.logger.info("Foram encontrados %s candidato(s) nessa estratégia.", len(candidatos))
                    todos.extend(candidatos)
                else:
                    self.logger.info("Nenhum candidato nessa estratégia refinada.")

        self.limpar_filtros_com_calma()
        return todos

    # ==========================================================
    # AÇÕES PEQUENAS E AUDITÁVEIS
    # ==========================================================

    def clicar_coluna(self, nome_coluna: str) -> None:
        frame = self._frame()
        self.logger.info("Clicando na coluna: %s", nome_coluna)

        header = frame.locator("#c0_PH1_GridView1 th").filter(has_text=nome_coluna).first
        header.wait_for(state="visible", timeout=20000)

        link = header.locator("a").first
        if link.count() > 0:
            link.click()
        else:
            header.click()

        self.esperar(f"coluna {nome_coluna} selecionada")

    def preencher_search(self, valor: str) -> None:
        frame = self._frame()
        self.logger.info("Preenchendo search com: %s", valor)

        campo = frame.locator(SELECTORS["campo_busca"]).first
        campo.wait_for(state="visible", timeout=20000)
        campo.click()
        campo.fill("")
        self.esperar("campo search limpo")
        campo.fill(str(valor))
        self.esperar("campo search preenchido")

    def clicar_botao_pesquisar(self) -> None:
        frame = self._frame()
        self.logger.info("Clicando no botão de pesquisar")

        botao = frame.locator(SELECTORS["botao_buscar"]).first
        if botao.count() > 0:
            botao.click()
        else:
            frame.locator(SELECTORS["campo_busca"]).first.press("Enter")

        self.esperar("pesquisa enviada")

    def clicar_manter_pesquisa(self) -> None:
        frame = self._frame()
        self.logger.info("Marcando Manter Pesquisa")

        checkbox = frame.locator(SELECTORS["manter_pesquisa"]).first
        if checkbox.count() == 0:
            self.logger.warning("Checkbox Manter Pesquisa não encontrado.")
            return

        try:
            if not checkbox.is_checked():
                checkbox.check(force=True)
        except Exception:
            checkbox.click(force=True)

        self.esperar("manter pesquisa marcado")

    def desmarcar_manter_pesquisa(self) -> None:
        frame = self._frame()
        checkbox = frame.locator(SELECTORS["manter_pesquisa"]).first
        if checkbox.count() == 0:
            return

        try:
            if checkbox.is_checked():
                checkbox.uncheck(force=True)
                self.esperar("manter pesquisa desmarcado")
        except Exception:
            try:
                checkbox.click(force=True)
                self.esperar("manter pesquisa desmarcado via click")
            except Exception:
                self.logger.warning("Não foi possível desmarcar Manter Pesquisa.")

    def limpar_filtros_com_calma(self) -> None:
        self.logger.info("Limpando filtros para começar próximo passo/item")
        frame = self._frame()

        self.desmarcar_manter_pesquisa()

        limpar = frame.locator(SELECTORS["limpar_filtros"]).first
        if limpar.count() > 0:
            try:
                limpar.click(force=True)
                self.esperar("limpar filtros clicado")
            except Exception as exc:
                self.logger.warning("Falha ao clicar em limpar filtros: %s", exc)
        else:
            self.logger.warning("Botão Limpar/Cancelar filtro não encontrado. Limpando campo search manualmente.")
            try:
                campo = frame.locator(SELECTORS["campo_busca"]).first
                if campo.count() > 0:
                    campo.fill("")
                    self.esperar("campo search limpo manualmente")
            except Exception:
                pass

        self.aguardar_campo_busca()

    def coletar_resultados_da_tabela(self, origem_busca: str) -> list[CandidatoVenda]:
        frame = self._frame()
        self.logger.info("Validando resultados da tabela. Origem=%s", origem_busca)

        if frame.locator("text=Nenhum registro encontrado").count() > 0:
            self.logger.info("Mensagem 'Nenhum registro encontrado' detectada.")
            return []

        linhas = frame.locator(SELECTORS["linhas_grid"])
        total_linhas = linhas.count()
        if total_linhas <= 1:
            self.logger.info("Tabela sem linhas de resultado.")
            return []

        headers = self._obter_headers_grid()
        self.logger.info("Colunas detectadas: %s", headers)

        candidatos: list[CandidatoVenda] = []

        for i in range(1, total_linhas):  # pula header
            linha = linhas.nth(i)
            celulas = linha.locator("td")
            qtd_celulas = celulas.count()

            if qtd_celulas == 0:
                continue

            valores = [celulas.nth(j).inner_text().strip() for j in range(qtd_celulas)]
            dados = self._mapear_linha_por_headers(headers, valores)

            candidato = CandidatoVenda(
                indice_tabela=i,
                codigo_venda=self._valor_coluna(dados, "Venda"),
                data_emissao=self._valor_coluna(dados, "Data de Emissão"),
                data_inicio=self._valor_coluna(dados, "Data de Início"),
                data_termino=self._valor_coluna(dados, "Data de Término"),
                fornecedor=self._valor_coluna(dados, "Fornecedor"),
                fornecedor_servico=self._valor_coluna(dados, "Fornecedor Serviço"),
                localizador=self._valor_coluna(dados, "Localizador"),
                total_cliente=self._parse_valor_monetario(self._valor_coluna(dados, "Total Cliente")),
                total_fornecedor=self._parse_valor_monetario(self._valor_coluna(dados, "Total Fornecedor")),
                origem_busca=origem_busca,
                texto_linha=" | ".join(valores),
            )

            self.logger.info(
                "Candidato coletado | Venda=%s | Fornecedor=%s | Forn.Serviço=%s | Emissão=%s | Início=%s | Término=%s | TotalCliente=%s | TotalFornecedor=%s",
                candidato.codigo_venda,
                candidato.fornecedor,
                candidato.fornecedor_servico,
                candidato.data_emissao,
                candidato.data_inicio,
                candidato.data_termino,
                candidato.total_cliente,
                candidato.total_fornecedor,
            )

            candidatos.append(candidato)

        self.logger.info("Total de candidatos coletados: %s", len(candidatos))
        return candidatos

    def seguir_fluxo_venda_ok(self, candidato: CandidatoVenda, data_vencimento: str | None) -> None:
        """
        Fluxo executado somente quando a venda foi encontrada com valor confiável.

        Passos:
        - abrir o ícone Editar da linha encontrada na listagem;
        - editar o primeiro pagamento do fornecedor;
        - selecionar Cartão de Crédito Agência;
        - selecionar Fabio Antununcio - CARTÃO DIGITAL;
        - preencher Data de Vencimento da aba Capa;
        - clicar OK;
        - clicar Gravar;
        - clicar Voltar para retornar à listagem.
        """
        if not data_vencimento:
            raise ValueError("Data de vencimento da aba Capa não encontrada. Não vou alterar pagamento da venda.")

        self.logger.info(
            "Venda validada. Iniciando alteração de pagamento | Venda=%s | Linha tabela=%s | Vencimento=%s",
            candidato.codigo_venda,
            candidato.indice_tabela,
            data_vencimento,
        )

        self.abrir_edicao_venda(candidato)
        self.editar_primeiro_pagamento_fornecedor()
        self.preencher_pagamento_cartao_agencia(data_vencimento)
        self.gravar_venda_e_voltar()

    def abrir_edicao_venda(self, candidato: CandidatoVenda) -> None:
        frame = self._frame()
        self.logger.info("Abrindo edição da venda na linha da tabela: %s", candidato.indice_tabela)

        linhas = frame.locator(SELECTORS["linhas_grid"])
        linha = linhas.nth(candidato.indice_tabela)
        botao_editar = linha.locator("input[id*='ImgEditar']").first

        botao_editar.wait_for(state="visible", timeout=20000)
        botao_editar.click(force=True)

        self.esperar("abrir tela de edição da venda")
        frame.locator(SELECTORS["botao_gravar_venda"]).first.wait_for(state="visible", timeout=30000)
        self.logger.info("Tela de edição da venda carregada.")

    def editar_primeiro_pagamento_fornecedor(self) -> None:
        frame = self._frame()
        self.logger.info("Abrindo edição do primeiro pagamento do fornecedor.")

        grid_pagamentos = frame.locator(SELECTORS["grid_pagamentos_fornecedor"]).first
        grid_pagamentos.wait_for(state="visible", timeout=30000)

        botao_editar_pagamento = frame.locator(SELECTORS["editar_pagamento_fornecedor"]).first
        botao_editar_pagamento.wait_for(state="visible", timeout=20000)
        botao_editar_pagamento.click(force=True)

        self.esperar("abrir modal de pagamento do fornecedor")
        frame.locator(SELECTORS["modal_forma_pagamento"]).first.wait_for(state="visible", timeout=20000)
        self.logger.info("Modal de pagamento do fornecedor aberto.")

    def preencher_pagamento_cartao_agencia(self, data_vencimento: str) -> None:
        frame = self._frame()

        self.logger.info("Selecionando forma de pagamento: Cartão de Crédito Agência.")
        radio_cartao = frame.locator(SELECTORS["radio_cartao_credito_agencia"]).first
        radio_cartao.wait_for(state="visible", timeout=20000)
        radio_cartao.click(force=True)
        self.esperar("forma Cartão de Crédito Agência selecionada")

        self.logger.info("Selecionando titular: Fabio Antununcio - CARTÃO DIGITAL.")
        select_titular = frame.locator(SELECTORS["select_titular_cartao_agencia"]).first
        select_titular.wait_for(state="visible", timeout=20000)
        select_titular.select_option("29")
        self.esperar("titular do cartão selecionado")

        # Em geral o STUR sincroniza o número do cartão ao selecionar o titular.
        # Mantemos este fallback para garantir que o número correspondente também esteja selecionado.
        try:
            select_numero_cartao = frame.locator(SELECTORS["select_numero_cartao_agencia"]).first
            if select_numero_cartao.count() > 0:
                select_numero_cartao.select_option("29")
                self.esperar("número do cartão sincronizado")
        except Exception as exc:
            self.logger.warning("Não consegui sincronizar select do número do cartão. Seguindo. Detalhe: %s", exc)

        self.logger.info("Preenchendo vencimento do cartão: %s", data_vencimento)
        campo_vencimento = frame.locator(SELECTORS["data_vencimento_cartao_agencia"]).first
        campo_vencimento.wait_for(state="visible", timeout=20000)
        campo_vencimento.click()
        campo_vencimento.fill("")
        campo_vencimento.fill(data_vencimento)
        self.esperar("data de vencimento preenchida")

        self.logger.info("Confirmando modal de pagamento no botão OK.")
        frame.locator(SELECTORS["botao_ok_pagamento"]).first.click(force=True)
        self.esperar("OK do pagamento clicado")

    def gravar_venda_e_voltar(self) -> None:
        frame = self._frame()

        self.logger.info("Gravando venda.")
        frame.locator(SELECTORS["botao_gravar_venda"]).first.wait_for(state="visible", timeout=20000)
        frame.locator(SELECTORS["botao_gravar_venda"]).first.click(force=True)
        self.esperar("venda gravada")

        self.logger.info("Voltando para a listagem de vendas.")
        frame.locator(SELECTORS["botao_voltar_venda"]).first.wait_for(state="visible", timeout=20000)
        frame.locator(SELECTORS["botao_voltar_venda"]).first.click(force=True)
        self.esperar("voltar para listagem")
        self.aguardar_campo_busca()
        self.logger.info("Retorno para a listagem concluído.")

    # ==========================================================
    # SUPORTE
    # ==========================================================

    def aguardar_campo_busca(self) -> None:
        self._frame().locator(SELECTORS["campo_busca"]).first.wait_for(state="visible", timeout=20000)

    def esperar(self, motivo: str = "") -> None:
        if motivo:
            self.logger.info("Aguardando %ss — %s", ESPERA_SEGUNDOS, motivo)
        self._page().wait_for_timeout(ESPERA_SEGUNDOS * 1000)

    def salvar_screenshot_erro(self, codigo: str) -> Path | None:
        if not self.config.salvar_screenshot_erro:
            return None
        screenshots_dir = self.config.logs_dir / "screenshots"
        screenshots_dir.mkdir(exist_ok=True)
        arquivo = screenshots_dir / f"erro_{codigo}_{datetime.now():%Y%m%d_%H%M%S}.png"
        self._page().screenshot(path=str(arquivo), full_page=True)
        return arquivo

    def _obter_headers_grid(self) -> list[str]:
        frame = self._frame()
        headers_locator = frame.locator("#c0_PH1_GridView1 th")
        headers: list[str] = []

        for i in range(headers_locator.count()):
            th = headers_locator.nth(i)
            texto = " ".join(th.inner_text().split()).strip()
            colspan = th.get_attribute("colspan")
            qtd_colunas = int(colspan) if colspan and colspan.isdigit() else 1

            if not texto or texto == "\xa0":
                texto = f"__COLUNA_VAZIA_{i}"

            for indice in range(qtd_colunas):
                if qtd_colunas > 1:
                    headers.append(f"{texto}_{indice + 1}")
                else:
                    headers.append(texto)

        return headers

    def _mapear_linha_por_headers(self, headers: list[str], valores: list[str]) -> dict[str, str]:
        if len(headers) != len(valores):
            self.logger.warning(
                "Quantidade de headers (%s) diferente de células (%s). Headers=%s | Valores=%s",
                len(headers), len(valores), headers, valores,
            )

        dados: dict[str, str] = {}
        for index, valor in enumerate(valores):
            if index < len(headers):
                dados[headers[index]] = valor
            else:
                dados[f"__EXTRA_{index}"] = valor

        return dados

    def _valor_coluna(self, dados: dict[str, str], nome_coluna: str) -> str | None:
        alvo = self._normalizar(nome_coluna)
        for coluna, valor in dados.items():
            if self._normalizar(coluna) == alvo:
                return valor.strip() if valor is not None else None
        return None

    def _parse_valor_monetario(self, valor: str | None) -> Decimal | None:
        if valor is None:
            return None

        texto_original = str(valor).strip()
        if not texto_original or texto_original.lower() == "nan":
            return None

        # Evita aceitar CNPJ/CPF/códigos grandes como se fossem valor.
        apenas_digitos = re.sub(r"\D", "", texto_original)
        if "," not in texto_original and "." not in texto_original and len(apenas_digitos) > 6:
            return None

        texto = texto_original.replace("R$", "").replace(" ", "")
        if "," in texto and "." in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")

        texto = re.sub(r"[^0-9.-]", "", texto)
        if not texto:
            return None

        try:
            return Decimal(texto)
        except InvalidOperation:
            return None

    def _normalizar(self, texto: str) -> str:
        return (
            str(texto).strip().lower()
            .replace("ç", "c")
            .replace("ã", "a").replace("á", "a").replace("à", "a").replace("â", "a")
            .replace("é", "e").replace("ê", "e")
            .replace("í", "i")
            .replace("ó", "o").replace("ô", "o").replace("õ", "o")
            .replace("ú", "u")
        )

    def _existe_page(self, selector: str, timeout: int = 1000) -> bool:
        try:
            self._page().locator(selector).first.wait_for(state="visible", timeout=timeout)
            return True
        except PlaywrightTimeoutError:
            return False

    def _frame(self) -> FrameLocator:
        return self._page().frame_locator(SELECTORS["iframe_stur"])

    def _page(self) -> Page:
        if self.page is None:
            raise RuntimeError("Browser não inicializado.")
        return self.page
