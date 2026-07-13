import re
from decimal import Decimal
from logging import Logger

from playwright.sync_api import PlaywrightTimeoutError

from config import AppConfig
from models import CandidatoVenda, TransacaoHotel
from stur_automation import SELECTORS, SturAutomation


SELECTORS_HOTEL = {
    # Habilitar coluna Cód. Integração
    "chk_cod_integracao": "#c0_PH1_ChkCodIntegracao",

    # Botão + para novo recebimento (par do ImgNovoPag para pagamentos)
    "botao_novo_recebimento": "#c0_PH1_UFRPV1_ImgNovoRec",

    # Engrenagem por linha da grid — abre menu de ações da venda
    "engrenagem_linha": "input[id*='ImgEngrenagem']",

    # Modal Copiar Venda
    "select_produto_copia": "#c0_PH1_UsrSelecaoProduto_Dbl",
    "radio_adicionar_item": "input[id*='RblTipoCopia'][value='1']",
    "campo_venda_ref": "#c0_PH1_EdtReservaOrigem",
    "botao_ok_copia": "#c0_PH1_BtnOkCopia",

    # Tela Vendas de Hotelaria - Inclusão
    "campo_cod_integracao_inclusao": "#c0_PH1_UsrCabecVenda1_EDCH",
    "botao_gravar_inclusao":         "#c0_PH1_UsrRodapeEdicao1_BtnGravar",
    "botao_gravar_confirmacao":      "#c0_PH1_UsrRodapeEdicao1_BtnFuncao",
    "botao_ok_vc":                   "#c0_PH1_BtnOkVC",
    "label_erro_gravacao":           "#c0_PH1_Label5",
    "grid_confirmacao":              "#c0_PH1_GridView1",
    "botao_confirmar":               "#c0_PH1_BtnConfirmar",
}

# Valor do option "EXTRA HOTELARIA" no select de produto
VALUE_EXTRA_HOTELARIA = "54"


class SturHoteisAutomation(SturAutomation):
    """
    Fluxo de Hotelaria sobre SturAutomation.

    Reutiliza login, browser lifecycle, grid reading e helpers de SturAutomation.
    Adiciona busca por Cód. Integração, leitura/escrita de
    FORMAS DE RECEBIMENTO/PAGAMENTO para hotéis e o sub-fluxo Extra Hotelaria.
    """

    # ==========================================================
    # BUSCA — COD. INTEGRAÇÃO
    # ==========================================================

    def habilitar_coluna_cod_integracao(self) -> None:
        """Garante que a coluna 'Cód. Integração' está visível na grid de Vendas."""
        frame = self._frame()
        self.logger.info("Verificando coluna Cód. Integração na tela de Vendas")

        # Verifica se já está visível (style sem display:none)
        col = frame.locator("#c0_PH1_GridView1 th").filter(has_text="Cód. Integração")
        if col.count() > 0:
            style = col.first.get_attribute("style") or ""
            if "display:none" not in style and "display: none" not in style:
                self.logger.info("Coluna Cód. Integração já visível")
                return

        self.logger.info("Habilitando coluna Cód. Integração via ícone olho")
        img_eye = frame.locator("#c0_PH1_UsrCabecLista1_ImgCustom")
        img_eye.wait_for(state="visible", timeout=15000)
        img_eye.click()
        self.esperar("painel de colunas aberto")

        chk = frame.locator(SELECTORS_HOTEL["chk_cod_integracao"])
        chk.wait_for(state="visible", timeout=10000)
        if not chk.is_checked():
            chk.check(force=True)
            self.esperar("Cód. Integração marcado")

        img_eye.click()
        self.esperar("painel de colunas fechado")
        self.logger.info("Coluna Cód. Integração habilitada com sucesso")

    def buscar_por_cod_integracao(self, observacao: str) -> list[CandidatoVenda]:
        """Busca na tela de Vendas pelo valor de OBSERVAÇÃO (Cód. Integração)."""
        self.logger.info("Buscando por Cód. Integração: %s", observacao)
        self.limpar_filtros_com_calma()
        self.clicar_coluna("Cód. Integração")
        self.preencher_search(observacao)
        self.clicar_botao_pesquisar()
        self.esperar("consulta por Cód. Integração")
        return self.coletar_resultados_da_tabela(origem_busca="Cód. Integração")

    def encontrar_linha_hotelaria(self, candidatos: list[CandidatoVenda]) -> list[CandidatoVenda]:
        """Filtra candidatos onde o texto da linha contém 'HOTELARIA' (coluna Produto)."""
        return [c for c in candidatos if "hotelaria" in (c.texto_linha or "").lower()]

    def refinar_por_cliente(self, observacao: str, cliente: str) -> list[CandidatoVenda]:
        """Manter pesquisa no Cód. Integração + refina pela coluna Cliente."""
        self.logger.info("Refinando resultado por Cliente: %s", cliente)
        self.limpar_filtros_com_calma()
        self.clicar_coluna("Cód. Integração")
        self.preencher_search(observacao)
        self.clicar_botao_pesquisar()
        self.esperar("busca inicial antes de Manter Pesquisa")
        self.clicar_manter_pesquisa()
        self.clicar_coluna("Cliente")
        self.preencher_search(cliente)
        self.clicar_botao_pesquisar()
        self.esperar("refinamento por Cliente")
        return self.coletar_resultados_da_tabela(origem_busca="Cód. Integração + Cliente")

    def buscar_hotel(self, transacao: TransacaoHotel) -> CandidatoVenda | None:
        """
        Orquestra a busca completa de um hotel:
        1. Busca por Cód. Integração
        2. Filtra por produto HOTELARIA
        3. Se múltiplos, refina por Cliente
        """
        candidatos = self.buscar_por_cod_integracao(transacao.observacao)

        if not candidatos:
            self.logger.info(
                "Nenhum resultado para Cód. Integração=%s", transacao.observacao
            )
            return None

        hotelaria = self.encontrar_linha_hotelaria(candidatos)

        if not hotelaria:
            self.logger.info(
                "Nenhum candidato com 'HOTELARIA' no Produto. Total candidatos=%d | Cod=%s",
                len(candidatos), transacao.observacao,
            )
            return None

        if len(hotelaria) == 1:
            return hotelaria[0]

        # Múltiplos — refina por Cliente
        self.logger.info(
            "%d candidatos HOTELARIA. Refinando por Cliente: %s",
            len(hotelaria), transacao.cliente,
        )
        if not transacao.cliente:
            self.logger.warning("Cliente não informado — usando primeiro candidato HOTELARIA.")
            return hotelaria[0]

        refinados = self.refinar_por_cliente(transacao.observacao, transacao.cliente)
        hotelaria_r = self.encontrar_linha_hotelaria(refinados)
        if hotelaria_r:
            return hotelaria_r[0]

        self.logger.warning(
            "Refinamento por Cliente não produziu resultado. Usando primeiro candidato HOTELARIA."
        )
        return hotelaria[0]

    # ==========================================================
    # FORMAS DE RECEBIMENTO/PAGAMENTO — LEITURA
    # ==========================================================

    def ler_estado_formas_rec_pag(self) -> dict:
        """
        Lê o estado de FORMAS DE RECEBIMENTO/PAGAMENTO.

        Retorna:
          tem_faturado: bool   — recebimento do tipo Faturado já existe
          tem_ccrag: bool      — pagamento Cartão de Crédito Agência já existe
          valor_ccrag: Decimal | None — valor encontrado na linha CCRAG
        """
        frame = self._frame()
        resultado: dict = {"tem_faturado": False, "tem_ccrag": False, "valor_ccrag": None}

        # Tabela de recebimentos
        grd_r = frame.locator(SELECTORS["grid_recebimentos"])
        if grd_r.count() > 0:
            linhas_r = grd_r.locator("tbody tr")
            for i in range(linhas_r.count()):
                texto = linhas_r.nth(i).inner_text().lower()
                if "faturado" in texto:
                    resultado["tem_faturado"] = True
                    self.logger.info("Recebimento Faturado detectado na linha %d", i)
                    break

        # Tabela de pagamentos
        grd_p = frame.locator(SELECTORS["grid_pagamentos_fornecedor"])
        if grd_p.count() > 0:
            linhas_p = grd_p.locator("tbody tr")
            for i in range(linhas_p.count()):
                linha = linhas_p.nth(i)
                texto = linha.inner_text()
                texto_lower = texto.lower()
                if "cart" in texto_lower and ("ag" in texto_lower or "agencia" in texto_lower or "agência" in texto_lower):
                    resultado["tem_ccrag"] = True
                    resultado["valor_ccrag"] = self._extrair_primeiro_valor_monetario(texto)
                    self.logger.info(
                        "Pagamento CCRAG detectado na linha %d | valor=%s",
                        i, resultado["valor_ccrag"],
                    )
                    break

        self.logger.info(
            "Estado FORMAS: faturado=%s | ccrag=%s | valor_ccrag=%s",
            resultado["tem_faturado"], resultado["tem_ccrag"], resultado["valor_ccrag"],
        )
        return resultado

    def _extrair_primeiro_valor_monetario(self, texto: str) -> Decimal | None:
        """Extrai o primeiro valor decimal (ex: 3.022,20 ou 3022.20) de um texto."""
        match = re.search(r"\d{1,3}(?:\.\d{3})*(?:,\d{1,2})?|\d+(?:,\d{1,2})?", texto)
        if match:
            return self._parse_valor_monetario(match.group(0))
        return None

    # ==========================================================
    # FORMAS DE RECEBIMENTO/PAGAMENTO — ESCRITA (CASO 3)
    # ==========================================================

    def adicionar_recebimento_faturado(self) -> None:
        """Adiciona recebimento Faturado via botão + na tabela de recebimentos."""
        frame = self._frame()
        self.logger.info("Adicionando recebimento Faturado")

        botao_ok_rec = frame.locator(SELECTORS["botao_ok_recebimento"]).first

        # Verifica se modal já está aberto
        modal_aberto = False
        try:
            botao_ok_rec.wait_for(state="visible", timeout=2000)
            modal_aberto = True
        except PlaywrightTimeoutError:
            pass

        if not modal_aberto:
            botao_novo_rec = frame.locator(SELECTORS_HOTEL["botao_novo_recebimento"]).first
            botao_novo_rec.wait_for(state="visible", timeout=10000)
            botao_novo_rec.click(force=True)
            self.esperar("modal novo recebimento aberto")
            botao_ok_rec.wait_for(state="visible", timeout=10000)

        radio_faturado = frame.locator(SELECTORS["radio_faturado_recebimento"]).first
        radio_faturado.wait_for(state="visible", timeout=10000)
        if not radio_faturado.is_checked():
            radio_faturado.click(force=True)
            self.esperar("radio Faturado selecionado")

        botao_ok_rec.click(force=True)
        self.esperar("recebimento Faturado confirmado")
        self.logger.info("Recebimento Faturado adicionado")

    def adicionar_pagamento_ccrag(self, codigo_autorizacao: str = "") -> None:
        """Adiciona pagamento CCRAG via botão +, reutilizando preencher_pagamento_cartao_agencia."""
        frame = self._frame()
        self.logger.info("Adicionando pagamento Cartão de Crédito Agência")

        botao_ok_pag = frame.locator(SELECTORS["botao_ok_pagamento"]).first

        modal_aberto = False
        try:
            botao_ok_pag.wait_for(state="visible", timeout=2000)
            modal_aberto = True
        except PlaywrightTimeoutError:
            pass

        if not modal_aberto:
            botao_novo_pag = frame.locator(SELECTORS["novo_pagamento_fornecedor"]).first
            botao_novo_pag.wait_for(state="visible", timeout=10000)
            botao_novo_pag.click(force=True)
            self.esperar("modal novo pagamento aberto")
            botao_ok_pag.wait_for(state="visible", timeout=10000)

        # Preenche CCRAG (herdado de SturAutomation; o método já clica em OK)
        self.preencher_pagamento_cartao_agencia(codigo_autorizacao=codigo_autorizacao)

    def gravar_venda_hotel(self) -> None:
        """Grava venda de Hotelaria e volta para a listagem."""
        self.gravar_venda_e_voltar()

    def voltar_sem_gravar(self) -> None:
        """Volta para a listagem sem gravar (quando o estado já está correto)."""
        frame = self._frame()
        self.logger.info("Voltando para listagem sem alterar venda")
        botao_voltar = frame.locator(SELECTORS["botao_voltar_venda"]).first
        botao_voltar.wait_for(state="visible", timeout=15000)
        botao_voltar.click(force=True)
        self.esperar("voltou para listagem")
        self.aguardar_campo_busca()

    # ==========================================================
    # SUB-FLUXO EXTRA HOTELARIA (discrepância de valor)
    # ==========================================================

    def executar_copiar_venda_extra(
        self,
        candidato: CandidatoVenda,
        diferenca: Decimal,
        observacao: str,
        codigo_autorizacao: str = "",
    ) -> None:
        """
        Sub-fluxo para discrepância entre valor da planilha e valor do STUR:
        1. Clica na engrenagem da linha
        2. Seleciona 'Copiar Venda'
        3. Seleciona produto 'EXTRA HOTELARIA', venda de referência
        4. OK → abre 'Vendas de Hotelaria - Inclusão'
        5. Preenche Cód. Integração, edita linha, zera diárias/taxas, coloca diferença
        6. Grava → adiciona Faturado + CCRAG na nova venda
        """
        frame = self._frame()
        self.logger.info(
            "Iniciando Extra Hotelaria | Venda=%s | Diferença=%s",
            candidato.codigo_venda, diferenca,
        )

        # --- 1. Engrenagem da linha ---
        linhas = frame.locator(SELECTORS["linhas_grid"])
        linha = linhas.nth(candidato.indice_tabela)
        engrenagem = linha.locator(SELECTORS_HOTEL["engrenagem_linha"]).first
        engrenagem.wait_for(state="visible", timeout=10000)
        engrenagem.click(force=True)
        self.esperar("menu engrenagem aberto")

        # --- 2. Copiar Venda ---
        copiar_link = frame.locator("text=Copiar Venda").first
        copiar_link.wait_for(state="visible", timeout=5000)
        copiar_link.click(force=True)
        self.esperar("modal Copiar Venda aberto")

        # --- 3. Preenche modal ---
        select_produto = frame.locator(SELECTORS_HOTEL["select_produto_copia"]).first
        select_produto.wait_for(state="visible", timeout=10000)
        select_produto.select_option(VALUE_EXTRA_HOTELARIA)
        self.esperar("EXTRA HOTELARIA selecionado")

        radio_novo_item = frame.locator(SELECTORS_HOTEL["radio_adicionar_item"]).first
        if radio_novo_item.count() > 0:
            radio_novo_item.click(force=True)
            self.esperar("radio Adicionar novo ítem selecionado")

        if candidato.codigo_venda:
            campo_venda = frame.locator(SELECTORS_HOTEL["campo_venda_ref"]).first
            if campo_venda.count() > 0:
                campo_venda.triple_click()
                campo_venda.fill(candidato.codigo_venda)
                self.esperar("número da venda preenchido")

        botao_ok = frame.locator(SELECTORS_HOTEL["botao_ok_copia"]).first
        botao_ok.wait_for(state="visible", timeout=5000)
        botao_ok.click(force=True)
        self.esperar("OK Copiar Venda")

        # --- 4-6. Tela Vendas de Hotelaria - Inclusão ---
        self._preencher_inclusao_extra_hotelaria(
            observacao=observacao,
            diferenca=diferenca,
            codigo_autorizacao=codigo_autorizacao,
        )

    def _preencher_inclusao_extra_hotelaria(
        self, observacao: str, diferenca: Decimal, codigo_autorizacao: str = ""
    ) -> None:
        """
        Preenche 'Vendas de Hotelaria - Inclusão' após Copiar Venda:
        cód. integração → editar hóspede → editar apartamento → zerar diárias →
        OK → zerar taxas → digitar diferença → + → OK (aceita alert de duplicata) →
        validar erro de gravação → Faturado + CCRAG → Gravar e Ir para Confirmação →
        selecionar EXTRA HOTELARIA RESERVADA Definido → Confirmar.
        """
        frame = self._frame()
        page = self._page()
        self.logger.info(
            "Preenchendo Vendas de Hotelaria - Inclusão | obs=%s | dif=%s", observacao, diferenca
        )

        # Aguarda tela carregar
        frame.locator(SELECTORS_HOTEL["botao_gravar_inclusao"]).first.wait_for(
            state="visible", timeout=20000
        )

        # Cód. Integração
        campo_cod = frame.locator(SELECTORS_HOTEL["campo_cod_integracao_inclusao"]).first
        campo_cod.triple_click()
        campo_cod.fill(observacao)
        self.esperar("Cód. Integração preenchido")

        # Editar hóspede (GrdPax)
        botao_pax = frame.locator("input[id*='GrdPax'][id*='ImgEditar']").first
        botao_pax.wait_for(state="visible", timeout=10000)
        botao_pax.click(force=True)
        self.esperar("tela edição hóspede aberta")

        # Editar apartamento (abre modal de diárias)
        botao_apto = frame.locator("input[id*='GrdAptosCli'][id*='ImgEditar']").first
        botao_apto.wait_for(state="visible", timeout=10000)
        botao_apto.click(force=True)
        self.esperar("modal edição apartamento aberto")

        # Zerar diária cliente e fornecedor
        for selector in (
            "#c0_PH1_UTVH_UsrVlrDiariaCli_ED",
            "#c0_PH1_UTVH_UsrVlrDiariaFor_ED",
        ):
            campo = frame.locator(selector).first
            campo.wait_for(state="visible", timeout=10000)
            campo.triple_click()
            campo.fill("0,00")
        self.esperar("diárias zeradas")

        # OK do modal de apartamento
        frame.locator("#c0_PH1_UTVH_BtnOk").first.click(force=True)
        self.esperar("modal apartamento confirmado")

        # Zerar taxas
        campo_taxas = frame.locator("#c0_PH1_UTVH_UsrTaxasCli_ED").first
        campo_taxas.wait_for(state="visible", timeout=10000)
        campo_taxas.triple_click()
        campo_taxas.fill("0,00")
        self.esperar("taxas zeradas")

        # Digitar diferença no campo Extra
        valor_str = f"{diferenca:.2f}".replace(".", ",")
        campo_extra = frame.locator("#c0_PH1_UTVH_UsrTxExtraCli_ED").first
        campo_extra.triple_click()
        campo_extra.fill(valor_str)
        self.esperar("valor extra preenchido")

        # Clicar "+" (adicionar outras taxas / confirmar extra)
        frame.locator("#c0_PH1_UTVH_ImgEC").first.click(force=True)
        self.esperar("extras adicionados")

        # OK da tela de hospedagem — aceita alert de reserva duplicada se aparecer
        def _aceitar_dialog(dialog):
            self.logger.info("Alert BtnOkVC: %s — aceitando", dialog.message[:120])
            dialog.accept()

        page.once("dialog", _aceitar_dialog)
        frame.locator(SELECTORS_HOTEL["botao_ok_vc"]).first.click(force=True)
        self.esperar("BtnOkVC confirmado")

        # Verificar erro de gravação (coordenador inativo etc.)
        label_erro = frame.locator(SELECTORS_HOTEL["label_erro_gravacao"])
        try:
            label_erro.wait_for(state="visible", timeout=3000)
            texto_erro = (label_erro.text_content() or "").strip()
            if "Problemas na gravação" in texto_erro:
                raise RuntimeError(f"Erro de gravação STUR: {texto_erro}")
        except PlaywrightTimeoutError:
            pass  # sem erro — continua

        # Faturado + CCRAG na nova venda
        self.adicionar_recebimento_faturado()
        self.adicionar_pagamento_ccrag(codigo_autorizacao=codigo_autorizacao)

        # Gravar e Ir para Confirmação (fallback: Gravar)
        botao_func = frame.locator(SELECTORS_HOTEL["botao_gravar_confirmacao"]).first
        try:
            botao_func.wait_for(state="visible", timeout=5000)
            botao_func.click(force=True)
        except PlaywrightTimeoutError:
            frame.locator(SELECTORS_HOTEL["botao_gravar_inclusao"]).first.click(force=True)
        self.esperar("gravando Extra Hotelaria")

        # Tela de Confirmação — seleciona EXTRA HOTELARIA RESERVADA com formas Definidas
        frame = self._frame()
        grid_conf = frame.locator(SELECTORS_HOTEL["grid_confirmacao"])
        grid_conf.wait_for(state="visible", timeout=20000)

        checkbox_alvo = None
        for linha in grid_conf.locator("tr").all():
            texto = (linha.text_content() or "").upper()
            if "EXTRA HOTELARIA" not in texto or "RESERVADA" not in texto:
                continue
            chk = linha.locator("input[type='checkbox']").first
            if chk.count() == 0:
                continue
            onclick_val = chk.get_attribute("onclick") or ""
            if "alert(" not in onclick_val:
                checkbox_alvo = chk
                break

        if checkbox_alvo is None:
            raise RuntimeError(
                "Linha EXTRA HOTELARIA RESERVADA com formas Definidas não encontrada na tela de Confirmação"
            )

        checkbox_alvo.check(force=True)
        self.esperar("EXTRA HOTELARIA selecionado")

        frame.locator(SELECTORS_HOTEL["botao_confirmar"]).first.click(force=True)
        self.esperar("Extra Hotelaria confirmado")
        self.logger.info(
            "Extra Hotelaria concluído | obs=%s | dif=%s", observacao, diferenca
        )
