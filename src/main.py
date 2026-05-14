import argparse
from decimal import Decimal
from pathlib import Path

from config import load_config
from excel_service import ExcelService
from logger_config import setup_logger
from models import CandidatoVenda, ResultadoProcessamento, Transacao
from stur_automation import SturAutomation


VALOR_PROXIMO_RANGE = Decimal("600.00")


def valores_batem(valor_excel: Decimal | None, valor_stur: Decimal | None, tolerancia: Decimal) -> bool:
    if valor_excel is None or valor_stur is None:
        return False
    return abs(abs(valor_excel) - abs(valor_stur)) <= tolerancia


def valores_proximos(valor_excel: Decimal | None, valor_stur: Decimal | None) -> bool:
    if valor_excel is None or valor_stur is None:
        return False
    return abs(abs(valor_excel) - abs(valor_stur)) <= VALOR_PROXIMO_RANGE


def candidato_tem_valor_exato(candidato: CandidatoVenda, transacao: Transacao, tolerancia: Decimal) -> bool:
    return (
        valores_batem(transacao.valor_excel, candidato.total_fornecedor, tolerancia)
        or valores_batem(transacao.valor_excel, candidato.total_cliente, tolerancia)
    )


def candidato_tem_valor_proximo(candidato: CandidatoVenda, transacao: Transacao) -> bool:
    return (
        valores_proximos(transacao.valor_excel, candidato.total_fornecedor)
        or valores_proximos(transacao.valor_excel, candidato.total_cliente)
    )


def escolher_candidato_exato(
    candidatos: list[CandidatoVenda],
    transacao: Transacao,
    tolerancia: Decimal,
    logger,
) -> CandidatoVenda | None:
    logger.info("Validando candidatos com valor exato. Valor Excel=%s", transacao.valor_excel)

    for candidato in candidatos:
        logger.info(
            "Analisando candidato | Venda=%s | TotalFornecedor=%s | TotalCliente=%s | Origem=%s",
            candidato.codigo_venda,
            candidato.total_fornecedor,
            candidato.total_cliente,
            candidato.origem_busca,
        )

        if candidato_tem_valor_exato(candidato, transacao, tolerancia):
            logger.info("Candidato EXATO aprovado. Venda=%s", candidato.codigo_venda)
            return candidato

    logger.info("Nenhum candidato exato encontrado.")
    return None


def montar_observacao_possiveis_vendas(candidatos: list[CandidatoVenda], transacao: Transacao) -> str:
    proximos = [c for c in candidatos if candidato_tem_valor_proximo(c, transacao)]
    base = proximos[:8] if proximos else candidatos[:8]

    partes = []
    for c in base:
        partes.append(
            f"Venda {c.codigo_venda or '-'} "
            f"({c.origem_busca}; Forn={c.total_fornecedor}; Cliente={c.total_cliente})"
        )

    if proximos:
        return "POSSÍVEL VENDA | Validar manualmente: " + " | ".join(partes)

    return "ERRO | Encontrou candidatos, mas nenhum valor bate/proxima: " + " | ".join(partes)


def processar_transacao(
    stur: SturAutomation,
    excel_service: ExcelService,
    df,
    transacao: Transacao,
    tolerancia: Decimal,
    data_vencimento_capa: str | None,
    logger,
) -> bool:
    logger.info("============================================================")
    logger.info("Linha Excel %s", transacao.linha_excel)
    logger.info("Tipo busca: %s", transacao.tipo_busca)
    logger.info("Estabelecimento: %s", transacao.estabelecimento)
    logger.info("VCN: %s | Código Venda VCN: %s", transacao.vcn, transacao.codigo_venda_vcn)
    logger.info("Localizador extraído: %s", transacao.localizador_extraido)
    logger.info("Termo: %s | Coluna: %s", transacao.termo_busca, transacao.coluna_busca)
    logger.info("Data STUR: %s | Valor Excel: %s", transacao.data_stur, transacao.valor_excel)

    if not transacao.termo_busca:
        excel_service.escrever_resultado(df, transacao, "LANÇAMENTO MANUAL | Sem termo de busca")
        return False

    if transacao.tipo_busca in {"GENERICO", "GENERICO_COM_LOCALIZADOR"} and not transacao.data_stur:
        excel_service.escrever_resultado(df, transacao, "LANÇAMENTO MANUAL | Sem data para consultar")
        return False

    candidatos: list[CandidatoVenda] = []

    if transacao.tipo_busca == "VCN":
        candidatos = stur.buscar_vcn_por_venda(transacao)

        if len(candidatos) == 1:
            candidato = candidatos[0]
            excel_service.escrever_resultado(df, transacao, f"OK | Venda={candidato.codigo_venda} | Localizada por VCN")
            stur.seguir_fluxo_venda_ok(candidato, data_vencimento_capa)
            stur.limpar_filtros_com_calma()
            return True

        if not candidatos:
            excel_service.escrever_resultado(df, transacao, "ERRO | Venda do VCN não localizada no STUR")
            stur.limpar_filtros_com_calma()
            return False

        excel_service.escrever_resultado(df, transacao, f"ERRO | VCN retornou {len(candidatos)} vendas. Validar manualmente")
        stur.limpar_filtros_com_calma()
        return False

    if transacao.tipo_busca == "LATAM":
        candidatos = stur.buscar_latam_por_localizador(transacao)
    else:
        candidatos = stur.buscar_generico_por_datas(transacao)

    if not candidatos:
        mensagem = "LANÇAMENTO MANUAL | Não localizado no STUR"
        excel_service.escrever_resultado(df, transacao, mensagem)
        logger.warning(mensagem)
        stur.limpar_filtros_com_calma()
        return False

    candidato_exato = escolher_candidato_exato(candidatos, transacao, tolerancia, logger)

    if candidato_exato:
        mensagem = (
            f"OK | Venda={candidato_exato.codigo_venda} | Origem={candidato_exato.origem_busca} | "
            f"Excel={transacao.valor_excel} | Fornecedor={candidato_exato.total_fornecedor} | Cliente={candidato_exato.total_cliente}"
        )
        excel_service.escrever_resultado(df, transacao, mensagem)
        stur.seguir_fluxo_venda_ok(candidato_exato, data_vencimento_capa)
        stur.limpar_filtros_com_calma()
        return True

    observacao = montar_observacao_possiveis_vendas(candidatos, transacao)
    excel_service.escrever_resultado(df, transacao, observacao)
    logger.warning(observacao)
    stur.limpar_filtros_com_calma()
    return False


def processar_arquivo(arquivo: Path, headless: bool) -> ResultadoProcessamento:
    config = load_config()
    logger = setup_logger(config.logs_dir)
    excel_service = ExcelService(config)

    logger.info("Iniciando processamento: %s", arquivo)

    df, aba = excel_service.carregar_transacoes(arquivo)
    logger.info("Aba carregada: %s | Colunas: %s", aba, list(df.columns))

    data_vencimento_capa = excel_service.obter_vencimento_capa(arquivo)
    logger.info("Data de vencimento encontrada na aba Capa: %s", data_vencimento_capa)

    transacoes = excel_service.montar_transacoes(df)
    logger.info("Total de linhas a processar: %d", len(transacoes))

    total_sucesso = 0
    total_erro = 0

    with SturAutomation(config=config, logger=logger, headless=headless) as stur:
        stur.login()
        stur.acessar_tela_vendas()

        for transacao in transacoes:
            try:
                sucesso = processar_transacao(
                    stur=stur,
                    excel_service=excel_service,
                    df=df,
                    transacao=transacao,
                    tolerancia=config.tolerancia_valor,
                    data_vencimento_capa=data_vencimento_capa,
                    logger=logger,
                )

                if sucesso:
                    total_sucesso += 1
                else:
                    total_erro += 1

                arquivo_parcial = excel_service.salvar_saida(df, arquivo)
                logger.info("Backup parcial salvo em: %s", arquivo_parcial)

            except Exception as exc:
                total_erro += 1
                screenshot = None
                try:
                    screenshot = stur.salvar_screenshot_erro(str(transacao.linha_excel))
                except Exception:
                    pass

                mensagem = f"ERRO inesperado | {type(exc).__name__}: {exc}"
                if screenshot:
                    mensagem += f" | Screenshot={screenshot}"

                excel_service.escrever_resultado(df, transacao, mensagem)
                try:
                    arquivo_parcial = excel_service.salvar_saida(df, arquivo)
                    logger.info("Backup parcial salvo após erro em: %s", arquivo_parcial)
                except Exception:
                    logger.warning("Não foi possível salvar backup parcial após erro.")

                logger.exception("Erro inesperado na linha %s", transacao.linha_excel)

                try:
                    stur.limpar_filtros_com_calma()
                except Exception:
                    logger.warning("Não foi possível limpar filtros após erro inesperado.")

                continue

    validacao = excel_service.validar_total_primeira_aba(arquivo, df)
    logger.info("Validação final: %s", validacao)

    arquivo_saida = excel_service.salvar_saida(df, arquivo)
    logger.info("Arquivo salvo: %s", arquivo_saida)

    return ResultadoProcessamento(
        arquivo_saida=arquivo_saida,
        total_linhas=len(transacoes),
        total_sucesso=total_sucesso,
        total_erro=total_erro,
    )


def buscar_arquivo_mais_recente(pasta: Path) -> Path:
    arquivos = list(pasta.glob("*.xlsx")) + list(pasta.glob("*.xls")) + list(pasta.glob("*.csv"))
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo Excel/CSV encontrado em: {pasta}")
    return max(arquivos, key=lambda f: f.stat().st_mtime)


def main() -> None:
    parser = argparse.ArgumentParser(description="Robô STUR — Conciliação passo a passo com regras da reunião")
    parser.add_argument("--arquivo", required=False, help="Caminho do arquivo Excel/CSV")
    parser.add_argument("--headless", action="store_true", help="Executa navegador oculto")
    args = parser.parse_args()

    arquivo = Path(args.arquivo) if args.arquivo else buscar_arquivo_mais_recente(Path.home() / "Downloads")
    resultado = processar_arquivo(arquivo, headless=args.headless)

    print()
    print("Processamento finalizado.")
    print(f"Arquivo saída : {resultado.arquivo_saida}")
    print(f"Total         : {resultado.total_linhas}")
    print(f"Sucesso       : {resultado.total_sucesso}")
    print(f"Erro          : {resultado.total_erro}")


if __name__ == "__main__":
    main()
