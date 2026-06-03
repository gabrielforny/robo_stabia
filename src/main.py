
import argparse
from pathlib import Path

from config import load_config
from excel_service import ExcelService
from logger_config import setup_logger
from models import ResultadoProcessamento, Transacao
from stur_automation import SturAutomation
from stur_financeiro_automation import SturFinanceiroAutomation


EXTENSOES_SUPORTADAS = {".xlsx", ".xls", ".csv"}


def processar_transacao_financeiro(
    financeiro: SturFinanceiroAutomation,
    excel_service: ExcelService,
    df,
    transacao: Transacao,
    logger,
) -> bool:
    logger.info("============================================================")
    logger.info("Arquivo origem: %s", transacao.origem_arquivo)
    logger.info("Tipo layout: %s", transacao.tipo_layout)
    logger.info("Linha arquivo %s", transacao.linha_excel)
    logger.info("Descrição/Transação: %s", transacao.estabelecimento)
    logger.info("Dt.Conferência usada: %s", transacao.data_stur)
    logger.info("Valor arquivo: %s", transacao.valor_excel)

    if not transacao.estabelecimento:
        excel_service.escrever_resultado(df, transacao, "LANÇAMENTO MANUAL | Sem descrição/transação")
        return False

    if transacao.valor_excel is None:
        excel_service.escrever_resultado(df, transacao, "LANÇAMENTO MANUAL | Sem valor para comparação")
        return False

    resultado = financeiro.buscar_conferencia_por_descricao_e_data(
        descricao=transacao.estabelecimento,
        data_conferencia=transacao.data_stur,
        valor_excel=transacao.valor_excel,
        clicar_editar_quando_bater=True,
    )

    excel_service.escrever_resultado(df, transacao, resultado.mensagem)
    return resultado.encontrada and resultado.valor_bateu


def processar_arquivo_aberto(
    arquivo: Path,
    excel_service: ExcelService,
    financeiro: SturFinanceiroAutomation,
    logger,
) -> ResultadoProcessamento:
    logger.info("Iniciando processamento do arquivo: %s", arquivo)

    df, aba = excel_service.carregar_transacoes(arquivo)
    logger.info("Aba/tipo carregado: %s | Colunas: %s", aba, list(df.columns))

    transacoes = excel_service.montar_transacoes(df, origem_arquivo=arquivo.name)
    logger.info("Total de linhas a processar no arquivo %s: %d", arquivo.name, len(transacoes))

    total_sucesso = 0
    total_erro = 0

    for transacao in transacoes:
        try:
            sucesso = processar_transacao_financeiro(
                financeiro=financeiro,
                excel_service=excel_service,
                df=df,
                transacao=transacao,
                logger=logger,
            )

            if sucesso:
                total_sucesso += 1
                arquivo_parcial = excel_service.salvar_saida(df, arquivo)
                logger.info("Backup parcial salvo em: %s", arquivo_parcial)

                # Nesta etapa ainda paramos após abrir o primeiro Editar da conferência correta,
                # porque as próximas navegações serão definidas depois.
                logger.info("Match encontrado e tela de edição aberta. Encerrando esta etapa de teste.")
                return ResultadoProcessamento(
                    arquivo_saida=arquivo_parcial,
                    total_linhas=len(transacoes),
                    total_sucesso=total_sucesso,
                    total_erro=total_erro,
                )

            total_erro += 1
            arquivo_parcial = excel_service.salvar_saida(df, arquivo)
            logger.info("Backup parcial salvo em: %s", arquivo_parcial)
            financeiro.limpar_filtros_com_calma()

        except Exception as exc:
            total_erro += 1
            mensagem = f"ERRO inesperado | {type(exc).__name__}: {exc}"
            excel_service.escrever_resultado(df, transacao, mensagem)

            try:
                arquivo_parcial = excel_service.salvar_saida(df, arquivo)
                logger.info("Backup parcial salvo após erro em: %s", arquivo_parcial)
            except Exception:
                logger.warning("Não foi possível salvar backup parcial após erro.")

            logger.exception("Erro inesperado na linha %s do arquivo %s", transacao.linha_excel, arquivo.name)

            try:
                financeiro.limpar_filtros_com_calma()
            except Exception:
                logger.warning("Não foi possível limpar filtros após erro inesperado.")

            continue

    arquivo_saida = excel_service.salvar_saida(df, arquivo)
    logger.info("Arquivo salvo: %s", arquivo_saida)

    return ResultadoProcessamento(
        arquivo_saida=arquivo_saida,
        total_linhas=len(transacoes),
        total_sucesso=total_sucesso,
        total_erro=total_erro,
    )


def processar_arquivos(arquivos: list[Path], headless: bool) -> list[ResultadoProcessamento]:
    config = load_config()
    logger = setup_logger(config.logs_dir)
    excel_service = ExcelService(config)

    arquivos = [arquivo for arquivo in arquivos if arquivo.suffix.lower() in EXTENSOES_SUPORTADAS]

    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo Excel/CSV válido encontrado para processamento.")

    logger.info("Iniciando processamento - NOVO FLUXO FINANCEIRO / múltiplos arquivos")
    logger.info("Arquivos recebidos: %s", [str(arquivo) for arquivo in arquivos])

    resultados: list[ResultadoProcessamento] = []

    with SturAutomation(config=config, logger=logger, headless=headless) as stur:
        stur.login()

        financeiro = SturFinanceiroAutomation(
            page=stur._page(),
            logger=logger,
            espera_padrao_segundos=3,
        )
        financeiro.acessar_tela_conferencias_baixas()

        for arquivo in arquivos:
            resultado = processar_arquivo_aberto(
                arquivo=arquivo,
                excel_service=excel_service,
                financeiro=financeiro,
                logger=logger,
            )
            resultados.append(resultado)

            # Se encontrou match e abriu edição, paramos tudo por enquanto.
            if resultado.total_sucesso > 0:
                break

            try:
                financeiro.limpar_filtros_com_calma()
            except Exception:
                logger.warning("Não foi possível limpar filtros entre arquivos.")

    return resultados


def listar_arquivos_da_pasta(pasta: Path) -> list[Path]:
    if not pasta.exists():
        return []

    arquivos = [arquivo for arquivo in pasta.iterdir() if arquivo.is_file() and arquivo.suffix.lower() in EXTENSOES_SUPORTADAS]
    return sorted(arquivos, key=lambda arquivo: arquivo.stat().st_mtime)


def resolver_arquivos(args) -> list[Path]:
    if args.arquivo:
        return [Path(caminho) for caminho in args.arquivo]

    if args.pasta:
        return listar_arquivos_da_pasta(Path(args.pasta))

    # Padrão mais seguro para o robô: processa tudo que estiver na pasta input do projeto.
    pasta_input = Path(__file__).resolve().parent.parent / "input"
    arquivos_input = listar_arquivos_da_pasta(pasta_input)
    if arquivos_input:
        return arquivos_input

    # Fallback para testes locais: Downloads.
    return listar_arquivos_da_pasta(Path.home() / "Downloads")


def main() -> None:
    parser = argparse.ArgumentParser(description="Robô STUR — Financeiro / Conferências e Baixas")
    parser.add_argument(
        "--arquivo",
        action="append",
        required=False,
        help="Caminho do arquivo Excel/CSV. Pode informar mais de uma vez: --arquivo a.xlsx --arquivo b.csv",
    )
    parser.add_argument("--pasta", required=False, help="Pasta contendo um ou mais arquivos Excel/CSV para processar")
    parser.add_argument("--headless", action="store_true", help="Executa navegador oculto")
    args = parser.parse_args()

    arquivos = resolver_arquivos(args)
    resultados = processar_arquivos(arquivos, headless=args.headless)

    print()
    print("Processamento finalizado.")
    for resultado in resultados:
        print("-" * 60)
        print(f"Arquivo saída : {resultado.arquivo_saida}")
        print(f"Total         : {resultado.total_linhas}")
        print(f"Sucesso       : {resultado.total_sucesso}")
        print(f"Erro          : {resultado.total_erro}")


if __name__ == "__main__":
    main()
