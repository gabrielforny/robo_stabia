
import argparse
from collections import defaultdict
from pathlib import Path

from config import load_config
from excel_service import ExcelService
from logger_config import setup_logger
from models import ResultadoProcessamento, Transacao
from stur_automation import SturAutomation
from stur_financeiro_automation import SturFinanceiroAutomation


EXTENSOES_SUPORTADAS = {".xlsx", ".xls", ".csv"}


def processar_latam_conferencia(
    financeiro: SturFinanceiroAutomation,
    excel_service: ExcelService,
    df,
    transacoes_latam: list[Transacao],
    logger,
) -> tuple[int, int]:
    """
    Agrupa os itens LATAM por mês/ano de fatura, busca ou cria a conferência
    correspondente, adiciona todos os localizadores e grava.
    """
    grupos: dict[str, list[Transacao]] = defaultdict(list)
    sem_fatura: list[Transacao] = []

    for t in transacoes_latam:
        if t.data_fatura:
            # data_fatura = dd/mm/aaaa → chave = mm/aaaa
            partes = t.data_fatura.split("/")
            chave = f"{partes[1]}/{partes[2]}" if len(partes) == 3 else t.data_fatura
            grupos[chave].append(t)
        else:
            sem_fatura.append(t)

    total_sucesso = 0
    total_erro = 0

    for t in sem_fatura:
        msg = "ERRO | LATAM sem data de fatura (extrato da conta) para identificar conferência"
        excel_service.escrever_resultado(df, t, msg)
        total_erro += 1

    for chave_mes, grupo in grupos.items():
        # TODO: remover sufixo quando sair de testes
        descricao_busca = f"Clara {chave_mes} - teste robo"
        descricao_criar = f"Clara {chave_mes} - teste robo"
        data_fatura = grupo[0].data_fatura

        logger.info("Processando conferência LATAM: %s | %d item(ns)", descricao_busca, len(grupo))

        try:
            financeiro.buscar_ou_criar_conferencia_latam(
                descricao_busca=descricao_busca,
                descricao_criar=descricao_criar,
                data_fatura=data_fatura,
            )

            financeiro.abrir_adicionar_titulos()
            financeiro.garantir_coluna_localizador_visivel()

            for transacao in grupo:
                if not transacao.localizador_extraido:
                    msg = "ERRO | LATAM sem localizador extraído"
                    excel_service.escrever_resultado(df, transacao, msg)
                    total_erro += 1
                    continue

                encontrado, motivo = financeiro.buscar_e_selecionar_localizador(
                    localizador=transacao.localizador_extraido,
                    valor_excel=transacao.valor_excel,
                )

                if encontrado:
                    msg = f"OK | Localizador {transacao.localizador_extraido} selecionado | conferência {descricao_busca}"
                    total_sucesso += 1
                else:
                    msg = f"ERRO | {motivo}"
                    total_erro += 1

                excel_service.escrever_resultado(df, transacao, msg)

            financeiro.gravar_titulos()
            financeiro.gravar_conferencia()

        except Exception as exc:
            logger.exception("Erro no processamento da conferência LATAM %s", descricao_busca)
            for transacao in grupo:
                msg = f"ERRO inesperado na conferência | {type(exc).__name__}: {exc}"
                excel_service.escrever_resultado(df, transacao, msg)
                total_erro += 1

            try:
                financeiro.limpar_filtros_com_calma()
            except Exception:
                logger.warning("Não foi possível limpar filtros após erro LATAM.")

    return total_sucesso, total_erro


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
    logger.info("Total de linhas a processar: %d", len(transacoes))

    transacoes_latam = [t for t in transacoes if t.tipo_busca == "LATAM"]
    transacoes_outras = [t for t in transacoes if t.tipo_busca != "LATAM"]

    logger.info("LATAM: %d | Outros: %d", len(transacoes_latam), len(transacoes_outras))

    total_sucesso = 0
    total_erro = 0

    if transacoes_latam:
        sucesso_latam, erro_latam = processar_latam_conferencia(
            financeiro=financeiro,
            excel_service=excel_service,
            df=df,
            transacoes_latam=transacoes_latam,
            logger=logger,
        )
        total_sucesso += sucesso_latam
        total_erro += erro_latam

        arquivo_parcial = excel_service.salvar_saida(df, arquivo)
        logger.info("Backup parcial salvo após LATAM em: %s", arquivo_parcial)

    if transacoes_outras:
        logger.info(
            "%d transação(ões) não-LATAM ignoradas nesta versão (v2 implementará outros tipos).",
            len(transacoes_outras),
        )

    arquivo_saida = excel_service.salvar_saida(df, arquivo)
    logger.info("Arquivo salvo: %s", arquivo_saida)

    return ResultadoProcessamento(
        arquivo_saida=arquivo_saida,
        total_linhas=len(transacoes_latam),
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

    logger.info("Iniciando processamento - fluxo LATAM / Conferências")
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

            try:
                financeiro.acessar_tela_conferencias_baixas()
            except Exception:
                logger.warning("Não foi possível retornar à lista de conferências entre arquivos.")

    return resultados


def listar_arquivos_da_pasta(pasta: Path) -> list[Path]:
    if not pasta.exists():
        return []

    arquivos = [
        arquivo for arquivo in pasta.iterdir()
        if arquivo.is_file() and arquivo.suffix.lower() in EXTENSOES_SUPORTADAS
    ]
    return sorted(arquivos, key=lambda arquivo: arquivo.stat().st_mtime)


def resolver_arquivos(args) -> list[Path]:
    if args.arquivo:
        return [Path(caminho) for caminho in args.arquivo]

    if args.pasta:
        return listar_arquivos_da_pasta(Path(args.pasta))

    pasta_input = Path(__file__).resolve().parent.parent / "input"
    arquivos_input = listar_arquivos_da_pasta(pasta_input)
    if arquivos_input:
        return arquivos_input

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
