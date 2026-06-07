
import argparse
from collections import defaultdict
from pathlib import Path

from config import load_config
from excel_service import ExcelService
from logger_config import setup_logger
from models import ResultadoProcessamento, Transacao
from stur_automation import SturAutomation, VendaJaFaturadaError
from stur_financeiro_automation import SturFinanceiroAutomation


EXTENSOES_SUPORTADAS = {".xlsx", ".xls", ".csv"}


# ==========================================================
# FASE 1 — VENDAS (Operacional → Vendas)
# ==========================================================

def processar_latam_vendas(
    stur: SturAutomation,
    excel_service: ExcelService,
    df,
    transacoes_latam: list[Transacao],
    logger,
) -> tuple[int, int]:
    """
    Para cada item LATAM: busca o localizador na tela de Vendas, valida
    Total Fornecedor e executa seguir_fluxo_venda_ok.
    """

    total_sucesso = 0
    total_erro = 0

    for transacao in transacoes_latam:
        if not transacao.localizador_extraido:
            msg = "ERRO | LATAM sem localizador extraído"
            excel_service.escrever_resultado(df, transacao, msg)
            total_erro += 1
            continue

        logger.info("Buscando localizador nas Vendas: %s", transacao.localizador_extraido)

        try:
            candidatos = stur.buscar_latam_por_localizador(transacao)
        except Exception as exc:
            logger.exception("Erro ao buscar localizador %s nas Vendas", transacao.localizador_extraido)
            excel_service.escrever_resultado(df, transacao, f"ERRO | Busca Vendas: {type(exc).__name__}: {exc}")
            total_erro += 1
            continue

        if not candidatos:
            msg = f"ERRO | Localizador {transacao.localizador_extraido} não encontrado nas Vendas"
            excel_service.escrever_resultado(df, transacao, msg)
            total_erro += 1
            continue

        # Valida Total Fornecedor (ignora sinal)
        candidato_ok = None
        for c in candidatos:
            if transacao.valor_excel is not None and c.total_fornecedor is not None:
                if abs(c.total_fornecedor) == abs(transacao.valor_excel):
                    candidato_ok = c
                    break

        if not candidato_ok:
            vals = [str(c.total_fornecedor) for c in candidatos]
            msg = (
                f"ERRO | Valor não bate nas Vendas | "
                f"Excel={transacao.valor_excel} | Tabela={vals}"
            )
            excel_service.escrever_resultado(df, transacao, msg)
            total_erro += 1
            continue

        logger.info(
            "Localizador %s → Venda=%s | TotalForn=%s",
            transacao.localizador_extraido,
            candidato_ok.codigo_venda,
            candidato_ok.total_fornecedor,
        )

        try:
            stur.seguir_fluxo_venda_ok(candidato_ok, codigo_autorizacao=transacao.codigo_autorizacao)
            msg = f"OK Vendas | Venda {candidato_ok.codigo_venda} | Loc {transacao.localizador_extraido}"
            excel_service.escrever_resultado(df, transacao, msg)
            total_sucesso += 1
        except VendaJaFaturadaError:
            logger.warning("Venda %s já faturada — marcando e seguindo.", candidato_ok.codigo_venda)
            msg = f"JÁ FATURADO | Venda {candidato_ok.codigo_venda} | Loc {transacao.localizador_extraido}"
            excel_service.escrever_resultado(df, transacao, msg)
            total_erro += 1
        except Exception as exc:
            logger.exception("Erro ao processar venda %s", candidato_ok.codigo_venda)
            msg = f"ERRO | Venda {candidato_ok.codigo_venda}: {type(exc).__name__}: {exc}"
            excel_service.escrever_resultado(df, transacao, msg)
            total_erro += 1

    return total_sucesso, total_erro


# ==========================================================
# FASE 2 — CONFERÊNCIAS (Financeiro → Conferências e Baixas)
# ==========================================================

def processar_latam_conferencia(
    financeiro: SturFinanceiroAutomation,
    excel_service: ExcelService,
    df,
    transacoes_latam: list[Transacao],
    logger,
) -> None:
    """
    Agrupa os itens LATAM por mês/ano de fatura, busca ou cria a conferência,
    adiciona todos os localizadores e grava.
    Usa acrescentar_resultado para appendar ao resultado já escrito pela Fase 1.
    """
    grupos: dict[str, list[Transacao]] = defaultdict(list)
    sem_fatura: list[Transacao] = []

    for t in transacoes_latam:
        if t.data_fatura:
            partes = t.data_fatura.split("/")
            chave = f"{partes[1]}/{partes[2]}" if len(partes) == 3 else t.data_fatura
            grupos[chave].append(t)
        else:
            sem_fatura.append(t)

    for t in sem_fatura:
        excel_service.acrescentar_resultado(
            df, t, "ERRO Conferência | LATAM sem data de fatura para identificar conferência"
        )

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
                    excel_service.acrescentar_resultado(
                        df, transacao, "ERRO Conferência | sem localizador"
                    )
                    continue

                encontrado, motivo = financeiro.buscar_e_selecionar_localizador(
                    localizador=transacao.localizador_extraido,
                    valor_excel=transacao.valor_excel,
                )

                if encontrado:
                    excel_service.acrescentar_resultado(
                        df, transacao,
                        f"OK Conferência | {descricao_busca} | Loc {transacao.localizador_extraido}",
                    )
                else:
                    excel_service.acrescentar_resultado(
                        df, transacao, f"ERRO Conferência | {motivo}"
                    )

            financeiro.gravar_titulos()
            financeiro.gravar_conferencia()

        except Exception as exc:
            logger.exception("Erro no processamento da conferência LATAM %s", descricao_busca)
            for transacao in grupo:
                excel_service.acrescentar_resultado(
                    df, transacao,
                    f"ERRO Conferência inesperado | {type(exc).__name__}: {exc}",
                )
            try:
                financeiro.limpar_filtros_com_calma()
            except Exception:
                logger.warning("Não foi possível limpar filtros após erro LATAM.")


# ==========================================================
# ORQUESTRAÇÃO POR ARQUIVO
# ==========================================================

def processar_arquivo_aberto(
    arquivo: Path,
    excel_service: ExcelService,
    stur: SturAutomation,
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

    logger.info("LATAM: %d | Outros (ignorados nesta versão): %d", len(transacoes_latam), len(transacoes_outras))

    total_sucesso = 0
    total_erro = 0

    if transacoes_latam:
        # Fase 1: Vendas
        logger.info("=== FASE 1: Vendas ===")
        sucesso_v, erro_v = processar_latam_vendas(
            stur=stur,
            excel_service=excel_service,
            df=df,
            transacoes_latam=transacoes_latam,
            logger=logger,
        )
        total_sucesso += sucesso_v
        total_erro += erro_v

        arquivo_parcial = excel_service.salvar_saida(df, arquivo)
        logger.info("Backup parcial (pós Vendas) salvo em: %s", arquivo_parcial)

        # Fase 2: Conferências
        logger.info("=== FASE 2: Conferências ===")
        financeiro.acessar_tela_conferencias_baixas()
        processar_latam_conferencia(
            financeiro=financeiro,
            excel_service=excel_service,
            df=df,
            transacoes_latam=transacoes_latam,
            logger=logger,
        )

    arquivo_saida = excel_service.salvar_saida(df, arquivo)
    logger.info("Arquivo final salvo: %s", arquivo_saida)

    return ResultadoProcessamento(
        arquivo_saida=arquivo_saida,
        total_linhas=len(transacoes_latam),
        total_sucesso=total_sucesso,
        total_erro=total_erro,
    )


# ==========================================================
# PONTO DE ENTRADA
# ==========================================================

def processar_arquivos(arquivos: list[Path], headless: bool) -> list[ResultadoProcessamento]:
    config = load_config()
    logger = setup_logger(config.logs_dir)
    excel_service = ExcelService(config)

    arquivos = [a for a in arquivos if a.suffix.lower() in EXTENSOES_SUPORTADAS]

    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo Excel/CSV válido encontrado para processamento.")

    logger.info("Iniciando processamento — fluxo LATAM: Vendas + Conferências")
    logger.info("Arquivos recebidos: %s", [str(a) for a in arquivos])

    resultados: list[ResultadoProcessamento] = []

    with SturAutomation(config=config, logger=logger, headless=headless) as stur:
        stur.login()

        financeiro = SturFinanceiroAutomation(
            page=stur._page(),
            logger=logger,
            espera_padrao_segundos=3,
        )

        # Começa na tela de Vendas (Fase 1)
        stur.acessar_tela_vendas()
        stur.garantir_coluna_localizador_visivel()

        for arquivo in arquivos:
            resultado = processar_arquivo_aberto(
                arquivo=arquivo,
                excel_service=excel_service,
                stur=stur,
                financeiro=financeiro,
                logger=logger,
            )
            resultados.append(resultado)

            # Entre arquivos: volta para Vendas para o próximo
            if arquivo != arquivos[-1]:
                try:
                    stur.acessar_tela_vendas()
                    stur.garantir_coluna_localizador_visivel()
                except Exception:
                    logger.warning("Não foi possível retornar à tela de Vendas entre arquivos.")

    return resultados


def listar_arquivos_da_pasta(pasta: Path) -> list[Path]:
    if not pasta.exists():
        return []
    arquivos = [
        a for a in pasta.iterdir()
        if a.is_file() and a.suffix.lower() in EXTENSOES_SUPORTADAS
    ]
    return sorted(arquivos, key=lambda a: a.stat().st_mtime)


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
    parser.add_argument("--pasta", required=False, help="Pasta contendo um ou mais arquivos Excel/CSV")
    parser.add_argument("--headless", action="store_true", help="Executa navegador oculto")
    args = parser.parse_args()

    arquivos = resolver_arquivos(args)
    resultados = processar_arquivos(arquivos, headless=args.headless)

    print()
    print("Processamento finalizado.")
    for resultado in resultados:
        print("-" * 60)
        print(f"Arquivo saída : {resultado.arquivo_saida}")
        print(f"Total LATAM   : {resultado.total_linhas}")
        print(f"Sucesso Vendas: {resultado.total_sucesso}")
        print(f"Erro Vendas   : {resultado.total_erro}")


if __name__ == "__main__":
    main()
