import argparse
from decimal import Decimal
from pathlib import Path

from config import load_config
from excel_service import ExcelService
from logger_config import setup_logger
from models import ResultadoProcessamento
from stur_automation import SturAutomation


def comparar_valores(valor_excel: Decimal | None, valor_site: Decimal | None, tolerancia: Decimal) -> tuple[bool, str]:
    if valor_excel is None:
        return False, "Valor do Excel não identificado."

    if valor_site is None:
        return False, "Total do fornecedor no site não identificado."

    diferenca = abs(valor_excel - valor_site)

    if diferenca <= tolerancia:
        return True, f"Valor OK | Excel={valor_excel} | Site={valor_site}"

    return False, f"Valor divergente | Excel={valor_excel} | Site={valor_site} | Diferença={diferenca}"


def processar_arquivo(arquivo: Path, headless: bool) -> ResultadoProcessamento:
    config = load_config()
    logger = setup_logger(config.logs_dir)
    excel_service = ExcelService(config)

    logger.info("Iniciando processamento do arquivo: %s", arquivo)

    df, aba = excel_service.carregar_transacoes(arquivo)
    logger.info("Aba/Origem carregada: %s", aba)
    logger.info("Colunas encontradas: %s", list(df.columns))

    transacoes = excel_service.montar_transacoes(df)

    logger.info("Total de transações com código após '*': %s", len(transacoes))

    total_sucesso = 0
    total_erro = 0

    with SturAutomation(config=config, logger=logger, headless=headless) as stur:
        stur.login()
        stur.acessar_vendas()

        for transacao in transacoes:
            logger.info(
                "Processando linha Excel %s | Código Companhia: %s",
                transacao.linha_excel,
                transacao.codigo_companhia,
            )

            try:
                resultado_venda = stur.buscar_venda(transacao.codigo_companhia)

                if not resultado_venda.encontrada:
                    total_erro += 1
                    mensagem = f"ERRO | {resultado_venda.mensagem}"
                    excel_service.escrever_resultado(df, transacao, mensagem)
                    logger.warning(mensagem)
                    continue

                valores_ok, mensagem_comparacao = comparar_valores(
                    valor_excel=transacao.valor_excel,
                    valor_site=resultado_venda.total_fornecedor,
                    tolerancia=config.tolerancia_valor,
                )

                if not valores_ok:
                    total_erro += 1
                    mensagem = f"ERRO | Venda {resultado_venda.codigo_venda or 'sem código'} | {mensagem_comparacao}"
                    excel_service.escrever_resultado(df, transacao, mensagem)
                    logger.warning(mensagem)
                    continue

                # Quando terminar o estudo do vídeo, complemente esse método com os campos corretos.
                stur.processar_pagamento_fornecedor(transacao)

                total_sucesso += 1
                mensagem = f"Venda {resultado_venda.codigo_venda or 'sem código'} | {mensagem_comparacao}"
                excel_service.escrever_resultado(df, transacao, mensagem)
                logger.info("SUCESSO | %s", mensagem)

            except Exception as exc:
                total_erro += 1
                screenshot = None

                try:
                    screenshot = stur.salvar_screenshot_erro(transacao.codigo_companhia)
                except Exception:
                    pass

                mensagem = f"ERRO inesperado | {type(exc).__name__}: {exc}"

                if screenshot:
                    mensagem += f" | Screenshot: {screenshot}"

                excel_service.escrever_resultado(df, transacao, mensagem)
                logger.exception("Erro ao processar linha %s", transacao.linha_excel)

                # Continua para o próximo item, como solicitado.
                continue

    validacao_total = excel_service.validar_total_primeira_aba(arquivo, df)
    logger.info("Validação final: %s", validacao_total)

    arquivo_saida = excel_service.salvar_saida(df, arquivo)
    logger.info("Arquivo final salvo em: %s", arquivo_saida)

    return ResultadoProcessamento(
        arquivo_saida=arquivo_saida,
        total_linhas=len(transacoes),
        total_sucesso=total_sucesso,
        total_erro=total_erro,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Robô STUR - Processamento de vendas")
    parser.add_argument(
        "--arquivo",
        required=False,
        help="Caminho do arquivo Excel/CSV de entrada"
    )
    parser.add_argument("--headless", action="store_true", help="Executa navegador oculto")
    args = parser.parse_args()

    def buscar_arquivo_mais_recente(pasta: Path) -> Path:
        arquivos = list(pasta.glob("*.xlsx")) + list(pasta.glob("*.xls")) + list(pasta.glob("*.csv"))

        if not arquivos:
            raise FileNotFoundError(f"Nenhum arquivo Excel/CSV encontrado em: {pasta}")

        return max(arquivos, key=lambda arquivo: arquivo.stat().st_mtime)
    
    arquivo = Path(args.arquivo) if args.arquivo else buscar_arquivo_mais_recente(
        Path.home() / "Downloads"
    )

    resultado = processar_arquivo(arquivo, headless=args.headless)
    
    print()
    print("Processamento finalizado.")
    print(f"Arquivo saída: {resultado.arquivo_saida}")
    print(f"Total processado: {resultado.total_linhas}")
    print(f"Sucesso: {resultado.total_sucesso}")
    print(f"Erro: {resultado.total_erro}")


if __name__ == "__main__":
    main()