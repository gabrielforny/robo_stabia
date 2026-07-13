
import argparse
import os
import shutil
import sys
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from config import load_config
from excel_service import ExcelService
from logger_config import setup_logger
from models import ProcessamentoCancelado, ResultadoProcessamento, Transacao
from stur_automation import SturAutomation, VendaJaFaturadaError
from stur_financeiro_automation import SturFinanceiroAutomation


EXTENSOES_SUPORTADAS = {".xlsx", ".xls", ".csv"}
COMPANHIAS_SUPORTADAS = {"LATAM", "GOL", "AZUL"}
MAX_TENTATIVAS_POR_ITEM = 3


def _pasta_documentos() -> Path:
    """
    Encontra a pasta real de Documentos do usuário no Windows.

    Path.home() / "Documents" não funciona quando o OneDrive redireciona a
    pasta Documentos (Backup de Pastas Conhecidas) — nesse caso o caminho real
    fica em algo como "OneDrive - EMPRESA\\Documentos". O registro do Windows
    sempre reflete o local atual, então é a fonte confiável aqui.
    """
    if sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            ) as key:
                valor, _ = winreg.QueryValueEx(key, "Personal")
                return Path(os.path.expandvars(valor))
        except Exception:
            pass

    return Path.home() / "Documents"


PASTA_AUTOMACAO_STUR = _pasta_documentos() / "automacao-stur"
PASTA_FINALIZADAS = PASTA_AUTOMACAO_STUR / "finalizadas"


def mover_para_finalizadas(arquivo_saida: Path) -> Path:
    """Move o arquivo já processado (com cores) para automacao-stur/finalizadas,
    renomeando com sufixo -finalizada-DD-MM-AAAA-HH-MM."""
    PASTA_FINALIZADAS.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%d-%m-%Y-%H-%M")
    novo_nome = f"{arquivo_saida.stem}-finalizada-{timestamp}{arquivo_saida.suffix}"
    destino = PASTA_FINALIZADAS / novo_nome
    shutil.move(str(arquivo_saida), str(destino))
    return destino


# ==========================================================
# FASE 1 — VENDAS (Operacional → Vendas)
# ==========================================================

def processar_latam_vendas(
    config,
    headless: bool,
    excel_service: ExcelService,
    df,
    transacoes_latam: list[Transacao],
    logger,
    deve_parar=None,
) -> tuple[int, int]:
    """
    Para cada item LATAM: busca o localizador na tela de Vendas, valida
    Total Fornecedor e executa seguir_fluxo_venda_ok.

    Gerencia o ciclo de vida do browser internamente. Em caso de falha em
    qualquer item, fecha o browser, reabre, faz login novamente e retenta
    o mesmo item — até MAX_TENTATIVAS_POR_ITEM vezes. Se todas as tentativas
    falharem, registra erro na planilha e segue para o próximo item.
    """
    total_sucesso = 0
    total_erro = 0
    stur: SturAutomation | None = None

    def _fechar_sessao():
        nonlocal stur
        if stur is not None:
            try:
                stur.__exit__(None, None, None)
            except Exception:
                pass
            stur = None

    def _abrir_nova_sessao():
        nonlocal stur
        _fechar_sessao()
        s = SturAutomation(config=config, logger=logger, headless=headless)
        s.__enter__()
        try:
            s.login()
            s.acessar_tela_vendas()
            s.garantir_coluna_localizador_visivel()
        except Exception:
            try:
                s.__exit__(None, None, None)
            except Exception:
                pass
            raise
        stur = s

    try:
        _abrir_nova_sessao()

        for transacao in transacoes_latam:
            if deve_parar and deve_parar():
                logger.warning("Parada solicitada — interrompendo Fase 1 (Vendas).")
                raise ProcessamentoCancelado()

            if transacao.venda_ja_ok:
                logger.info("Venda já OK para %s — pulando Fase 1.", transacao.localizador_extraido)
                continue

            if not transacao.localizador_extraido:
                excel_service.escrever_resultado(df, transacao, "ERRO | LATAM sem localizador extraído")
                total_erro += 1
                continue

            ultima_exc: Exception | None = None
            resultado_msg: str | None = None
            foi_sucesso = False

            for tentativa in range(1, MAX_TENTATIVAS_POR_ITEM + 1):
                if deve_parar and deve_parar():
                    raise ProcessamentoCancelado()

                if tentativa > 1:
                    logger.warning(
                        "Tentativa %d/%d para localizador %s — fechando browser e reabrindo do zero...",
                        tentativa, MAX_TENTATIVAS_POR_ITEM, transacao.localizador_extraido,
                    )
                    try:
                        _abrir_nova_sessao()
                    except Exception as exc_sessao:
                        logger.error(
                            "Falha ao reabrir browser na tentativa %d: %s",
                            tentativa, exc_sessao,
                        )
                        ultima_exc = exc_sessao
                        continue

                logger.info(
                    "Buscando localizador nas Vendas: %s%s",
                    transacao.localizador_extraido,
                    f" (tentativa {tentativa}/{MAX_TENTATIVAS_POR_ITEM})" if tentativa > 1 else "",
                )

                try:
                    candidatos = stur.buscar_latam_por_localizador(transacao)
                except Exception as exc:
                    logger.warning(
                        "Erro ao buscar localizador %s (tentativa %d/%d): %s",
                        transacao.localizador_extraido, tentativa, MAX_TENTATIVAS_POR_ITEM, exc,
                    )
                    ultima_exc = exc
                    continue

                if not candidatos:
                    resultado_msg = (
                        f"ERRO | Localizador {transacao.localizador_extraido} "
                        f"não encontrado nas Vendas"
                    )
                    break  # não é falha de browser — não retentar

                candidato_ok = None
                candidato_comissao = None
                valor_comissao = None
                for c in candidatos:
                    if transacao.valor_excel is not None and c.total_fornecedor is not None:
                        if abs(c.total_fornecedor) == abs(transacao.valor_excel):
                            candidato_ok = c
                            break
                        # Verifica se a diferença é de até 20% (tabela > excel)
                        diferenca = abs(c.total_fornecedor) - abs(transacao.valor_excel)
                        if diferenca > 0 and diferenca / abs(c.total_fornecedor) <= Decimal("0.20"):
                            candidato_comissao = c
                            valor_comissao = diferenca

                if not candidato_ok and not candidato_comissao:
                    vals = [str(c.total_fornecedor) for c in candidatos]
                    resultado_msg = (
                        f"ERRO | Valor não bate nas Vendas | "
                        f"Excel={transacao.valor_excel} | Tabela={vals}"
                    )
                    break  # não é falha de browser — não retentar

                candidato_final = candidato_ok or candidato_comissao
                logger.info(
                    "Localizador %s → Venda=%s | TotalForn=%s%s",
                    transacao.localizador_extraido,
                    candidato_final.codigo_venda,
                    candidato_final.total_fornecedor,
                    f" | Comissão={valor_comissao}" if valor_comissao else "",
                )

                try:
                    if candidato_ok:
                        stur.seguir_fluxo_venda_ok(
                            candidato_ok, codigo_autorizacao=transacao.codigo_autorizacao
                        )
                    else:
                        stur.seguir_fluxo_venda_com_comissao(
                            candidato_comissao,
                            valor_comissao=valor_comissao,
                            codigo_autorizacao=transacao.codigo_autorizacao,
                        )
                    resultado_msg = (
                        f"OK Vendas | Venda {candidato_final.codigo_venda} | "
                        f"Loc {transacao.localizador_extraido}"
                    )
                    foi_sucesso = True
                    ultima_exc = None
                    break
                except VendaJaFaturadaError:
                    logger.warning("Venda %s já faturada — marcando e seguindo.", candidato_final.codigo_venda)
                    resultado_msg = (
                        f"JÁ FATURADO | Venda {candidato_final.codigo_venda} | "
                        f"Loc {transacao.localizador_extraido}"
                    )
                    break  # condição esperada — não retentar
                except Exception as exc:
                    logger.warning(
                        "Erro ao processar venda %s (tentativa %d/%d): %s",
                        candidato_final.codigo_venda, tentativa, MAX_TENTATIVAS_POR_ITEM, exc,
                    )
                    ultima_exc = exc
                    continue

            # Esgotadas as tentativas sem resultado definido
            if resultado_msg is None:
                resultado_msg = (
                    f"ERRO | {MAX_TENTATIVAS_POR_ITEM} tentativas falharam para "
                    f"{transacao.localizador_extraido} | "
                    f"{type(ultima_exc).__name__}: {ultima_exc}"
                )

            excel_service.escrever_resultado(df, transacao, resultado_msg)
            if foi_sucesso:
                total_sucesso += 1
            else:
                total_erro += 1

    finally:
        _fechar_sessao()

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
    deve_parar=None,
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
        if deve_parar and deve_parar():
            logger.warning("Parada solicitada — interrompendo Fase 2 (Conferências).")
            raise ProcessamentoCancelado()

        descricao_busca = f"Clara {chave_mes}"
        descricao_criar = f"Clara {chave_mes}"
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

            ok_conferencia: list[Transacao] = []

            for transacao in grupo:
                if deve_parar and deve_parar():
                    logger.warning("Parada solicitada — interrompendo seleção de localizadores.")
                    raise ProcessamentoCancelado()

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
                    ok_conferencia.append(transacao)
                    resultado_conf = f"OK Conferência | {descricao_busca} | Loc {transacao.localizador_extraido}"
                    if transacao.venda_ja_ok:
                        # Sobrescreve limpo: remove o ERRO Conferência anterior
                        excel_service.escrever_resultado(
                            df, transacao,
                            f"{transacao.resultado_venda_anterior} | {resultado_conf}",
                        )
                    else:
                        excel_service.acrescentar_resultado(df, transacao, resultado_conf)
                else:
                    if transacao.venda_ja_ok:
                        excel_service.escrever_resultado(
                            df, transacao,
                            f"{transacao.resultado_venda_anterior} | ERRO Conferência | {motivo}",
                        )
                    else:
                        excel_service.acrescentar_resultado(
                            df, transacao, f"ERRO Conferência | {motivo}"
                        )

            soma_ok = sum(t.valor_excel for t in ok_conferencia if t.valor_excel is not None)
            logger.info(
                "Conferência %s — %d/%d adicionados | Soma dos valores: R$ %s",
                descricao_busca,
                len(ok_conferencia),
                len(grupo),
                soma_ok,
            )

            financeiro.gravar_titulos()
            financeiro.gravar_conferencia()

        except ProcessamentoCancelado:
            raise
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
    config,
    headless: bool,
    logger,
    deve_parar=None,
) -> ResultadoProcessamento:
    logger.info("Iniciando processamento do arquivo: %s", arquivo)

    df, aba = excel_service.carregar_transacoes(arquivo)
    logger.info("Aba/tipo carregado: %s | Colunas: %s", aba, list(df.columns))

    transacoes, transacoes_negativas = excel_service.montar_transacoes(df, origem_arquivo=arquivo.name)
    logger.info("Total de linhas a processar: %d | Valores negativos (reservados): %d", len(transacoes), len(transacoes_negativas))

    transacoes_latam = [t for t in transacoes if t.tipo_busca in COMPANHIAS_SUPORTADAS]
    transacoes_outras = [t for t in transacoes if t.tipo_busca not in COMPANHIAS_SUPORTADAS]

    logger.info(
        "LATAM/GOL/AZUL: %d | Outros (ignorados nesta versão): %d",
        len(transacoes_latam), len(transacoes_outras),
    )

    total_sucesso = 0
    total_erro = 0

    if transacoes_latam:
        try:
            # Fase 1: Vendas — gerencia o próprio browser com retry por item
            logger.info("=== FASE 1: Vendas ===")
            sucesso_v, erro_v = processar_latam_vendas(
                config=config,
                headless=headless,
                excel_service=excel_service,
                df=df,
                transacoes_latam=transacoes_latam,
                logger=logger,
                deve_parar=deve_parar,
            )
            total_sucesso += sucesso_v
            total_erro += erro_v

            arquivo_parcial = excel_service.salvar_saida(df, arquivo)
            logger.info("Backup parcial (pós Vendas) salvo em: %s", arquivo_parcial)
            excel_service.salvar_no_local_com_cores(df, arquivo)

            # Fase 2: Conferências — nova sessão de browser
            logger.info("=== FASE 2: Conferências ===")
            with SturAutomation(config=config, logger=logger, headless=headless) as stur_conf:
                stur_conf.login()
                financeiro = SturFinanceiroAutomation(
                    page=stur_conf._page(),
                    logger=logger,
                    espera_padrao_segundos=3,
                )
                financeiro.acessar_tela_conferencias_baixas()
                processar_latam_conferencia(
                    financeiro=financeiro,
                    excel_service=excel_service,
                    df=df,
                    transacoes_latam=transacoes_latam,
                    logger=logger,
                    deve_parar=deve_parar,
                )
        except ProcessamentoCancelado:
            excel_service.salvar_no_local_com_cores(df, arquivo)
            logger.warning(
                "Processamento interrompido pelo usuário — progresso parcial salvo em: %s", arquivo
            )
            raise
        except Exception:
            logger.exception("Erro inesperado durante processamento do arquivo %s", arquivo)
            raise

    excel_service.salvar_saida(df, arquivo)
    arquivo_saida = excel_service.salvar_no_local_com_cores(df, arquivo)
    logger.info("Arquivo final salvo: %s", arquivo_saida)

    # Quando a entrada é CSV, salvar_no_local_com_cores cria um .xlsx ao lado e
    # deixa o .csv original intacto — sem isso ele ficaria na pasta e seria
    # reprocessado na próxima execução.
    if arquivo != arquivo_saida and arquivo.exists():
        arquivo.unlink()

    arquivo_saida = mover_para_finalizadas(arquivo_saida)
    logger.info("Arquivo movido para finalizadas: %s", arquivo_saida)

    return ResultadoProcessamento(
        arquivo_saida=arquivo_saida,
        total_linhas=len(transacoes_latam),
        total_sucesso=total_sucesso,
        total_erro=total_erro,
    )


# ==========================================================
# PONTO DE ENTRADA
# ==========================================================

def processar_arquivos(
    arquivos: list[Path],
    headless: bool,
    logger=None,
    deve_parar=None,
) -> list[ResultadoProcessamento]:
    config = load_config()
    if logger is None:
        logger = setup_logger(config.logs_dir)
    else:
        # Adiciona FileHandler diretamente — setup_logger faz handlers.clear() e
        # apagaria o queue handler da GUI, fazendo os logs sumirem da tela.
        import logging as _logging
        fh_existe = any(isinstance(h, _logging.FileHandler) for h in logger.handlers)
        if not fh_existe:
            from datetime import datetime as _dt
            config.logs_dir.mkdir(exist_ok=True)
            log_file = config.logs_dir / f"robo_stur_{_dt.now():%Y%m%d_%H%M%S}.log"
            fh = _logging.FileHandler(log_file, encoding="utf-8")
            fh.setFormatter(_logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s"))
            fh.setLevel(_logging.INFO)
            logger.addHandler(fh)

    excel_service = ExcelService(config)

    arquivos = [a for a in arquivos if a.suffix.lower() in EXTENSOES_SUPORTADAS]

    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo Excel/CSV válido encontrado para processamento.")

    logger.info("Iniciando processamento — fluxo LATAM: Vendas + Conferências")
    logger.info("Arquivos recebidos: %s", [str(a) for a in arquivos])

    resultados: list[ResultadoProcessamento] = []

    for arquivo in arquivos:
        if deve_parar and deve_parar():
            logger.warning("Parada solicitada — encerrando antes do próximo arquivo.")
            raise ProcessamentoCancelado(resultados)

        try:
            resultado = processar_arquivo_aberto(
                arquivo=arquivo,
                excel_service=excel_service,
                config=config,
                headless=headless,
                logger=logger,
                deve_parar=deve_parar,
            )
        except ProcessamentoCancelado as exc:
            exc.resultados_parciais = resultados
            raise
        except Exception:
            logger.exception("Erro inesperado ao processar %s", arquivo)
            raise

        resultados.append(resultado)

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

    # Padrão: pasta ~/Documents/automacao-stur — pega só o arquivo mais recente.
    # Sem fallback para Downloads/projeto: se a pasta não existir ou estiver vazia,
    # é melhor avisar o usuário do que processar arquivos de outro lugar por engano.
    PASTA_AUTOMACAO_STUR.mkdir(parents=True, exist_ok=True)
    arquivos = listar_arquivos_da_pasta(PASTA_AUTOMACAO_STUR)
    if arquivos:
        return [arquivos[-1]]

    return []


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
