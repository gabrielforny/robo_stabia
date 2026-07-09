from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path


class ProcessamentoCancelado(Exception):
    """Levantada quando o usuário pede para parar o robô pela GUI.

    Carrega os resultados dos arquivos já concluídos antes do cancelamento,
    para que a GUI possa mostrar um resumo parcial.
    """

    def __init__(self, resultados_parciais: list["ResultadoProcessamento"] | None = None):
        super().__init__("Processamento cancelado pelo usuário.")
        self.resultados_parciais = resultados_parciais or []


@dataclass(slots=True)
class Transacao:
    indice_planilha: int
    linha_excel: int
    estabelecimento: str
    data_aprovacao: str
    data_stur: str
    valor_excel: Decimal | None
    vcn: str
    codigo_venda_vcn: str | None
    localizador_extraido: str | None
    termo_busca: str
    coluna_busca: str
    tipo_busca: str  # VCN, LATAM, GENERICO, GENERICO_COM_LOCALIZADOR
    origem_arquivo: str = ""
    tipo_layout: str = "PADRAO"
    extrato_conta: str = ""
    data_fatura: str = ""
    codigo_autorizacao: str = ""
    venda_ja_ok: bool = False
    resultado_venda_anterior: str = ""


@dataclass(slots=True)
class TransacaoHotel:
    indice_planilha: int
    linha_excel: int
    estabelecimento: str
    data_aprovacao: str
    valor_excel: Decimal | None
    codigo_autorizacao: str
    titular: str
    observacao: str          # col V — chave de busca "Cod. Integração" no STUR
    cliente: str             # col W — filtro por empresa no STUR
    data_fatura: str         # data de hoje no formato dd/mm/aaaa
    origem_arquivo: str = ""


@dataclass(slots=True)
class CandidatoVenda:
    indice_tabela: int
    codigo_venda: str | None
    data_emissao: str | None
    data_inicio: str | None
    data_termino: str | None
    fornecedor: str | None
    fornecedor_servico: str | None
    localizador: str | None
    total_cliente: Decimal | None
    total_fornecedor: Decimal | None
    origem_busca: str
    texto_linha: str
    status: str | None = None


@dataclass(slots=True)
class ResultadoProcessamento:
    arquivo_saida: Path
    total_linhas: int
    total_sucesso: int
    total_erro: int
