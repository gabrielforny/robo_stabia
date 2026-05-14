from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path


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


@dataclass(slots=True)
class ResultadoProcessamento:
    arquivo_saida: Path
    total_linhas: int
    total_sucesso: int
    total_erro: int
