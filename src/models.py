from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path


@dataclass(slots=True)
class Transacao:
    indice_planilha: int
    linha_excel: int
    localizador_original: str
    codigo_companhia: str
    valor_excel: Decimal | None


@dataclass(slots=True)
class ResultadoVenda:
    encontrada: bool
    codigo_venda: str | None = None
    total_fornecedor: Decimal | None = None
    mensagem: str = ""


@dataclass(slots=True)
class ResultadoProcessamento:
    arquivo_saida: Path
    total_linhas: int
    total_sucesso: int
    total_erro: int