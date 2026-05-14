import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd

from config import AppConfig
from models import Transacao

MESES_PT = {
    "Jan": "01", "Fev": "02", "Mar": "03", "Abr": "04",
    "Mai": "05", "Jun": "06", "Jul": "07", "Ago": "08",
    "Set": "09", "Out": "10", "Nov": "11", "Dez": "12",
}


def converter_data_excel_para_stur(data_excel: str) -> str:
    """Converte datas do Excel para dd/mm/aaaa quando possível."""
    texto = str(data_excel or "").strip()
    if not texto or texto.lower() == "nan":
        return ""

    # Ex: 28/Mar/2026
    partes = texto.split("/")
    if len(partes) == 3:
        dia, mes, ano = partes
        mes_num = MESES_PT.get(mes.strip().title(), mes.strip())
        return f"{dia.strip().zfill(2)}/{mes_num.zfill(2)}/{ano.strip()}"

    # Ex: pandas Timestamp ou string ISO
    try:
        dt = pd.to_datetime(texto, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    return texto


class ExcelService:
    def __init__(self, config: AppConfig):
        self.config = config

    def carregar_transacoes(self, arquivo: Path) -> tuple[pd.DataFrame, str]:
        if not arquivo.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {arquivo}")

        extensao = arquivo.suffix.lower()

        if extensao == ".csv":
            df = pd.read_csv(arquivo, dtype=str, sep=None, engine="python")
            return self._normalizar_df(df), "CSV"

        if extensao in {".xlsx", ".xls"}:
            sheet_name = self._definir_aba_transacoes(arquivo)
            header_row = self._encontrar_linha_cabecalho(arquivo, sheet_name)
            df = pd.read_excel(arquivo, sheet_name=sheet_name, dtype=str, header=header_row)
            return self._normalizar_df(df), str(sheet_name)

        raise ValueError(f"Formato não suportado: {extensao}")

    def montar_transacoes(self, df: pd.DataFrame) -> list[Transacao]:
        coluna_estabelecimento = self._resolver_coluna_estabelecimento(df)
        coluna_data = self._resolver_coluna_data_aprovacao(df)
        coluna_valor = self._resolver_coluna_valor(df)
        coluna_vcn = self._resolver_coluna_vcn(df)

        transacoes: list[Transacao] = []

        for index, row in df.iterrows():
            estabelecimento = str(row.get(coluna_estabelecimento, "") or "").strip()
            if not estabelecimento or estabelecimento.lower() == "nan":
                continue

            data_excel = str(row.get(coluna_data, "") or "").strip() if coluna_data else ""
            data_stur = converter_data_excel_para_stur(data_excel)
            valor_excel = self._parse_decimal(row.get(coluna_valor)) if coluna_valor else None
            vcn = str(row.get(coluna_vcn, "") or "").strip() if coluna_vcn else ""

            codigo_venda_vcn = self._extrair_codigo_venda_vcn(vcn)
            localizador_extraido = self._extrair_localizador_apos_asterisco(estabelecimento)
            termo_busca, coluna_busca, tipo_busca = self._definir_estrategia_busca(
                estabelecimento=estabelecimento,
                codigo_venda_vcn=codigo_venda_vcn,
                localizador_extraido=localizador_extraido,
            )

            transacoes.append(
                Transacao(
                    indice_planilha=index,
                    linha_excel=index + 2,
                    estabelecimento=estabelecimento,
                    data_aprovacao=data_excel,
                    data_stur=data_stur,
                    valor_excel=valor_excel,
                    vcn=vcn,
                    codigo_venda_vcn=codigo_venda_vcn,
                    localizador_extraido=localizador_extraido,
                    termo_busca=termo_busca,
                    coluna_busca=coluna_busca,
                    tipo_busca=tipo_busca,
                )
            )

        return transacoes

    def escrever_resultado(self, df: pd.DataFrame, transacao: Transacao, resultado: str) -> None:
        coluna = self.config.coluna_resultado
        if coluna not in df.columns:
            df[coluna] = ""
        df.at[transacao.indice_planilha, coluna] = resultado

    def salvar_saida(self, df: pd.DataFrame, arquivo_original: Path) -> Path:
        saida = self.config.output_dir / f"{arquivo_original.stem}_processado.xlsx"
        df.to_excel(saida, index=False)
        return saida

    def validar_total_primeira_aba(self, arquivo: Path, df_processado: pd.DataFrame) -> str:
        if arquivo.suffix.lower() not in {".xlsx", ".xls"}:
            return "Validação de total não aplicada para CSV."

        primeira_aba = pd.read_excel(arquivo, sheet_name=0, header=None, dtype=str)
        total_a_pagar = None

        for _, row in primeira_aba.iterrows():
            valores = [str(v).strip() for v in row.tolist() if str(v).strip() and str(v) != "nan"]
            linha_texto = " ".join(valores).lower()
            if "total a pagar" in linha_texto:
                total_a_pagar = self._primeiro_decimal_na_linha(valores)
                break

        if total_a_pagar is None:
            return "Não encontrei 'Total a pagar' na primeira aba."

        return f"Total a pagar localizado na primeira aba: {total_a_pagar}"

    def _definir_estrategia_busca(
        self,
        estabelecimento: str,
        codigo_venda_vcn: str | None,
        localizador_extraido: str | None,
    ) -> tuple[str, str, str]:
        """
        Define a primeira estratégia da linha.

        Regras:
        - Se o VCN trouxer venda, busca direto por Venda.
        - Se for LATAM e tiver '*', busca direto pelo Localizador extraído.
        - Se NÃO for LATAM, mas tiver '*', mantém o estabelecimento como termo principal,
          porém sinaliza que também deve tentar o Localizador extraído.
        - Se não tiver '*', fluxo genérico normal.
        """
        if codigo_venda_vcn:
            return codigo_venda_vcn, "Venda", "VCN"

        if localizador_extraido and "latam" in estabelecimento.lower():
            return localizador_extraido, "Localizador", "LATAM"

        if localizador_extraido:
            return estabelecimento, "Fornecedor", "GENERICO_COM_LOCALIZADOR"

        return estabelecimento, "Fornecedor", "GENERICO"

    def _extrair_localizador_apos_asterisco(self, estabelecimento: str) -> str | None:
        texto = str(estabelecimento or "").strip()
        if "*" not in texto:
            return None

        match = re.search(r"\*\s*([A-Za-z0-9]{6})", texto)
        if not match:
            return None

        return match.group(1).upper()

    def _extrair_codigo_venda_vcn(self, vcn: str) -> str | None:
        """
        Cenário futuro explicado pelo cliente:
        VCN pode vir com algo como 'Venda 2238' ou texto contendo o número da venda.
        """
        texto = str(vcn or "").strip()
        if not texto or texto.lower() == "nan":
            return None

        match = re.search(r"venda\D*(\d{3,})", texto, flags=re.IGNORECASE)
        if match:
            return match.group(1)

        # fallback conservador: se o VCN for só número, assume venda
        if re.fullmatch(r"\d{3,}", texto):
            return texto

        return None

    def _encontrar_linha_cabecalho(self, arquivo: Path, sheet_name: str | int) -> int:
        amostra = pd.read_excel(arquivo, sheet_name=sheet_name, header=None, nrows=12, dtype=str)
        palavras_chave = {"estabelecimento", "valor em r$", "data de aprovação", "vcn"}
        for i, row in amostra.iterrows():
            valores = [str(v).strip().lower() for v in row if str(v).strip() and str(v) != "nan"]
            if any(any(kw in v for kw in palavras_chave) for v in valores):
                return i
        return 0

    def _definir_aba_transacoes(self, arquivo: Path) -> str | int:
        if self.config.excel_sheet_transacoes:
            return self.config.excel_sheet_transacoes

        excel = pd.ExcelFile(arquivo)
        for aba in excel.sheet_names:
            if self._normalizar_texto(aba) in {"transacoes", "transacao", "transactions"}:
                return aba
        return 0

    def _normalizar_df(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(col).strip() for col in df.columns]
        df = df.dropna(how="all").reset_index(drop=True)
        return df

    def _resolver_coluna_estabelecimento(self, df: pd.DataFrame) -> str:
        candidatos = ["estabelecimento", "fornecedor", "descricao", "descrição", "historico", "histórico"]
        return self._procurar_coluna(df, candidatos, obrigatoria=True, finalidade="estabelecimento")

    def _resolver_coluna_data_aprovacao(self, df: pd.DataFrame) -> str | None:
        candidatos = ["data de aprovação", "data de aprovacao", "data aprovacao", "data aprovação", "data da transação", "data da transacao", "data"]
        return self._procurar_coluna(df, candidatos, obrigatoria=False, finalidade="data de aprovação")

    def _resolver_coluna_valor(self, df: pd.DataFrame) -> str | None:
        candidatos = ["valor em r$", "valor em reais", "valor", "valor pago", "total", "total a pagar"]
        return self._procurar_coluna(df, candidatos, obrigatoria=False, finalidade="valor")

    def _resolver_coluna_vcn(self, df: pd.DataFrame) -> str | None:
        candidatos = ["vcn", "venda", "codigo venda", "código venda"]
        return self._procurar_coluna(df, candidatos, obrigatoria=False, finalidade="VCN")

    def _procurar_coluna(self, df: pd.DataFrame, candidatos: list[str], obrigatoria: bool, finalidade: str) -> str | None:
        mapa = {self._normalizar_texto(col): col for col in df.columns}
        for candidato in candidatos:
            normalizado = self._normalizar_texto(candidato)
            if normalizado in mapa:
                return mapa[normalizado]

        if obrigatoria:
            raise ValueError(
                f"Não consegui identificar a coluna de {finalidade}. Colunas encontradas: {list(df.columns)}"
            )
        return None

    def _parse_decimal(self, value) -> Decimal | None:
        if value is None:
            return None
        texto = str(value).strip()
        if not texto or texto.lower() == "nan":
            return None
        texto = texto.replace("R$", "").replace(" ", "")
        if "," in texto and "." in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        texto = re.sub(r"[^0-9.-]", "", texto)
        if not texto:
            return None
        try:
            return Decimal(texto)
        except InvalidOperation:
            return None

    def _primeiro_decimal_na_linha(self, valores: list[str]) -> Decimal | None:
        for valor in valores:
            parsed = self._parse_decimal(valor)
            if parsed is not None:
                return parsed
        return None

    def _normalizar_texto(self, texto: str) -> str:
        return (
            str(texto).strip().lower()
            .replace("ç", "c")
            .replace("ã", "a").replace("á", "a").replace("à", "a").replace("â", "a")
            .replace("é", "e").replace("ê", "e")
            .replace("í", "i")
            .replace("ó", "o").replace("ô", "o").replace("õ", "o")
            .replace("ú", "u")
        )
