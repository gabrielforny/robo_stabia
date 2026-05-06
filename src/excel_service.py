import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd

from config import AppConfig
from models import Transacao


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
            df = pd.read_excel(arquivo, sheet_name=sheet_name, dtype=str)
            return self._normalizar_df(df), str(sheet_name)

        raise ValueError(f"Formato não suportado: {extensao}")

    def _definir_aba_transacoes(self, arquivo: Path) -> str | int:
        if self.config.excel_sheet_transacoes:
            return self.config.excel_sheet_transacoes

        excel = pd.ExcelFile(arquivo)
        abas = excel.sheet_names

        for aba in abas:
            if self._normalizar_texto(aba) in {"transacoes", "transação", "transacao", "transactions"}:
                return aba

        return 0

    def _normalizar_df(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(col).strip() for col in df.columns]
        df = df.dropna(how="all").reset_index(drop=True)
        return df

    def montar_transacoes(self, df: pd.DataFrame) -> list[Transacao]:
        coluna_localizador = self._resolver_coluna_localizador(df)
        coluna_valor = self._resolver_coluna_valor(df)

        transacoes: list[Transacao] = []

        for index, row in df.iterrows():
            localizador_original = str(row.get(coluna_localizador, "") or "").strip()
            codigo_companhia = self.extrair_codigo_companhia(localizador_original)

            if not codigo_companhia:
                continue

            valor_excel = self._parse_decimal(row.get(coluna_valor)) if coluna_valor else None

            transacoes.append(
                Transacao(
                    indice_planilha=index,
                    linha_excel=index + 2,
                    localizador_original=localizador_original,
                    codigo_companhia=codigo_companhia,
                    valor_excel=valor_excel,
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
        """
        Base para validar a primeira aba com a última linha "Total a pagar".

        Como o layout real pode variar, deixei uma busca genérica:
        - Procura por qualquer célula contendo "Total a pagar"
        - Tenta capturar o primeiro valor numérico na mesma linha
        """
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

    def extrair_codigo_companhia(self, valor: str) -> str | None:
        """
        Regra:
        - Procura o conteúdo depois do *
        - Captura 6 caracteres alfanuméricos ou 6 dígitos

        Exemplos:
        ABC*123456 -> 123456
        LOCALIZADOR * G3AB12 -> G3AB12
        """
        if not valor or "*" not in valor:
            return None

        depois_asterisco = valor.split("*", 1)[1].strip()
        match = re.search(r"([A-Za-z0-9]{6})", depois_asterisco)

        if not match:
            return None

        return match.group(1).upper()

    def _resolver_coluna_localizador(self, df: pd.DataFrame) -> str:
        if self.config.coluna_localizador and self.config.coluna_localizador in df.columns:
            return self.config.coluna_localizador

        candidatos = [
            "localizador",
            "localizador companhia",
            "localizador da companhia",
            "descricao",
            "descrição",
            "historico",
            "histórico",
            "documento",
            "referencia",
            "referência",
        ]

        return self._procurar_coluna(df, candidatos, obrigatoria=True, finalidade="localizador")

    def _resolver_coluna_valor(self, df: pd.DataFrame) -> str | None:
        if self.config.coluna_valor_excel and self.config.coluna_valor_excel in df.columns:
            return self.config.coluna_valor_excel

        candidatos = [
            "valor",
            "valor excel",
            "valor pago",
            "valor pagamento",
            "total",
            "total a pagar",
            "valor transacao",
            "valor transação",
        ]

        return self._procurar_coluna(df, candidatos, obrigatoria=False, finalidade="valor")

    def _procurar_coluna(
        self,
        df: pd.DataFrame,
        candidatos: list[str],
        obrigatoria: bool,
        finalidade: str,
    ) -> str | None:
        mapa = {self._normalizar_texto(col): col for col in df.columns}

        for candidato in candidatos:
            normalizado = self._normalizar_texto(candidato)
            if normalizado in mapa:
                return mapa[normalizado]

        if finalidade == "localizador":
            for coluna in df.columns:
                serie = df[coluna].dropna().astype(str)
                if serie.str.contains(r"\*[A-Za-z0-9]{6}", regex=True, na=False).any():
                    return coluna

        if obrigatoria:
            raise ValueError(
                f"Não consegui identificar a coluna de {finalidade}. "
                f"Configure no .env. Colunas encontradas: {list(df.columns)}"
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
            str(texto)
            .strip()
            .lower()
            .replace("ç", "c")
            .replace("ã", "a")
            .replace("á", "a")
            .replace("à", "a")
            .replace("â", "a")
            .replace("é", "e")
            .replace("ê", "e")
            .replace("í", "i")
            .replace("ó", "o")
            .replace("ô", "o")
            .replace("õ", "o")
            .replace("ú", "u")
        )