from dataclasses import dataclass
from decimal import Decimal
from os import getenv
from pathlib import Path

from dotenv import load_dotenv


@dataclass(frozen=True, slots=True)
class AppConfig:
    stur_url: str
    stur_user: str
    stur_password: str
    excel_sheet_transacoes: str | None
    coluna_localizador: str | None
    coluna_valor_excel: str | None
    coluna_resultado: str
    tolerancia_valor: Decimal
    salvar_screenshot_erro: bool
    base_dir: Path
    input_dir: Path
    output_dir: Path
    logs_dir: Path


def _as_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default

    return value.strip().lower() in {"true", "1", "yes", "sim", "s"}


def load_config() -> AppConfig:
    load_dotenv()

    base_dir = Path(__file__).resolve().parent.parent

    config = AppConfig(
        stur_url=getenv("STUR_URL", "").strip(),
        stur_user=getenv("STUR_USER", "").strip(),
        stur_password=getenv("STUR_PASSWORD", "").strip(),
        excel_sheet_transacoes=(getenv("EXCEL_SHEET_TRANSACOES") or "").strip() or None,
        coluna_localizador=(getenv("COLUNA_LOCALIZADOR") or "").strip() or None,
        coluna_valor_excel=(getenv("COLUNA_VALOR_EXCEL") or "").strip() or None,
        coluna_resultado=(getenv("COLUNA_RESULTADO") or "Resultado Robo").strip(),
        tolerancia_valor=Decimal((getenv("TOLERANCIA_VALOR") or "0.01").strip()),
        salvar_screenshot_erro=_as_bool(getenv("SALVAR_SCREENSHOT_ERRO"), True),
        base_dir=base_dir,
        input_dir=base_dir / "input",
        output_dir=base_dir / "output",
        logs_dir=base_dir / "logs",
    )

    config.input_dir.mkdir(exist_ok=True)
    config.output_dir.mkdir(exist_ok=True)
    config.logs_dir.mkdir(exist_ok=True)

    if not config.stur_url:
        raise ValueError("STUR_URL não configurado no .env")

    if not config.stur_user:
        raise ValueError("STUR_USER não configurado no .env")

    if not config.stur_password:
        raise ValueError("STUR_PASSWORD não configurado no .env")

    return config