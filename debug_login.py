"""
Script de debug manual — só faz login e tenta acessar a tela de Vendas, sem
mexer em nenhuma venda de verdade. Serve pra observar ao vivo, com o
navegador visível, o que acontece exatamente no ponto onde tem travado.

Uso:
    python3 debug_login.py

O navegador abre visível (headless=False) e fica parado no final (com
input()) pra você poder olhar a tela, inspecionar elementos (botão direito
> Inspecionar) e só then fechar.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from config import load_config
from logger_config import setup_logger
from stur_automation import SturAutomation

config = load_config()
logger = setup_logger(config.logs_dir)

with SturAutomation(config=config, logger=logger, headless=False) as stur:
    print("Fazendo login...")
    stur.login()
    print("Login OK. Tentando acessar tela de Vendas...")
    try:
        stur.acessar_tela_vendas()
        print("Tela de Vendas carregou OK!")
    except Exception as exc:
        print(f"FALHOU: {exc}")
        print("Confira o log e a pasta logs/screenshots/ pra ver o que apareceu na tela.")

    input("\nPressione Enter para fechar o navegador...")
