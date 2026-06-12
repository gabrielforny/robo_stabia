import argparse
import logging
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import scrolledtext
from pathlib import Path

# Garante que módulos locais são encontrados quando rodando como .exe empacotado
if getattr(sys, "frozen", False):
    sys.path.insert(0, sys._MEIPASS)
else:
    sys.path.insert(0, str(Path(__file__).resolve().parent))


class _QueueHandler(logging.Handler):
    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Robô STUR — LATAM")
        self.geometry("860x560")
        self.minsize(700, 450)
        self.configure(bg="#f5f5f5")
        self._log_queue: queue.Queue = queue.Queue()
        self._build_ui()
        self._poll_logs()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        # ── topo ──────────────────────────────────────────────────────
        top = tk.Frame(self, bg="#f5f5f5", padx=14, pady=12)
        top.pack(fill=tk.X)

        self.btn_iniciar = tk.Button(
            top,
            text="▶   Iniciar Processamento",
            font=("Helvetica", 13, "bold"),
            bg="#1565c0",
            fg="white",
            activebackground="#0d47a1",
            activeforeground="white",
            padx=20,
            pady=10,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._iniciar,
        )
        self.btn_iniciar.pack(side=tk.LEFT)

        self.lbl_status = tk.Label(
            top, text="Pronto para iniciar.", font=("Helvetica", 11),
            fg="#555", bg="#f5f5f5",
        )
        self.lbl_status.pack(side=tk.LEFT, padx=16)

        # ── área de log ───────────────────────────────────────────────
        self.txt_log = scrolledtext.ScrolledText(
            self,
            state=tk.DISABLED,
            font=("Courier New", 10),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="white",
            wrap=tk.WORD,
            padx=10,
            pady=8,
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 8))
        self.txt_log.tag_config("ERROR",   foreground="#f44747")
        self.txt_log.tag_config("WARNING", foreground="#ffcc00")
        self.txt_log.tag_config("OK",      foreground="#6bbf59")
        self.txt_log.tag_config("INFO",    foreground="#9cdcfe")

        # ── painel de resumo (oculto até o fim) ───────────────────────
        self.frame_resumo = tk.Frame(self, bg="#e8f5e9", padx=14, pady=10)
        self.lbl_resumo = tk.Label(
            self.frame_resumo,
            text="",
            font=("Courier New", 11),
            bg="#e8f5e9",
            fg="#1b5e20",
            justify=tk.LEFT,
            anchor="w",
        )
        self.lbl_resumo.pack(fill=tk.X)

    # ------------------------------------------------------------------
    # Processamento
    # ------------------------------------------------------------------

    def _iniciar(self):
        self.btn_iniciar.config(state=tk.DISABLED)
        self.lbl_status.config(text="Processando…", fg="#1565c0")
        self.frame_resumo.pack_forget()
        self._limpar_log()
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        try:
            self._garantir_playwright()

            from config import load_config
            from main import processar_arquivos, resolver_arquivos

            load_config()  # valida .env antes de continuar

            # Redireciona logger para a fila da UI
            logger = logging.getLogger("robo_stur")
            logger.handlers.clear()
            logger.setLevel(logging.INFO)
            handler = _QueueHandler(self._log_queue)
            handler.setFormatter(
                logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s",
                                  datefmt="%H:%M:%S")
            )
            logger.addHandler(handler)

            args = argparse.Namespace(arquivo=None, pasta=None)
            arquivos = resolver_arquivos(args)

            if not arquivos:
                self._log("ERRO: Nenhum arquivo encontrado em ~/Documents/automacao-stur/", "ERROR")
                self.after(0, self._finalizar_erro, "Nenhum arquivo encontrado.")
                return

            resultados = processar_arquivos(arquivos, headless=True)
            self.after(0, self._finalizar_ok, resultados)

        except ValueError as exc:
            # Erros de configuração (.env incompleto etc.)
            self._log(f"ERRO DE CONFIGURAÇÃO: {exc}", "ERROR")
            self.after(0, self._finalizar_erro, str(exc))
        except Exception as exc:
            import traceback
            self._log(f"ERRO FATAL: {exc}", "ERROR")
            self._log(traceback.format_exc(), "ERROR")
            self.after(0, self._finalizar_erro, str(exc))

    def _garantir_playwright(self):
        """Instala o browser Chromium na primeira execução, se necessário."""
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                browser.close()
        except Exception:
            self._log("Instalando Chromium (apenas na primeira execução)…", "WARNING")
            subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                check=True,
                capture_output=True,
            )
            self._log("Chromium instalado com sucesso.", "OK")

    # ------------------------------------------------------------------
    # Callbacks do thread principal
    # ------------------------------------------------------------------

    def _finalizar_ok(self, resultados):
        self.btn_iniciar.config(state=tk.NORMAL)
        self.lbl_status.config(text="Concluído com sucesso!", fg="#2e7d32")

        linhas = ["Processamento finalizado.", "─" * 50]
        for r in resultados:
            linhas += [
                f"Arquivo saída : {r.arquivo_saida}",
                f"Total LATAM   : {r.total_linhas}",
                f"Sucesso Vendas: {r.total_sucesso}",
                f"Erro Vendas   : {r.total_erro}",
                "",
            ]

        self.lbl_resumo.config(text="\n".join(linhas).strip(), fg="#1b5e20",
                                bg="#e8f5e9")
        self.frame_resumo.pack(fill=tk.X, padx=14, pady=(0, 12))

    def _finalizar_erro(self, msg: str):
        self.btn_iniciar.config(state=tk.NORMAL)
        self.lbl_status.config(text=f"Erro: {msg}", fg="#c62828")
        self.lbl_resumo.config(
            text=f"Processamento encerrado com erro.\n{msg}",
            fg="#b71c1c", bg="#ffebee",
        )
        self.frame_resumo.configure(bg="#ffebee")
        self.frame_resumo.pack(fill=tk.X, padx=14, pady=(0, 12))

    # ------------------------------------------------------------------
    # Log helpers
    # ------------------------------------------------------------------

    def _poll_logs(self):
        while True:
            try:
                msg = self._log_queue.get_nowait()
                self._append_log(msg)
            except queue.Empty:
                break
        self.after(100, self._poll_logs)

    def _log(self, msg: str, tag: str = "INFO"):
        self._log_queue.put((msg, tag))

    def _append_log(self, item):
        if isinstance(item, tuple):
            msg, tag = item
        else:
            msg = item
            upper = msg.upper()
            if "| ERROR" in upper or "ERRO" in upper:
                tag = "ERROR"
            elif "| WARNING" in upper or "AVISO" in upper:
                tag = "WARNING"
            elif upper.startswith("OK") or "SUCESSO" in upper:
                tag = "OK"
            else:
                tag = "INFO"

        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.insert(tk.END, msg + "\n", tag)
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def _limpar_log(self):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state=tk.DISABLED)


if __name__ == "__main__":
    app = App()
    app.mainloop()
