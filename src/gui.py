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
        self._stop_event = threading.Event()
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

        self.btn_conferencia = tk.Button(
            top,
            text="🗂   Só Conferência",
            font=("Helvetica", 13, "bold"),
            bg="#2e7d32",
            fg="white",
            activebackground="#1b5e20",
            activeforeground="white",
            padx=20,
            pady=10,
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self._iniciar(somente_conferencia=True),
        )
        self.btn_conferencia.pack(side=tk.LEFT, padx=(10, 0))

        self.btn_parar = tk.Button(
            top,
            text="■   Parar",
            font=("Helvetica", 13, "bold"),
            bg="#c62828",
            fg="white",
            activebackground="#8e0000",
            activeforeground="white",
            padx=20,
            pady=10,
            relief=tk.FLAT,
            cursor="hand2",
            state=tk.DISABLED,
            command=self._parar,
        )
        self.btn_parar.pack(side=tk.LEFT, padx=(10, 0))

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

    def _iniciar(self, somente_conferencia: bool = False):
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_conferencia.config(state=tk.DISABLED)
        self.btn_parar.config(state=tk.NORMAL, text="■   Parar")
        if somente_conferencia:
            self.lbl_status.config(text="Processando (só conferência)…", fg="#2e7d32")
        else:
            self.lbl_status.config(text="Processando…", fg="#1565c0")
        self.frame_resumo.pack_forget()
        self._limpar_log()
        self._stop_event.clear()
        threading.Thread(target=self._run, args=(somente_conferencia,), daemon=True).start()

    def _parar(self):
        self.btn_parar.config(state=tk.DISABLED, text="Parando…")
        self.lbl_status.config(text="Parando com segurança…", fg="#c62828")
        self._log("Solicitação de parada recebida — encerrando no próximo ponto seguro (pode levar alguns segundos)…", "WARNING")
        self._stop_event.set()

    def _run(self, somente_conferencia: bool = False):
        try:
            # Quando rodando como .exe empacotado, o Playwright extrai seus arquivos numa
            # pasta temporária mas precisa encontrar o Chromium onde foi instalado de fato.
            if getattr(sys, "frozen", False):
                import os
                import platform
                if platform.system() == "Windows":
                    localappdata = os.environ.get("LOCALAPPDATA") or str(
                        Path.home() / "AppData" / "Local"
                    )
                    browsers_path = Path(localappdata) / "ms-playwright"
                else:
                    browsers_path = Path.home() / ".cache" / "ms-playwright"
                os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers_path)
                self._log(f"Browsers Playwright em: {browsers_path}", "INFO")

            self._garantir_playwright()

            from config import load_config, _base_dir_padrao
            from main import PASTA_AUTOMACAO_STUR, processar_arquivos, resolver_arquivos
            from models import ProcessamentoCancelado

            env_esperado = _base_dir_padrao() / ".env"
            self._log(f"Procurando .env em: {env_esperado}", "INFO")
            self._log(f"Diretório de trabalho: {Path.cwd()}", "INFO")

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
                self._log(f"ERRO: Nenhum arquivo encontrado em {PASTA_AUTOMACAO_STUR}", "ERROR")
                self.after(0, self._finalizar_erro, "Nenhum arquivo encontrado.")
                return

            resultados = processar_arquivos(
                arquivos, headless=False, logger=logger,
                deve_parar=self._stop_event.is_set,
                somente_conferencia=somente_conferencia,
            )
            self.after(0, self._finalizar_ok, resultados)

        except ProcessamentoCancelado as exc:
            self._log("Processamento interrompido pelo usuário.", "WARNING")
            self.after(0, self._finalizar_cancelado, exc.resultados_parciais)
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
        """
        Verifica se algum browser está disponível (Edge, Chrome ou Chromium).
        Edge e Chrome são usados diretamente se instalados no Windows.
        Só tenta instalar o Chromium do Playwright como último recurso.
        """
        import platform
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            for channel in ("msedge", "chrome"):
                try:
                    b = p.chromium.launch(channel=channel, headless=True)
                    b.close()
                    self._log(f"Browser disponível: {channel}", "OK")
                    return
                except Exception:
                    pass

            # Nenhum browser do sistema — tenta instalar Chromium via driver do Playwright
            try:
                b = p.chromium.launch(headless=True)
                b.close()
                self._log("Browser disponível: chromium (playwright)", "OK")
                return
            except Exception:
                pass

        self._log("Nenhum browser encontrado. Instalando Chromium…", "WARNING")
        # Usa o driver interno do Playwright para instalar (funciona em .exe congelado)
        try:
            from playwright._impl._driver import compute_driver_executable
            driver = compute_driver_executable()
            env = {**__import__("os").environ}
            subprocess.run([str(driver), "install", "chromium"], env=env, check=True, capture_output=True)
            self._log("Chromium instalado com sucesso.", "OK")
        except Exception as exc:
            self._log(f"Falha ao instalar Chromium: {exc}", "ERROR")

    # ------------------------------------------------------------------
    # Callbacks do thread principal
    # ------------------------------------------------------------------

    def _finalizar_ok(self, resultados):
        self.btn_iniciar.config(state=tk.NORMAL)
        self.btn_conferencia.config(state=tk.NORMAL)
        self.btn_parar.config(state=tk.DISABLED, text="■   Parar")
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
        self.btn_conferencia.config(state=tk.NORMAL)
        self.btn_parar.config(state=tk.DISABLED, text="■   Parar")
        self.lbl_status.config(text=f"Erro: {msg}", fg="#c62828")
        self.lbl_resumo.config(
            text=f"Processamento encerrado com erro.\n{msg}",
            fg="#b71c1c", bg="#ffebee",
        )
        self.frame_resumo.configure(bg="#ffebee")
        self.frame_resumo.pack(fill=tk.X, padx=14, pady=(0, 12))

    def _finalizar_cancelado(self, resultados_parciais):
        self.btn_iniciar.config(state=tk.NORMAL)
        self.btn_conferencia.config(state=tk.NORMAL)
        self.btn_parar.config(state=tk.DISABLED, text="■   Parar")
        self.lbl_status.config(text="Interrompido pelo usuário.", fg="#e65100")

        linhas = ["Processamento interrompido pelo usuário.", "─" * 50]
        if resultados_parciais:
            linhas.append("Arquivos concluídos antes da parada:")
            for r in resultados_parciais:
                linhas += [
                    f"Arquivo saída : {r.arquivo_saida}",
                    f"Total LATAM   : {r.total_linhas}",
                    f"Sucesso Vendas: {r.total_sucesso}",
                    f"Erro Vendas   : {r.total_erro}",
                    "",
                ]
        else:
            linhas.append("Nenhum arquivo foi concluído antes da parada.")

        self.lbl_resumo.config(text="\n".join(linhas).strip(), fg="#e65100",
                                bg="#fff3e0")
        self.frame_resumo.configure(bg="#fff3e0")
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
