"""Interface grafica simples (Tkinter) para a ferramenta de backup.

A GUI NAO executa o Redemption/MAPI: ela lanca o backup_tool.exe (ou main.py em
dev) como PROCESSO SEPARADO, que roda o MAPI na propria thread principal (sem
risco de travar). A GUI apenas:
  - coleta planilha, senha (mascarada) e flags;
  - envia a senha pelo stdin do processo (nunca pela linha de comando / log);
  - le o arquivo JSON de progresso e atualiza as barras (por conta e geral);
  - mostra o log em tempo real.
"""
import configparser
import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def _base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _cmd_backend():
    """Comando base para rodar o backend (CLI). Usa o .exe ao lado quando
    empacotado; em desenvolvimento usa o python + main.py."""
    exe = os.path.join(_base_dir(), "backup_tool.exe")
    if os.path.exists(exe):
        return [exe], os.path.dirname(exe)
    main_py = os.path.join(_base_dir(), "main.py")
    return [sys.executable, main_py], _base_dir()


class App:
    def __init__(self, root):
        self.root = root
        root.title("Backup de Caixas de Correio (PST)")
        root.geometry("720x560")
        self.proc = None
        self.pfile = None
        self.q = queue.Queue()
        self.fase_atual = ""
        self._build_ui()
        self._mostrar_config()

    # ---------------- UI ----------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}
        frm = ttk.LabelFrame(self.root, text="Entrada")
        frm.pack(fill="x", **pad)

        ttk.Label(frm, text="Planilha (.csv / .xlsx):").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.var_plan = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_plan).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="Procurar...", command=self._pick).grid(row=0, column=2, padx=6)

        ttk.Label(frm, text="Senha comum:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.var_pw = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_pw, show="*").grid(row=1, column=1, sticky="we", padx=6)

        ttk.Label(frm, text="Só esta conta (opcional):").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        self.var_conta = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_conta).grid(row=2, column=1, sticky="we", padx=6)

        self.var_dry = tk.BooleanVar()
        self.var_forcar = tk.BooleanVar()
        opt = ttk.Frame(frm)
        opt.grid(row=3, column=1, sticky="w", padx=6, pady=4)
        ttk.Checkbutton(opt, text="Dry-run (só conta, não baixa)", variable=self.var_dry).pack(side="left")
        ttk.Checkbutton(opt, text="Forçar reprocesso", variable=self.var_forcar).pack(side="left", padx=12)
        frm.columnconfigure(1, weight=1)

        self.lbl_cfg = ttk.Label(self.root, text="", foreground="#555")
        self.lbl_cfg.pack(anchor="w", padx=12)

        pf = ttk.LabelFrame(self.root, text="Progresso")
        pf.pack(fill="x", **pad)
        self.lbl_conta = ttk.Label(pf, text="Conta: —")
        self.lbl_conta.pack(anchor="w", padx=6, pady=(4, 0))
        self.pb_conta = ttk.Progressbar(pf, maximum=100)
        self.pb_conta.pack(fill="x", padx=6, pady=2)
        self.lbl_geral = ttk.Label(pf, text="Geral: —")
        self.lbl_geral.pack(anchor="w", padx=6, pady=(4, 0))
        self.pb_geral = ttk.Progressbar(pf, maximum=100)
        self.pb_geral.pack(fill="x", padx=6, pady=(2, 6))

        bf = ttk.Frame(self.root)
        bf.pack(fill="x", **pad)
        self.btn_start = ttk.Button(bf, text="Iniciar backup", command=self._start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(bf, text="Parar", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=8)

        lf = ttk.LabelFrame(self.root, text="Log")
        lf.pack(fill="both", expand=True, **pad)
        self.txt = tk.Text(lf, height=12, wrap="none")
        self.txt.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(lf, command=self.txt.yview)
        sb.pack(side="right", fill="y")
        self.txt.config(yscrollcommand=sb.set)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _mostrar_config(self):
        _, cfgdir = _cmd_backend()
        cfg = os.path.join(cfgdir, "config.ini")
        try:
            cp = configparser.ConfigParser(interpolation=None)
            cp.read(cfg, encoding="utf-8")
            host = cp.get("imap", "host", fallback="?")
            base = cp.get("saida", "base", fallback="?")
            self.lbl_cfg.config(text=f"Servidor: {host}   |   Pasta de saída: {base}   (config.ini)")
        except Exception:
            self.lbl_cfg.config(text="(config.ini não encontrado ao lado do executável)")

    # ---------------- acoes ----------------
    def _pick(self):
        f = filedialog.askopenfilename(
            filetypes=[("Planilhas", "*.csv *.xlsx"), ("Todos", "*.*")])
        if f:
            self.var_plan.set(f)

    def _log(self, s):
        self.txt.insert("end", s + "\n")
        self.txt.see("end")

    def _start(self):
        plan = self.var_plan.get().strip()
        pw = self.var_pw.get()
        if not plan or not os.path.exists(plan):
            messagebox.showerror("Erro", "Selecione uma planilha válida.")
            return
        if not pw:
            messagebox.showerror("Erro", "Digite a senha comum.")
            return

        self.pfile = os.path.join(tempfile.gettempdir(), f"backup_prog_{os.getpid()}.json")
        try:
            if os.path.exists(self.pfile):
                os.remove(self.pfile)
        except OSError:
            pass

        base_cmd, _ = _cmd_backend()
        cmd = base_cmd + [plan, "--progress", self.pfile, "--senha-stdin"]
        if self.var_conta.get().strip():
            cmd += ["--conta", self.var_conta.get().strip()]
        if self.var_dry.get():
            cmd += ["--dry-run"]
        if self.var_forcar.get():
            cmd += ["--forcar"]

        self.txt.delete("1.0", "end")
        self._log("Iniciando backup...")
        self.pb_conta["value"] = 0
        self.pb_geral["value"] = 0
        self.btn_start.config(state="disabled")

        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"  # stdout do backend em UTF-8
        flags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        try:
            self.proc = subprocess.Popen(
                cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                errors="replace", bufsize=1, env=env, creationflags=flags)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível iniciar o backend:\n{e}")
            self.btn_start.config(state="normal")
            return

        try:
            self.proc.stdin.write(pw + "\n")
            self.proc.stdin.flush()
            self.proc.stdin.close()
        except Exception as e:
            self._log(f"Erro ao enviar a senha: {e}")
        self.var_pw.set("")  # limpa a senha da tela

        threading.Thread(target=self._reader, daemon=True).start()
        self.root.after(300, self._poll)

    def _reader(self):
        try:
            for line in self.proc.stdout:
                self.q.put(("log", line.rstrip("\n")))
        except Exception:
            pass
        self.q.put(("end", None))

    def _poll(self):
        ended = False
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    self._log(val)
                elif kind == "end":
                    ended = True
        except queue.Empty:
            pass
        self._update_bars()
        if ended and (self.proc is None or self.proc.poll() is not None):
            self._finish()
        else:
            self.root.after(300, self._poll)

    def _update_bars(self):
        if not self.pfile or not os.path.exists(self.pfile):
            return
        try:
            with open(self.pfile, "r", encoding="utf-8") as f:
                d = json.load(f)
        except (OSError, json.JSONDecodeError, ValueError):
            return
        self.pb_conta["value"] = d.get("pct_conta", 0)
        self.pb_geral["value"] = d.get("pct_geral", 0)
        fase = d.get("phase", "")
        self.fase_atual = fase
        nome_fase = {"baixando": "Baixando", "montando": "Montando PST",
                     "verificando": "Verificando", "iniciando": "Conectando",
                     "CONCLUIDO": "Concluído", "FALHA": "Falha", "ERRO": "Erro",
                     "INTERROMPIDO": "Interrompido"}.get(fase, fase)
        det = ""
        if fase == "baixando":
            det = f"  pasta: {d.get('folder','')}  ({d.get('account_done',0)}/{d.get('account_total',0)})"
        elif fase == "montando":
            det = f"  ({d.get('build_done',0)}/{d.get('build_total',0)})"
        self.lbl_conta.config(
            text=f"Conta {d.get('index',0)}/{d.get('total_accounts',0)}: "
                 f"{d.get('email','')}  —  {nome_fase}{det}  —  {d.get('pct_conta',0)}%")
        self.lbl_geral.config(text=f"Geral: {d.get('pct_geral',0)}%")
        # Parar so e seguro na fase de download (matar durante montagem trava o MAPI).
        if self.proc and self.proc.poll() is None and fase in ("baixando", "iniciando"):
            self.btn_stop.config(state="normal")
        else:
            self.btn_stop.config(state="disabled")

    def _finish(self):
        self._update_bars()
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        rc = self.proc.poll() if self.proc else None
        resultados = {}
        try:
            with open(self.pfile, "r", encoding="utf-8") as f:
                resultados = json.load(f).get("resultados", {})
        except Exception:
            pass
        self._log("")
        self._log(f"--- Finalizado (código {rc}) ---")
        if resultados:
            self._log("Resumo:")
            for k, v in resultados.items():
                self._log(f"   {k}  ->  {v}")
            if all(v == "CONCLUIDO" for v in resultados.values()):
                messagebox.showinfo("Backup", "Backup concluído com sucesso!")
            else:
                messagebox.showwarning(
                    "Backup", "Backup finalizado com avisos. Veja o log e o resumo.")

    def _stop(self):
        if self.proc and self.proc.poll() is None:
            if self.fase_atual not in ("baixando", "iniciando"):
                messagebox.showwarning(
                    "Parar", "Não é seguro parar durante a montagem/verificação do "
                    "PST (pode travar o MAPI). Aguarde esta etapa terminar.")
                return
            if messagebox.askyesno("Parar", "Parar o backup agora?\n"
                                   "O progresso é retomável (rode de novo sem 'Forçar')."):
                try:
                    self.proc.terminate()
                except Exception:
                    pass
                self._log("Backup interrompido pelo usuário.")

    def _on_close(self):
        if self.proc and self.proc.poll() is None:
            if not messagebox.askyesno(
                    "Sair", "O backup ainda está rodando. Ele continuará em segundo "
                    "plano se você fechar.\nFechar a janela mesmo assim?"):
                return
        self.root.destroy()


def run():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(run())
