"""Relato de progresso para a interface grafica.

O backend (CLI) grava um pequeno arquivo JSON (--progress <arquivo>) que a GUI
le periodicamente para atualizar as barras. Sem conteudo de mensagem, so numeros.

Percentual por conta (pct_conta) combina as fases:
  baixando  -> 0% a 60% (account_done/account_total)
  montando  -> 60% a 90% (build_done/build_total)
  verificando -> ~92%
  CONCLUIDO -> 100%
Percentual geral (pct_geral) = (contas_concluidas + fracao_da_atual)/total_contas.
"""
import json
import os
import time


class Progress:
    def __init__(self, caminho, total_accounts):
        self.caminho = caminho
        self.d = {
            "total_accounts": total_accounts, "index": 0, "email": "",
            "phase": "", "folder": "",
            "account_total": 0, "account_done": 0,
            "build_total": 0, "build_done": 0,
            "pct_conta": 0.0, "pct_geral": 0.0,
            "status": "running", "resultados": {},
        }
        self._last = 0.0
        self.escrever(force=True)

    def start_account(self, index, email):
        self.d.update(index=index, email=email, phase="iniciando", folder="",
                      account_total=0, account_done=0, build_total=0, build_done=0,
                      pct_conta=0.0)
        self._calc(); self.escrever(force=True)

    def set_account_total(self, n):
        self.d["account_total"] = n; self._calc(); self.escrever(force=True)

    def set_account_done_inicial(self, n):
        self.d["account_done"] = n; self._calc(); self.escrever(force=True)

    def set_folder(self, nome):
        self.d["folder"] = nome; self.d["phase"] = "baixando"; self._calc(); self.escrever(force=True)

    def inc_download(self, k=1):
        self.d["account_done"] += k; self._calc(); self.escrever()

    def start_build(self, total):
        self.d.update(phase="montando", build_total=total, build_done=0)
        self._calc(); self.escrever(force=True)

    def inc_build(self, k=1):
        self.d["build_done"] += k; self._calc(); self.escrever()

    def set_phase(self, phase):
        self.d["phase"] = phase; self._calc(); self.escrever(force=True)

    def finish_account(self, status):
        self.d["resultados"][self.d["email"]] = status
        self.d["phase"] = status
        if status == "CONCLUIDO":
            self.d["pct_conta"] = 100.0
        self._calc(); self.escrever(force=True)

    def finish_all(self):
        self.d["status"] = "done"; self.escrever(force=True)

    def _calc(self):
        at = self.d["account_total"] or 0
        phase = self.d["phase"]
        if phase == "CONCLUIDO":
            pc = 100.0
        elif phase in ("baixando", "iniciando"):
            pc = (min(self.d["account_done"] / at, 1.0) if at else 0.0) * 60.0
        elif phase == "montando":
            bt = self.d["build_total"] or 0
            pc = 60.0 + (min(self.d["build_done"] / bt, 1.0) if bt else 0.0) * 30.0
        elif phase == "verificando":
            pc = 92.0
        else:  # FALHA/ERRO/outros: mantem o valor atual
            pc = self.d["pct_conta"]
        self.d["pct_conta"] = round(pc, 1)
        ta = self.d["total_accounts"] or 1
        idx = max(self.d["index"] - 1, 0)
        self.d["pct_geral"] = round(min((idx + pc / 100.0) / ta, 1.0) * 100.0, 1)

    def escrever(self, force=False):
        now = time.time()
        if not force and (now - self._last) < 0.2:
            return
        self._last = now
        try:
            tmp = self.caminho + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.d, f, ensure_ascii=False)
            os.replace(tmp, self.caminho)
        except OSError:
            pass


class NullProgress:
    """Usado quando nao ha --progress (execucao CLI pura). Ignora tudo."""
    def __getattr__(self, name):
        def _noop(*a, **k):
            return None
        return _noop
