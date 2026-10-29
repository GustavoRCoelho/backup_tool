"""Testes automatizados da ferramenta de backup.

Parte A (logica pura, sem MAPI): UTF-7 modificado, LIST, FLAGS/INTERNALDATE/
  Message-ID, util, planilha, state_manager, verifier, retomada do download.
Parte B (integracao com Redemption real, SESSAO UNICA, MAIN THREAD): download
  simulado (sem rede) -> cria o .pst de verdade -> verifica -> renomeia -> limpa;
  cobre DUAS contas (LogonPstStore na 1a, AddPSTStore na 2a), dry-run, PULADA e
  --forcar.

IMPORTANTE: a Parte B roda no MAIN THREAD, igual a ferramenta real. NAO use
threads/watchdog para chamar o Redemption: MAPI/COM e STA e trava (deadlock) em
thread secundaria sem message pump. Rode este teste numa maquina recem-reiniciada
(MAPI limpo). Se o MAPI estiver travado, a Parte B fica pendurada -> reinicie.

Executar:  <python64> test_backup_tool.py   (precisa do Redemption registrado)
"""
import logging
import os
import sys
import tempfile
import shutil

PROJ = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJ)

import logger as lg
_BASE = os.path.join(tempfile.gettempdir(), "backup_tests")
if os.path.isdir(_BASE):
    shutil.rmtree(_BASE, ignore_errors=True)
lg.setup_logger(os.path.join(_BASE, "logs"), logging.WARNING)

import imap_downloader as D
import pst_builder
import state_manager as sm
import util
import verifier
import main

_FALHAS = []


def check(desc, got, exp):
    ok = got == exp
    print(("OK   " if ok else "FALHA") + " | " + desc + (f"  (got={got!r} exp={exp!r})" if not ok else ""), flush=True)
    if not ok:
        _FALHAS.append(desc)


def checktrue(desc, cond):
    check(desc, bool(cond), True)


def eml(msgid, subject="Teste"):
    return (
        f"From: Fulano <fulano@exemplo.com>\r\n"
        f"To: beltrano@exemplo.com\r\n"
        f"Subject: {subject}\r\n"
        f"Date: Thu, 05 Oct 2023 14:30:00 +0000\r\n"
        f"Message-ID: {msgid}\r\n"
        f"MIME-Version: 1.0\r\n"
        f"Content-Type: text/plain; charset=utf-8\r\n"
        f"\r\nCorpo da mensagem.\r\n"
    ).encode()


# ==========================================================================
# PARTE A - LOGICA PURA (sem MAPI)
# ==========================================================================
print("\n== A1) Logica pura ==")
check("mutf7 INBOX", D.decode_mutf7("INBOX"), "INBOX")
check("mutf7 & literal", D.decode_mutf7("A &- B"), "A & B")
check("mutf7 acento", D.decode_mutf7("Enviadas &AOk- Rascunhos"), "Enviadas é Rascunhos")
check("list barra", D._parse_list_line(rb'(\HasNoChildren) "/" "INBOX"'),
      ("\\HasNoChildren", "/", "INBOX"))
check("list ponto", D._parse_list_line(rb'(\HasNoChildren) "." "INBOX.Sent"'),
      ("\\HasNoChildren", ".", "INBOX.Sent"))
check("list literal", D._parse_list_line((rb'(\HasNoChildren) "/" {5}', b'INBOX')),
      ("\\HasNoChildren", "/", "INBOX"))
check("fs ponto", util.caminho_pasta_fs("INBOX.Sent", "."), ["INBOX", "Sent"])
check("chave canonica", util.chave_canonica("INBOX.Sent", "."), "INBOX/Sent")
check("sanitiza invalidos", util.sanitizar_componente('a:b*c?"d'), "a_b_c__d")
check("sanitiza reservado", util.sanitizar_componente("CON"), "_CON")
_cab = rb'1 (UID 7 FLAGS (\Seen \Flagged) INTERNALDATE "05-Oct-2023 14:30:00 +0000" BODY[] {3}'
check("parse flags", D.ImapDownloader._parse_flags(_cab), ["\\Seen", "\\Flagged"])
check("parse internaldate", D.ImapDownloader._parse_internaldate(_cab), "05-Oct-2023 14:30:00 +0000")
check("parse msgid", D.ImapDownloader._parse_msgid(eml("<abc@h>")), "<abc@h>")

print("\n== A2) Planilha ==")
_contas = main.validar_planilha([
    {"nome": " Joao Silva ", "email": "joao1@x.com"},
    {"nome": "Joao Silva", "email": "joao2@x.com"},
    {"nome": "Maria", "email": "maria@x.com"},
    {"nome": "Dup", "email": "joao1@x.com"},
    {"nome": "", "email": "semnome@x.com"},
    {"nome": "Ninguem", "email": "invalido"},
])
check("planilha qtd", len(_contas), 4)
check("planilha pastas unicas", len(set(c["pasta"] for c in _contas)), 4)
checktrue("planilha homonimo usa email", any("joao1@x.com" in c["pasta"] for c in _contas))

print("\n== A3) State manager ==")
_sdir = os.path.join(_BASE, "state_rt")
os.makedirs(_sdir, exist_ok=True)
_est = sm.carregar_estado(_sdir)
_est["status"] = sm.BAIXADO
_est["pastas"]["INBOX"] = {"delim": "/", "uids_baixados": ["1", "2"], "total_servidor": 2}
sm.salvar_estado(_sdir, _est)
sm.salvar_meta(_sdir, {"INBOX": {"1": {"message_id": "<a@h>", "flags": ["\\Seen"], "internaldate": None}}})
check("state round-trip status", sm.carregar_estado(_sdir)["status"], sm.BAIXADO)
check("state round-trip uids", sm.carregar_estado(_sdir)["pastas"]["INBOX"]["uids_baixados"], ["1", "2"])
check("meta round-trip msgid", sm.carregar_meta(_sdir)["INBOX"]["1"]["message_id"], "<a@h>")

print("\n== A4) Verifier ==")
_estv = {"pastas": {"INBOX": {"delim": "/", "uids_baixados": ["1", "2"]}}}
_metav = {"INBOX": {"1": {"message_id": "<a@h>"}, "2": {"message_id": "<b@h>"}}}
ok_s, _ = verifier.verificar("c@x", _estv, _metav,
                             {"INBOX": {"count": 2, "message_ids": {"<a@h>", "<b@h>"}}})
checktrue("verifier sucesso", ok_s)
ok_d, rel_d = verifier.verificar("c@x", _estv, _metav,
                                 {"INBOX": {"count": 1, "message_ids": {"<a@h>"}}})
check("verifier divergencia ok=False", ok_d, False)
check("verifier faltantes", rel_d["pastas"]["INBOX"]["faltantes"], ["<b@h>"])

print("\n== A5) Retomada do download (stub IMAP) ==")


class StubImap:
    def __init__(self, msgs):
        self.msgs = msgs
        self.fetched = []

    def select(self, mbox, readonly=False):
        return ("OK", [str(len(self.msgs)).encode()])

    def uid(self, cmd, *args):
        if cmd == "SEARCH":
            return ("OK", [b" ".join(u.encode() for u in self.msgs)])
        if cmd == "FETCH":
            uid = args[0]
            self.fetched.append(uid)
            flags, idate, raw = self.msgs[uid]
            header = (b'1 (UID ' + uid.encode() + b' FLAGS (' + flags +
                      b') INTERNALDATE "' + idate.encode() + b'" BODY[] {' +
                      str(len(raw)).encode() + b"}")
            return ("OK", [(header, raw), b")"])


_dl = D.ImapDownloader("h", 993, 30, "c@x", "pw")
_dl.imap = StubImap({
    "1": (rb"\Seen", "05-Oct-2023 14:30:00 +0000", eml("<1@h>")),
    "2": (rb"", "06-Oct-2023 10:00:00 +0000", eml("<2@h>")),
})
_rdir = os.path.join(_BASE, "resume_eml")
os.makedirs(_rdir, exist_ok=True)
# UID 1 ja baixado: o .eml precisa existir no disco (senao a auto-cura re-baixa).
with open(os.path.join(_rdir, "1.eml"), "wb") as _fh:
    _fh.write(eml("<1@h>"))
_ep = {"uids_baixados": ["1"]}
_mp = {}
_dl.baixar_pasta({"encoded": "INBOX", "decoded": "INBOX", "delim": "/", "selecionavel": True},
                 _rdir, _ep, _mp, dry_run=False)
check("retomada busca so o UID novo", _dl.imap.fetched, ["2"])
check("retomada uids finais", _ep["uids_baixados"], ["1", "2"])
checktrue("retomada gravou 2.eml", os.path.exists(os.path.join(_rdir, "2.eml")))
check("retomada meta do novo", _mp["2"]["message_id"], "<2@h>")


# ==========================================================================
# PARTE B - INTEGRACAO (Redemption real, sessao unica, MAIN THREAD)
# ==========================================================================
_cfg = {"host": "stub", "porta": 993, "timeout": 30, "lote": 50,
        "base": os.path.join(_BASE, "saida"), "pastas_ignorar": []}

# Specs por e-mail: {email: {pasta_decodificada: [(uid, message_id, flags), ...]}}
_SPECS = {
    "c1@local": {"INBOX": [("1", "<m1@h>", ["\\Seen"]), ("2", "<m2@h>", [])]},
    "c2@local": {"INBOX": [("9", "<m9@h>", [])],
                 "INBOX/Trabalho": [("5", "<m5@h>", [])]},
    "dry@local": {"INBOX": [("1", "<d1@h>", [])]},
}


def _fake_download(cfg, email, senha, dir_conta, dir_eml, estado, meta, dry_run, salvar=None, progress=None):
    resumo = {}
    for dec, msgs in _SPECS[email].items():
        ep = estado["pastas"].setdefault(dec, {})
        ep["delim"] = "/"
        ep.setdefault("uids_baixados", [])
        mp = meta.setdefault(dec, {})
        resumo[dec] = len(msgs)
        ep["total_servidor"] = len(msgs)
        if dry_run:
            continue
        partes = util.caminho_pasta_fs(dec, "/")
        d = os.path.join(dir_eml, *partes) if partes else dir_eml
        os.makedirs(d, exist_ok=True)
        for (uid, msgid, flags) in msgs:
            with open(os.path.join(d, f"{uid}.eml"), "wb") as f:
                f.write(eml(msgid))
            if uid not in ep["uids_baixados"]:
                ep["uids_baixados"].append(uid)
            mp[uid] = {"message_id": msgid, "flags": flags,
                       "internaldate": "05-Oct-2023 14:30:00 +0000"}
        if salvar:
            salvar()
    return resumo


main.imap_downloader.executar_download = _fake_download  # injeta download sem rede

_reg1 = {"nome": "Conta Um", "email": "c1@local", "pasta": "Conta Um"}
_reg2 = {"nome": "Conta Dois", "email": "c2@local", "pasta": "Conta Dois"}
_reg_dry = {"nome": "So Dry", "email": "dry@local", "pasta": "So Dry"}
_dir1 = os.path.join(_cfg["base"], "Conta Um")
_dir2 = os.path.join(_cfg["base"], "Conta Dois")
_pst1 = os.path.join(_dir1, "Conta Um.pst")
_pst2 = os.path.join(_dir2, "Conta Dois.pst")

print("\n== B) Integracao (Redemption real, sessao unica, MAIN THREAD) ==", flush=True)
print("   (se ficar pendurado aqui, o MAPI esta travado -> reinicie o Windows)", flush=True)

sessao = pst_builder.RedemptionSession()  # UMA sessao para TODAS as contas
try:
    # dry-run (nao toca no Redemption, nao cria pasta)
    check("B dry-run status", main.processar_conta(_cfg, _reg_dry, "pw", True, False, sessao), "DRY_RUN")
    checktrue("B dry-run nao cria pasta", not os.path.exists(os.path.join(_cfg["base"], "So Dry")))

    # conta 1 -> LogonPstStore
    print("   conta 1 (LogonPstStore)...", flush=True)
    check("B conta1 CONCLUIDO", main.processar_conta(_cfg, _reg1, "pw", False, False, sessao), "CONCLUIDO")
    checktrue("B conta1 pst existe", os.path.exists(_pst1))
    checktrue("B conta1 _temp limpo", not os.path.exists(os.path.join(_dir1, "_temp")))
    check("B conta1 state", sm.carregar_estado(_dir1)["status"], sm.CONCLUIDO)

    # conta 2 -> AddPSTStore (mesma sessao) - este e o caminho multi-conta corrigido
    print("   conta 2 (AddPSTStore, mesma sessao)...", flush=True)
    check("B conta2 CONCLUIDO", main.processar_conta(_cfg, _reg2, "pw", False, False, sessao), "CONCLUIDO")
    checktrue("B conta2 pst existe", os.path.exists(_pst2))

    # reabre a conta 2 e confere contagens por pasta (inclui subpasta)
    print("   reabrindo conta 2 para conferir...", flush=True)
    b = pst_builder.PstBuilder(sessao, "c2@local", "Conta Dois")
    try:
        b.abrir_store(_pst2)
        idx = b.indexar()
    finally:
        b.fechar()
    check("B conta2 INBOX count", idx.get("INBOX", {}).get("count"), 1)
    check("B conta2 subpasta count", idx.get("INBOX/Trabalho", {}).get("count"), 1)
    checktrue("B conta2 msgids", {"<m9@h>", "<m5@h>"} <= (idx["INBOX"]["message_ids"] | idx["INBOX/Trabalho"]["message_ids"]))

    # reexecucao: PULADA sem --forcar; reprocessa com --forcar
    check("B PULADA sem forcar", main.processar_conta(_cfg, _reg1, "pw", False, False, sessao), "PULADA")
    print("   conta 1 --forcar...", flush=True)
    check("B forcar reprocessa", main.processar_conta(_cfg, _reg1, "pw", False, True, sessao), "CONCLUIDO")
    checktrue("B conta1 pst apos forcar", os.path.exists(_pst1))
finally:
    sessao.encerrar()
    print("   sessao MAPI encerrada", flush=True)

print("\n========================================")
print("RESULTADO:", "TODOS OS TESTES PASSARAM" if not _FALHAS else f"{len(_FALHAS)} FALHA(S): {_FALHAS}")
print("========================================")
shutil.rmtree(_BASE, ignore_errors=True)
sys.exit(1 if _FALHAS else 0)
