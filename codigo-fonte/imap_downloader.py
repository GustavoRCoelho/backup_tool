"""Download via IMAP (RFC 3501), salvando cada mensagem como .eml.

Pontos de projeto:
- IMAP4_SSL na porta 993 (padrao explicito).
- Nomes de pasta decodificados de modified UTF-7; o delimitador de hierarquia
  e lido do proprio servidor (nao se assume "/" nem ".").
- Download com BODY.PEEK[] (nao RFC822), para NAO marcar as mensagens como
  lidas no servidor. Flags e INTERNALDATE sao capturados junto.
- Cada .eml e nomeado pelo UID (nao pelo assunto).
- Retomada: os UIDs ja baixados ficam no state.json; mensagens ja baixadas sao
  puladas. O estado e persistido periodicamente (checkpoint) durante o download.

Decisao sinalizada (ver README): a busca do corpo e feita mensagem a mensagem
por UID, em vez de um unico FETCH em lote grande. Motivo: o parsing de literais
do imaplib em respostas de FETCH multiplas e fragil, e o download 1-a-1 da
granularidade exata para a retomada. O tamanho de "lote" do config.ini e usado
como intervalo de checkpoint/persistencia de estado.
"""
import base64
import imaplib
import os
import re
import socket
import time
from email.parser import BytesParser
from email.policy import compat32 as _policy_compat

from logger import log_etapa
from progress import NullProgress
from util import caminho_pasta_fs


# --------------------------------------------------------------------------
# Modified UTF-7 (RFC 3501, secao 5.1.3) - decodificacao de nomes de pasta.
# Os nomes vem do LIST ja em modified UTF-7 (ASCII); so precisamos decodificar
# para exibir/nomear pastas. Para operacoes IMAP (SELECT) usamos o nome
# ORIGINAL (encoded) como veio do servidor.
# --------------------------------------------------------------------------
def _b64_mutf7_para_unicode(chunk):
    b = chunk.replace(",", "/")
    b += "=" * ((-len(b)) % 4)
    return base64.b64decode(b).decode("utf-16-be")


def decode_mutf7(name):
    if isinstance(name, bytes):
        name = name.decode("ascii", "replace")
    res = []
    i, n = 0, len(name)
    while i < n:
        c = name[i]
        if c == "&":
            j = name.find("-", i + 1)
            if j == -1:
                res.append(name[i:])
                break
            chunk = name[i + 1:j]
            if chunk == "":
                res.append("&")  # "&-" representa um "&" literal
            else:
                try:
                    res.append(_b64_mutf7_para_unicode(chunk))
                except Exception:
                    res.append(name[i:j + 1])
            i = j + 1
        else:
            res.append(c)
            i += 1
    return "".join(res)


# --------------------------------------------------------------------------
# Parsing de respostas do comando LIST.
# --------------------------------------------------------------------------
_LIST_RE = re.compile(rb'^\((?P<flags>[^)]*)\) (?P<delim>"[^"]*"|NIL) (?P<name>.*)$')


def _limpar_delim(d):
    if d == b"NIL":
        return None
    d = d.strip(b'"')
    return d.decode("ascii", "replace") if d else None


def _desquotar(raw):
    s = raw.decode("ascii", "replace").strip()
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        s = s[1:-1].replace('\\"', '"').replace("\\\\", "\\")
    return s


def _parse_list_line(line):
    """Aceita tanto a forma simples (bytes) quanto a forma com literal (tupla)."""
    if isinstance(line, tuple):
        prefix, name_bytes = line[0], line[1]
        m = re.match(rb'^\((?P<flags>[^)]*)\) (?P<delim>"[^"]*"|NIL) ', prefix)
        if not m:
            return None
        flags = m.group("flags").decode("ascii", "replace")
        delim = _limpar_delim(m.group("delim"))
        name = name_bytes.decode("ascii", "replace")
        return flags, delim, name
    m = _LIST_RE.match(line)
    if not m:
        return None
    flags = m.group("flags").decode("ascii", "replace")
    delim = _limpar_delim(m.group("delim"))
    name = _desquotar(m.group("name"))
    return flags, delim, name


# --------------------------------------------------------------------------
class ImapDownloader:
    def __init__(self, host, port, timeout, conta_email, senha):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.email = conta_email
        self.senha = senha
        self.imap = None
        self._mbox = None  # mailbox atualmente selecionado (para reselecionar apos reconexao)

    def conectar(self):
        self.imap = imaplib.IMAP4_SSL(self.host, self.port, timeout=self.timeout)
        self.imap.login(self.email, self.senha)

    def reconectar(self):
        try:
            if self.imap is not None:
                try:
                    self.imap.logout()
                except Exception:
                    pass
        finally:
            self.imap = None
        time.sleep(2)
        self.conectar()
        if self._mbox:
            self.imap.select(self._mbox, readonly=True)

    def fechar(self):
        if self.imap is not None:
            try:
                self.imap.logout()
            except Exception:
                pass
            self.imap = None

    def listar_pastas(self):
        typ, data = self.imap.list()
        if typ != "OK":
            raise RuntimeError(f"LIST falhou: {typ}")
        pastas = []
        for item in data:
            if item is None:
                continue
            parsed = _parse_list_line(item)
            if not parsed:
                continue
            flags, delim, encoded = parsed
            pastas.append({
                "encoded": encoded,
                "decoded": decode_mutf7(encoded),
                "delim": delim,
                "selecionavel": "\\Noselect" not in flags,
                "flags": flags,
            })
        return pastas

    def contar(self, pasta):
        """SELECT (readonly) so para obter o total de mensagens da pasta."""
        encoded = pasta["encoded"]
        mbox = '"%s"' % encoded.replace("\\", "\\\\").replace('"', '\\"')
        typ, data = self.imap.select(mbox, readonly=True)
        if typ != "OK":
            return 0
        self._mbox = mbox
        return int(data[0]) if data and data[0] else 0

    def baixar_pasta(self, pasta, dir_eml, estado_pasta, meta_pasta,
                     dry_run, checkpoint=None, intervalo=50, progress=None):
        """Baixa (ou, em dry-run, apenas conta) uma pasta. Retorna o total de
        mensagens reportado pelo servidor."""
        progress = progress or NullProgress()
        encoded = pasta["encoded"]
        mbox = '"%s"' % encoded.replace("\\", "\\\\").replace('"', '\\"')
        typ, data = self.imap.select(mbox, readonly=True)
        if typ != "OK":
            raise RuntimeError(f"SELECT falhou para '{pasta['decoded']}': {typ} {data}")
        self._mbox = mbox
        total = int(data[0]) if data and data[0] else 0
        estado_pasta["total_servidor"] = total
        if dry_run or total == 0:
            return total

        typ, sdata = self.imap.uid("SEARCH", "ALL")
        if typ != "OK":
            raise RuntimeError(f"SEARCH falhou para '{pasta['decoded']}': {typ}")
        uids = sdata[0].split() if sdata and sdata[0] else []

        baixados = set(estado_pasta.get("uids_baixados", []))
        # Auto-cura: se um .eml ja baixado sumiu do disco (ex.: removido por
        # antivirus entre o download e a montagem do PST), tira-o do estado para
        # ser baixado de novo em vez de ficar faltando no PST.
        sumidos = {u for u in baixados
                   if not os.path.exists(os.path.join(dir_eml, f"{u}.eml"))}
        if sumidos:
            baixados -= sumidos
            for u in sumidos:
                meta_pasta.pop(u, None)
            estado_pasta["uids_baixados"] = sorted(baixados, key=int)
            log_etapa(self.email, "download_pasta", "RE-BAIXAR",
                      f"pasta={pasta['decoded']} sumiram_do_disco={len(sumidos)}")
        processados = 0
        for uid_b in uids:
            uid = uid_b.decode("ascii")
            if uid in baixados:
                continue
            self._baixar_mensagem(uid, dir_eml, meta_pasta)
            baixados.add(uid)
            processados += 1
            progress.inc_download()
            if processados % intervalo == 0:
                estado_pasta["uids_baixados"] = sorted(baixados, key=int)
                if checkpoint:
                    checkpoint()
        estado_pasta["uids_baixados"] = sorted(baixados, key=int)
        if checkpoint:
            checkpoint()
        return total

    def _baixar_mensagem(self, uid, dir_eml, meta_pasta, tentativas=3):
        ult_erro = None
        for tentativa in range(1, tentativas + 1):
            try:
                typ, data = self.imap.uid(
                    "FETCH", uid, "(FLAGS INTERNALDATE BODY.PEEK[])"
                )
                if typ != "OK" or not data or data[0] is None:
                    raise RuntimeError(f"FETCH UID {uid} retornou {typ}")
                cabecalho, raw = self._extrair(data)
                if raw is None:
                    raise RuntimeError(f"FETCH UID {uid} veio sem corpo")
                caminho = os.path.join(dir_eml, f"{uid}.eml")
                with open(caminho, "wb") as f:
                    f.write(raw)
                meta_pasta[uid] = {
                    "message_id": self._parse_msgid(raw),
                    "flags": self._parse_flags(cabecalho),
                    "internaldate": self._parse_internaldate(cabecalho),
                }
                return
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, socket.error, OSError) as e:
                ult_erro = e
                if tentativa < tentativas:
                    self.reconectar()
        raise RuntimeError(f"Falha ao baixar UID {uid}: {ult_erro}")

    # ---- parsing de uma resposta FETCH ----
    @staticmethod
    def _extrair(data):
        """Extrai (cabecalho_bytes, corpo_bytes) de uma resposta de FETCH.

        Forma tipica: [ (b'<n> (UID.. FLAGS.. INTERNALDATE.. BODY[] {N}', b'<raw>'), b')' ]
        """
        cabecalho, raw = b"", None
        for parte in data:
            if isinstance(parte, tuple):
                cabecalho = parte[0]
                raw = parte[1]
        return cabecalho, raw

    @staticmethod
    def _parse_flags(cabecalho):
        m = re.search(rb"FLAGS \(([^)]*)\)", cabecalho)
        if not m:
            return []
        return m.group(1).decode("ascii", "replace").split()

    @staticmethod
    def _parse_internaldate(cabecalho):
        m = re.search(rb'INTERNALDATE "([^"]+)"', cabecalho)
        return m.group(1).decode("ascii", "replace") if m else None

    @staticmethod
    def _parse_msgid(raw):
        try:
            msg = BytesParser(policy=_policy_compat).parsebytes(raw, headersonly=True)
            mid = msg.get("Message-ID")
            if mid:
                return str(mid).strip()
        except Exception:
            pass
        return None


# --------------------------------------------------------------------------
def _filtrar_pastas(pastas, ignorar):
    if not ignorar:
        return pastas
    low = {s.lower() for s in ignorar}
    return [p for p in pastas if p["decoded"].lower() not in low]


def executar_download(cfg, conta_email, senha, dir_conta, dir_eml,
                      estado, meta, dry_run, salvar=None, progress=None):
    """Conecta, lista pastas e baixa (ou conta, em dry-run) todas as pastas
    selecionaveis. Retorna {pasta_decodificada: total_no_servidor}.

    `salvar` (callable sem args) e chamado em cada checkpoint para persistir
    estado+meta; passe None em dry-run para nao gerar efeitos colaterais.
    `progress` (ver progress.py) atualiza a barra; None = sem progresso.
    """
    progress = progress or NullProgress()
    dl = ImapDownloader(cfg["host"], cfg["porta"], cfg["timeout"], conta_email, senha)
    dl.conectar()
    try:
        pastas = _filtrar_pastas(dl.listar_pastas(), cfg["pastas_ignorar"])
        selecionaveis = [p for p in pastas if p["selecionavel"]]

        # Pre-contagem (SELECT) para dimensionar a barra de progresso da conta.
        total_conta = 0
        for p in selecionaveis:
            total_conta += dl.contar(p)
        progress.set_account_total(total_conta)
        ja = sum(len(estado["pastas"].get(p["decoded"], {}).get("uids_baixados", []))
                 for p in selecionaveis)
        progress.set_account_done_inicial(ja)

        resumo = {}
        for pasta in pastas:
            dec = pasta["decoded"]
            if not pasta["selecionavel"]:
                log_etapa(conta_email, "listar_pasta", "IGNORADA_NOSELECT", f"pasta={dec}")
                continue
            estado_pasta = estado["pastas"].setdefault(dec, {})
            estado_pasta["encoded"] = pasta["encoded"]
            estado_pasta["delim"] = pasta["delim"]
            estado_pasta.setdefault("uids_baixados", [])
            meta_pasta = meta.setdefault(dec, {})

            if not dry_run:
                partes = caminho_pasta_fs(dec, pasta["delim"])
                dir_pasta_eml = os.path.join(dir_eml, *partes) if partes else dir_eml
                os.makedirs(dir_pasta_eml, exist_ok=True)
            else:
                dir_pasta_eml = dir_eml

            progress.set_folder(dec)
            total = dl.baixar_pasta(
                pasta, dir_pasta_eml, estado_pasta, meta_pasta, dry_run,
                checkpoint=salvar, intervalo=max(1, cfg["lote"]), progress=progress,
            )
            resumo[dec] = total
            if salvar:
                salvar()
            log_etapa(
                conta_email, "download_pasta", "OK",
                f"pasta={dec} servidor={total} baixados={len(estado_pasta['uids_baixados'])}",
            )
        return resumo
    finally:
        dl.fechar()
