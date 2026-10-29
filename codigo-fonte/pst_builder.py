"""Criacao do .pst e importacao dos .eml via Redemption (RDO).

API do Redemption usada (confirmada na documentacao oficial - dimastr.com):

  - RDOSession.LogonPstStore(Path, Format, DisplayName) -> RDOPstStore
      Cria (se nao existir) e abre o .pst. Format=2 (olStoreUnicode) evita o
      limite de 2 GB do formato ANSI. IMPORTANTE: LogonPstStore cria um profile
      MAPI *temporario*, removido logo em seguida -> NAO exige um profile do
      Outlook pre-configurado. (O prompt citou AddStoreEx, do Object Model do
      Outlook; no Redemption, LogonPstStore e o equivalente direto e dispensa
      profile, por isso foi o escolhido.)

  - RDOFolder.Items.Add("IPM.Note") -> RDOMail (item vazio)
  - RDOMail.Import(Path, 1024) -> importa um .eml (1024 = olRFC822)
  - RDOMail.ReceivedTime / RDOMail.UnRead -> propriedades gravaveis
  - RDOStore.Remove() -> desmonta UM store da sessao (sem derrubar a sessao)
  - RDOSession.Logoff() -> encerra a sessao MAPI (uma vez, no fim do processo)

IMPORTANTE (sessao unica): o MAPI nao gosta de ser reinicializado varias vezes
no mesmo processo. Fazer CoInitialize/LogonPstStore/Logoff/CoUninitialize em
loop (uma vez por conta) pode TRAVAR o subsistema MAPI ao processar varias
contas. Por isso ha uma unica RedemptionSession por processo: um Logon e um
Logoff, com AddPSTStore/Remove por conta.

PST travado apos o fechamento (comportamento conhecido do Outlook): a renomeacao
final do arquivo e feita com varias tentativas (ver renomear_com_retentativa).
"""
import gc
import os
import re
import time
from datetime import datetime, timedelta, timezone

import pythoncom
import pywintypes
import win32com.client

from logger import get_logger, log_etapa
from progress import NullProgress
from util import partes_pasta_pst

# Constantes do Redemption / Outlook (confirmadas na doc).
OL_STORE_UNICODE = 2       # rdoStoreType.olStoreUnicode
OL_RFC822 = 1024           # Import: arquivo .eml (RFC822)
PR_INTERNET_MESSAGE_ID_W = 0x1035001F
PR_INTERNET_MESSAGE_ID_A = 0x1035001E

_MESES = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}
_IDATE_RE = re.compile(
    r"(\d{1,2})-([A-Za-z]{3})-(\d{4}) (\d{2}):(\d{2}):(\d{2}) ([+-]\d{4})")


def _internaldate_to_pytime(idate):
    """Converte o INTERNALDATE do IMAP (ex.: '05-Oct-2023 14:30:00 +0000')
    em um pywintypes.Time (hora local, sem tz) para gravar em ReceivedTime.
    Parsing independente de locale (os meses do IMAP sao sempre em ingles)."""
    m = _IDATE_RE.search(idate or "")
    if not m:
        return None
    dia, mes, ano, hh, mm, ss, tz = m.groups()
    mnum = _MESES.get(mes.title())
    if not mnum:
        return None
    try:
        sinal = 1 if tz[0] == "+" else -1
        off = timezone(sinal * timedelta(hours=int(tz[1:3]), minutes=int(tz[3:5])))
        dt = datetime(int(ano), mnum, int(dia), int(hh), int(mm), int(ss), tzinfo=off)
        dt_local = dt.astimezone().replace(tzinfo=None)
        return pywintypes.Time(dt_local)
    except Exception:
        return None


def remover_com_retentativa(caminho, tentativas=10, espera=2):
    """Remove um arquivo tolerando travas temporarias (PST montado)."""
    for i in range(1, tentativas + 1):
        try:
            if os.path.exists(caminho):
                os.remove(caminho)
            return
        except PermissionError:
            if i == tentativas:
                raise
            time.sleep(espera)


def renomear_com_retentativa(origem, destino, tentativas=10, espera=2):
    """Renomeia tolerando travas temporarias do Outlook/Redemption sobre o .pst."""
    ult = None
    for i in range(1, tentativas + 1):
        try:
            if os.path.exists(destino):
                os.remove(destino)
            os.replace(origem, destino)
            return
        except PermissionError as e:
            ult = e
            time.sleep(espera)
    raise RuntimeError(
        f"Nao foi possivel renomear o .pst (arquivo ainda travado?): "
        f"{origem} -> {destino}: {ult}")


class RedemptionSession:
    """Mantem UMA sessao MAPI para todo o processo.

    Reinicializar o MAPI a cada conta (CoInitialize/LogonPstStore/Logoff/
    CoUninitialize em loop) pode TRAVAR o subsistema MAPI. Aqui ha um unico
    Logon (via LogonPstStore na 1a conta) e um unico Logoff; cada conta apenas
    monta (AddPSTStore) e desmonta (store.Remove) o seu .pst.
    """

    def __init__(self):
        self.session = None
        self._com_init = False
        self._logged = False

    def _garantir(self):
        if self.session is None:
            pythoncom.CoInitialize()
            self._com_init = True
            self.session = win32com.client.Dispatch("Redemption.RDOSession")

    def abrir_pst(self, pst_path, nome):
        self._garantir()
        os.makedirs(os.path.dirname(pst_path), exist_ok=True)
        disp = f"Backup {nome}"
        if not self._logged:
            # 1a conta: LogonPstStore faz o logon (profile temporario) e abre o store.
            store = self.session.LogonPstStore(pst_path, OL_STORE_UNICODE, disp)
            self._logged = True
        else:
            # Demais contas: mesma sessao, apenas monta mais um store Unicode.
            store = self.session.Stores.AddPSTStore(pst_path, OL_STORE_UNICODE, disp)
        return store

    def fechar_pst(self, store):
        if store is not None:
            try:
                store.Remove()   # desmonta o store, mantendo a sessao viva
            except Exception as e:
                get_logger().warning("store.Remove falhou: %s", e)

    def encerrar(self):
        """Logoff + CoUninitialize UMA vez, no fim do processo."""
        if self.session is not None:
            try:
                self.session.Logoff()
            except Exception as e:
                get_logger().warning("Logoff do Redemption falhou: %s", e)
        self.session = None
        gc.collect()
        if self._com_init:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_init = False
        self._logged = False


class PstBuilder:
    def __init__(self, sessao, conta_email, nome):
        self.sessao = sessao
        self.conta_email = conta_email
        self.nome = nome
        self.store = None

    def abrir_store(self, pst_path):
        self.store = self.sessao.abrir_pst(pst_path, self.nome)
        return self.store

    def importar(self, dir_eml, estado, meta, progress=None):
        """Recria a hierarquia de pastas no .pst e importa cada .eml.
        Retorna o total de mensagens importadas."""
        progress = progress or NullProgress()
        raiz = self.store.IPMRootFolder
        total = 0
        for pasta_decoded, estado_pasta in estado["pastas"].items():
            uids = estado_pasta.get("uids_baixados", [])
            if not uids:
                continue
            delim = estado_pasta.get("delim")
            partes = partes_pasta_pst(pasta_decoded, delim)
            folder = self._get_or_create_folder(raiz, partes)
            dir_pasta = os.path.join(dir_eml, *self._partes_fs(pasta_decoded, delim))
            meta_pasta = meta.get(pasta_decoded, {})
            for uid in uids:
                caminho = os.path.join(dir_pasta, f"{uid}.eml")
                if not os.path.exists(caminho):
                    log_etapa(self.conta_email, "importar", "EML_AUSENTE",
                              f"pasta={pasta_decoded} uid={uid}")
                    continue
                self._importar_um(folder, caminho, meta_pasta.get(uid, {}))
                total += 1
                progress.inc_build()
        return total

    @staticmethod
    def _partes_fs(decoded, delim):
        # Mesma logica do downloader para localizar os .eml no disco.
        from util import caminho_pasta_fs
        return caminho_pasta_fs(decoded, delim)

    def _get_or_create_folder(self, raiz, partes):
        atual = raiz
        for parte in partes:
            atual = self._filho(atual, parte)
        return atual

    @staticmethod
    def _filho(pai, nome):
        for f in pai.Folders:
            try:
                if f.Name == nome:
                    return f
            except Exception:
                continue
        return pai.Folders.Add(nome)

    def _importar_um(self, folder, caminho, meta_uid):
        msg = folder.Items.Add("IPM.Note")
        msg.Import(caminho, OL_RFC822)
        msg.Save()
        # Preserva lido/nao-lido a partir das flags IMAP.
        try:
            msg.UnRead = ("\\Seen" not in meta_uid.get("flags", []))
        except Exception:
            pass
        # Preserva a data de recebimento (INTERNALDATE) quando possivel.
        pyt = _internaldate_to_pytime(meta_uid.get("internaldate"))
        if pyt is not None:
            try:
                msg.ReceivedTime = pyt
            except Exception:
                pass
        try:
            msg.Save()
        except Exception:
            pass

    def indexar(self):
        """Percorre o .pst e retorna {chave_pasta: {count, message_ids:set()}}
        para a verificacao. A chave usa os nomes de pasta exatamente como foram
        criados (ver util.chave_canonica)."""
        indice = {}
        self._indexar_rec(self.store.IPMRootFolder, [], indice)
        return indice

    def _indexar_rec(self, folder, prefixo, indice):
        for sub in folder.Folders:
            caminho = prefixo + [sub.Name]
            chave = "/".join(caminho)
            itens = sub.Items
            count = itens.Count
            msgids = set()
            for i in range(1, count + 1):
                mid = self._msgid(itens.Item(i))
                if mid:
                    msgids.add(mid)
            indice[chave] = {"count": count, "message_ids": msgids}
            self._indexar_rec(sub, caminho, indice)

    @staticmethod
    def _msgid(item):
        for tag in (PR_INTERNET_MESSAGE_ID_W, PR_INTERNET_MESSAGE_ID_A):
            try:
                v = item.Fields(tag)
                if v:
                    return str(v).strip()
            except Exception:
                continue
        return None

    def fechar(self):
        """Desmonta o store desta conta (a sessao segue viva para a proxima)."""
        self.sessao.fechar_pst(self.store)
        self.store = None
