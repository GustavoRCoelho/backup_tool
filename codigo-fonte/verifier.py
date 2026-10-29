"""Verificacao por pasta e por Message-ID entre o baixado e o .pst.

Criterio de sucesso de uma pasta:
  - contagem no .pst == numero de UIDs baixados; E
  - nenhum Message-ID esperado faltando no .pst.

Mensagens sem Message-ID (cabecalho ausente) nao entram na comparacao por
Message-ID, mas continuam contando na contagem por pasta, que e o criterio
primario. Os Message-IDs faltantes sao registrados no log (sem assunto/conteudo).
"""
from logger import log_etapa
from util import chave_canonica


def verificar(conta_email, estado, meta, indice_pst):
    """Retorna (ok: bool, relatorio: dict)."""
    ok = True
    relatorio = {"pastas": {}, "faltantes_total": 0}

    for pasta_decoded, estado_pasta in estado["pastas"].items():
        uids = estado_pasta.get("uids_baixados", [])
        if not uids:
            continue
        delim = estado_pasta.get("delim")
        chave = chave_canonica(pasta_decoded, delim)

        meta_pasta = meta.get(pasta_decoded, {})
        esperado_ids = {m.get("message_id") for m in meta_pasta.values()
                        if m.get("message_id")}
        esperado_count = len(uids)

        info = indice_pst.get(chave, {"count": 0, "message_ids": set()})
        pst_count = info["count"]
        pst_ids = info["message_ids"]

        faltantes = sorted(esperado_ids - pst_ids)
        pasta_ok = (pst_count == esperado_count) and not faltantes
        if not pasta_ok:
            ok = False

        relatorio["pastas"][chave] = {
            "esperado": esperado_count,
            "pst": pst_count,
            "faltantes": faltantes,
        }
        relatorio["faltantes_total"] += len(faltantes)

        log_etapa(
            conta_email, "verificar_pasta",
            "OK" if pasta_ok else "DIVERGENCIA",
            f"pasta={chave} esperado={esperado_count} pst={pst_count} "
            f"faltantes={len(faltantes)}",
        )
        for mid in faltantes:
            log_etapa(conta_email, "verificar_faltante", "FALTANTE",
                      f"pasta={chave} message_id={mid}")

    return ok, relatorio
