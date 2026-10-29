"""Estado por conta (state.json) e metadados (meta.json) para retomada.

state.json  -> enxuto: status geral + por pasta os UIDs ja baixados e totais.
meta.json   -> por pasta: uid -> {message_id, flags, internaldate}. Fica
               separado do state.json para nao incha-lo, mas e necessario para
               a importacao (preservar lido/nao-lido e data) e a verificacao
               (comparacao por Message-ID).

Nenhum dos dois arquivos guarda conteudo/assunto de mensagem.
"""
import json
import os
import tempfile

# Status possiveis por conta (gravados em state.json["status"]).
PENDENTE = "PENDENTE"
BAIXANDO = "BAIXANDO"
BAIXADO = "BAIXADO"
CONSTRUINDO_PST = "CONSTRUINDO_PST"
VERIFICANDO = "VERIFICANDO"
CONCLUIDO = "CONCLUIDO"
FALHA = "FALHA"


def _state_path(dir_conta):
    return os.path.join(dir_conta, "state.json")


def _meta_path(dir_conta):
    return os.path.join(dir_conta, "meta.json")


def _ler_json(caminho, padrao):
    if os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return padrao


def _gravar_json_atomico(dir_conta, caminho, dados):
    """Escrita atomica: grava em arquivo temporario e substitui (os.replace)."""
    os.makedirs(dir_conta, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=dir_conta, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        os.replace(tmp, caminho)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def carregar_estado(dir_conta):
    return _ler_json(
        _state_path(dir_conta),
        {
            "status": PENDENTE,
            "email": None,
            "nome": None,
            # nome_pasta_decodificado -> {encoded, delim, uids_baixados[], total_servidor}
            "pastas": {},
        },
    )


def salvar_estado(dir_conta, estado):
    _gravar_json_atomico(dir_conta, _state_path(dir_conta), estado)


def marcar_status(dir_conta, estado, status):
    estado["status"] = status
    salvar_estado(dir_conta, estado)


def carregar_meta(dir_conta):
    return _ler_json(_meta_path(dir_conta), {})


def salvar_meta(dir_conta, meta):
    _gravar_json_atomico(dir_conta, _meta_path(dir_conta), meta)
