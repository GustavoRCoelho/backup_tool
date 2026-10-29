"""Funcoes utilitarias compartilhadas.

(Modulo adicionado alem da estrutura do prompt para evitar duplicacao de
helpers de sanitizacao de nomes entre o downloader, o construtor do PST e o
verificador.)
"""
import re

# Caracteres invalidos para nome de arquivo/pasta no Windows.
_INVALIDOS = '<>:"/\\|?*'
_TRANS = {ord(c): '_' for c in _INVALIDOS}
# Caracteres de controle (0-31) tambem sao invalidos.
_TRANS.update({c: '_' for c in range(32)})

_RESERVADOS = (
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)


def sanitizar_componente(nome):
    """Sanitiza um unico componente de caminho (sem separadores) para Windows."""
    nome = (nome or "").strip().translate(_TRANS)
    # Windows nao permite terminar com espaco ou ponto.
    nome = nome.rstrip(" .")
    if nome.upper() in _RESERVADOS:
        nome = "_" + nome
    return nome or "_"


def caminho_pasta_fs(decoded, delim):
    """Converte o nome (decodificado) de uma pasta IMAP em uma lista de
    componentes de caminho seguros para o sistema de arquivos, quebrando pela
    hierarquia reportada pelo servidor (delimitador)."""
    partes = decoded.split(delim) if delim else [decoded]
    return [sanitizar_componente(p) for p in partes if p != ""]


def _limpar_nome_pasta_pst(nome):
    """Limpa um nome de pasta para uso dentro do .pst (apenas barras sao
    problematicas; acentos e espacos sao permitidos)."""
    return (nome or "").replace("\\", "-").replace("/", "-").strip() or "_"


def partes_pasta_pst(decoded, delim):
    """Componentes de pasta para recriar a hierarquia dentro do .pst,
    preservando acentos e nomes originais do servidor."""
    partes = decoded.split(delim) if delim else [decoded]
    return [_limpar_nome_pasta_pst(p) for p in partes if p != ""]


def chave_canonica(decoded, delim):
    """Chave estavel para casar uma pasta IMAP com a pasta correspondente no
    .pst, usada na verificacao. Deriva do mesmo nome de exibicao usado para
    criar as pastas no PST."""
    return "/".join(partes_pasta_pst(decoded, delim))
