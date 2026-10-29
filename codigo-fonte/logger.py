"""Log estruturado por conta e por etapa.

Regra de LGPD (critica): este log NUNCA deve registrar assunto, remetente,
destinatario ou conteudo das mensagens, nem a senha. Apenas: conta, horario,
etapa, status, contagens por pasta e Message-IDs faltantes.
"""
import logging
import os
from datetime import datetime

_LOGGER_NAME = "backup_tool"


def setup_logger(log_dir, nivel=logging.INFO):
    """Configura o logger (arquivo + console). Retorna o logger."""
    os.makedirs(log_dir, exist_ok=True)
    nome_arquivo = datetime.now().strftime("backup_%Y%m%d_%H%M%S.log")
    caminho = os.path.join(log_dir, nome_arquivo)

    logger = logging.getLogger(_LOGGER_NAME)
    logger.setLevel(nivel)
    logger.handlers.clear()
    logger.propagate = False

    fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = logging.FileHandler(caminho, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    logger.info("Log iniciado em %s", caminho)
    return logger


def get_logger():
    return logging.getLogger(_LOGGER_NAME)


def log_etapa(conta, etapa, status, detalhe=""):
    """Registra uma etapa do pipeline de uma conta.

    IMPORTANTE: nunca passe assunto/conteudo de mensagem em `detalhe`.
    """
    logger = get_logger()
    msg = f"conta={conta} | etapa={etapa} | status={status}"
    if detalhe:
        msg += f" | {detalhe}"
    logger.info(msg)
