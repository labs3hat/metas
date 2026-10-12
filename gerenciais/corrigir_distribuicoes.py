"""Aplica corrige_distribuicoes ao "Resumo Distribuições Lucros" direto no Drive.

Variáveis de ambiente: GOOGLE_CREDENTIALS, GERENCIAIS_PASTA_ID (mesma pasta das gerenciais),
DISTRIBUICOES_NOME (nome exato do arquivo), SIMULAR ("true" padrão: só mostra; "false" grava).
Sai com 1 em qualquer falha. A revisão anterior fica preservada em Gerenciar versões.
"""
from __future__ import annotations

import logging
import os
import sys
import tempfile
from pathlib import Path

import corrige_distribuicoes as cd
import drive_gerenciais as dg

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
log = logging.getLogger("gerenciais.distribuicoes")


def main() -> int:
    pasta = os.environ["GERENCIAIS_PASTA_ID"]
    nome = os.environ["DISTRIBUICOES_NOME"]
    simular = os.environ.get("SIMULAR", "true").strip().lower() != "false"
    log.info("=== Resumo Distribuições — %s ===", "SIMULAÇÃO" if simular else "GRAVANDO NO DRIVE")
    s = dg.sessao()
    arq = dg.buscar_por_nome(s, pasta, nome)
    with tempfile.TemporaryDirectory() as d:
        origem = dg.baixar(s, arq, Path(d) / "orig.xlsx")
        destino = Path(d) / "novo.xlsx"
        mud = cd.corrige(str(origem), str(destino))
        cd.verificar(str(origem), str(destino))
        for m in mud:
            log.info("    %s", m)
        if simular:
            log.info("SIMULAÇÃO — %d mudança(s), nada gravado", len(mud))
        else:
            dg.subir_nova_versao(s, arq, destino, mime=dg.MIME_XLSX)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # falha visível no workflow
        log.error("FALHOU: %s: %s", type(exc).__name__, exc)
        sys.exit(1)
