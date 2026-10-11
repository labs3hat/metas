"""Corrige as planilhas gerenciais das lojas direto no Drive.

Para cada "Gestao Financeira <loja> 2026.xlsm" da pasta (só lojas conhecidas):
baixa → corrige (corrige_planilha.corrige) → verifica (corrige_planilha.verificar)
→ se mudou algo e não for simulação, sobe como nova versão do mesmo arquivo.

Variáveis de ambiente:
  GOOGLE_CREDENTIALS   JSON da conta de serviço (a mesma do scraper)
  GERENCIAIS_PASTA_ID  id da pasta do Drive com as planilhas
  LOJAS                "todas" ou lista "sjp1,ctba3" (padrão: todas)
  SIMULAR              "true" (padrão) só mostra o que mudaria; "false" grava
  RECEITA_LOJAS        lojas que recebem a receita do BIP360 na linha 101 da DRE
                       ("" = nenhuma, "todas", ou lista). Lê a planilha de indicadores.
  RECEITA_MESES        meses a gravar, "1-9" (padrão) ou "7,8,9"

Sai com código 1 se qualquer loja falhar, para o workflow ficar vermelho.
"""
from __future__ import annotations

import logging
import os
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import corrige_planilha as cp
import drive_gerenciais as dg

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s",
                    datefmt="%Y-%m-%d %H:%M:%S")
log = logging.getLogger("gerenciais.lote")


@dataclass
class Resultado:
    loja: str
    arquivo: str
    mudancas: list[str] = field(default_factory=list)
    gravado: bool = False
    erro: str | None = None


def _conjunto(var: str, padrao: str) -> set[str] | None:
    """"" → conjunto vazio; "todas" → None (= todas); lista → conjunto validado."""
    txt = os.environ.get(var, padrao).strip().lower()
    if txt == "":
        return set()
    if txt == "todas":
        return None
    pedidas = {l.strip() for l in txt.split(",") if l.strip()}
    desconhecidas = pedidas - set(cp.EMPRESAS)
    if desconhecidas:
        raise SystemExit(f"{var}: loja(s) desconhecida(s) {sorted(desconhecidas)}; use {sorted(cp.EMPRESAS)}")
    return pedidas


def lojas_pedidas() -> set[str] | None:
    return _conjunto("LOJAS", "todas") or None


def receitas_pedidas() -> dict[str, dict[int, float]]:
    """Receita do BIP360 por loja, para as lojas em RECEITA_LOJAS (vazio = nenhuma)."""
    import receita_indicadores as ri

    conj = _conjunto("RECEITA_LOJAS", "")
    lojas = sorted(cp.EMPRESAS) if conj is None else sorted(conj)
    if not lojas:
        return {}
    meses = ri.meses_do_intervalo(os.environ.get("RECEITA_MESES", "1-9"))
    log.info("Lendo receita do BIP360 (planilha de indicadores) para %s, meses %s", lojas, meses)
    return ri.ler_receitas(ri.abrir_aba(), lojas, meses)


def processar(s, arquivo: dg.ArquivoDrive, loja: str, simular: bool, tmp: Path,
              receita: dict[int, float] | None) -> Resultado:
    res = Resultado(loja=loja, arquivo=arquivo.nome)
    try:
        origem = dg.baixar(s, arquivo, tmp / f"{loja}.orig.xlsm")
        destino = tmp / f"{loja}.novo.xlsm"
        res.mudancas = cp.corrige(str(origem), str(destino), loja, receita)
        cp.verificar(str(origem), str(destino), loja, receita)
        for m in res.mudancas:
            log.info("    %s", m)
        if not res.mudancas:
            log.info("  %s: nada a corrigir", arquivo.nome)
        elif simular:
            log.info("  %s: SIMULAÇÃO — %d mudança(s), nada gravado", arquivo.nome, len(res.mudancas))
        else:
            dg.subir_nova_versao(s, arquivo, destino)
            res.gravado = True
    except Exception as exc:  # erro de uma loja não derruba as outras; vira falha no fim
        res.erro = f"{type(exc).__name__}: {exc}"
        log.error("  %s: %s", arquivo.nome, res.erro)
    return res


def main() -> int:
    pasta = os.environ["GERENCIAIS_PASTA_ID"]
    simular = os.environ.get("SIMULAR", "true").strip().lower() != "false"
    pedidas = lojas_pedidas()
    log.info("=== Planilhas gerenciais — %s ===", "SIMULAÇÃO" if simular else "GRAVANDO NO DRIVE")

    receitas = receitas_pedidas()
    s = dg.sessao()
    arquivos = dg.listar_planilhas(s, pasta)
    resultados: list[Resultado] = []
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        for arq in arquivos:
            loja = cp.loja_do_arquivo(arq.nome)
            if loja is None:
                log.info("  ignorado (loja não cadastrada): %s", arq.nome)
                continue
            if pedidas is not None and loja not in pedidas:
                continue
            log.info("--- %s (%s) ---", arq.nome, loja)
            resultados.append(processar(s, arq, loja, simular, tmp, receitas.get(loja)))

    faltando = (pedidas or set(cp.EMPRESAS)) - {r.loja for r in resultados}
    if faltando:
        log.error("Não encontrei na pasta a planilha de: %s", sorted(faltando))

    com_erro = [r for r in resultados if r.erro]
    log.info("=== Resumo: %d planilha(s) | %d com mudança | %d gravada(s) | %d com erro | %d não encontrada(s) ===",
             len(resultados), sum(1 for r in resultados if r.mudancas),
             sum(1 for r in resultados if r.gravado), len(com_erro), len(faltando))
    for r in com_erro:
        log.error("  FALHOU %s: %s", r.arquivo, r.erro)
    return 1 if (com_erro or faltando) else 0


if __name__ == "__main__":
    sys.exit(main())
