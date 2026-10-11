"""Receita mensal de cada loja na planilha de indicadores (a que o scraper do
BIP360 alimenta): é a fonte da receita que vai para a linha 101 das DREs.

Aba 'Indicadores de Venda 2026': bloco de 36 linhas por mês a partir da linha 4
(SJP 1), coluna AH = ACUM. MÊS. Ordem das lojas no bloco (mesma de
indicadores_scraper.LOJA_ORDER): sjp1, ctba3, ctba5, cl2, ctba7, ctba9(=mga8),
ctba11, mga3, mga5, mga7.
"""
from __future__ import annotations

import json
import logging
import os

import gspread
from google.oauth2.service_account import Credentials

log = logging.getLogger("gerenciais.receita")

SHEET_ID = "1qU8Ny_OqoF4VrI0IU4JuuRvoNnmBOMSf9JkFs1h4PRY"
ABA = "Indicadores de Venda 2026"
LINHA_BASE = 4
LINHAS_POR_MES = 36
COLUNA_ACUMULADO = "AH"
POSICAO = {"sjp1": 0, "ctba3": 1, "ctba5": 2, "cl2": 3, "ctba7": 4, "mga8": 5,
           "ctba11": 6, "mga3": 7, "mga5": 8, "mga7": 9}


def meses_do_intervalo(txt: str) -> list[int]:
    """'1-9' → [1..9]; '7,8,9' → [7, 8, 9]."""
    txt = txt.strip()
    if "-" in txt:
        a, b = (int(p) for p in txt.split("-", 1))
        meses = list(range(a, b + 1))
    else:
        meses = [int(p) for p in txt.split(",") if p.strip()]
    fora = [m for m in meses if not 1 <= m <= 12]
    if fora:
        raise ValueError(f"RECEITA_MESES: mês inválido {fora}")
    return meses


def abrir_aba() -> gspread.Worksheet:
    creds = Credentials.from_service_account_info(
        json.loads(os.environ["GOOGLE_CREDENTIALS"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return gspread.authorize(creds).open_by_key(SHEET_ID).worksheet(ABA)


def ler_receitas(ws: gspread.Worksheet, lojas: list[str], meses: list[int]) -> dict[str, dict[int, float]]:
    """{loja: {mês: receita}} — só meses com valor numérico maior que zero."""
    ranges = [f"{COLUNA_ACUMULADO}{LINHA_BASE + LINHAS_POR_MES * (m - 1) + POSICAO[l]}" for l in lojas for m in meses]
    blocos = ws.batch_get(ranges, value_render_option="UNFORMATTED_VALUE")
    out: dict[str, dict[int, float]] = {l: {} for l in lojas}
    i = 0
    for l in lojas:
        for m in meses:
            bloco = blocos[i]; i += 1
            v = bloco[0][0] if bloco and bloco[0] else None
            if isinstance(v, (int, float)) and v > 0:
                out[l][m] = round(float(v), 2)
            else:
                log.warning("  receita de %s no mês %d está vazia na planilha de indicadores (%r); não será gravada", l, m, v)
    for l in lojas:
        log.info("  receita BIP360 de %s: %s", l, {m: f"{v:,.2f}" for m, v in out[l].items()})
    return out
