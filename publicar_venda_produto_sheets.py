"""
Publica o resultado da extração avulsa de venda por produto numa planilha do
Google (Drive do André), com a conta de serviço que o robô já usa. Não grava
no banco. Não imprime dado de venda no log (o repositório é público).

Uso: python publicar_venda_produto_sheets.py <pasta do artifact> <id da planilha>
"""

import csv
import json
import os
import sys

import gspread
from google.oauth2.service_account import Credentials


def ler(caminho: str) -> list[list[str]]:
    with open(caminho, newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def numero(texto: str) -> float:
    return float(texto) if texto not in ("", None) else 0.0


def gravar_aba(planilha, titulo: str, linhas: list[list]) -> None:
    try:
        aba = planilha.worksheet(titulo)
        aba.clear()
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title=titulo, rows=1, cols=1)
    aba.resize(rows=max(len(linhas), 1), cols=max(len(linhas[0]), 1))
    # Em blocos, para não estourar o tamanho de uma requisição.
    for ini in range(0, len(linhas), 20000):
        aba.update(values=linhas[ini:ini + 20000], range_name=f"A{ini + 1}",
                   value_input_option="RAW")


def main() -> int:
    pasta, planilha_id = sys.argv[1], sys.argv[2]
    venda = ler(os.path.join(pasta, "venda_produto_por_loja_mes.csv"))
    situacao = ler(os.path.join(pasta, "venda_produto_por_loja_mes_situacao.csv"))
    cab, corpo = venda[0], venda[1:]
    linhas_venda = [cab] + [[l[0], l[1], l[2], numero(l[3]), numero(l[4])] for l in corpo]
    linhas_sit = [situacao[0]] + [[l[0], l[1], l[2], int(l[3] or 0), l[4]] for l in situacao[1:]]

    cred = Credentials.from_service_account_info(
        json.loads(os.environ["GOOGLE_CREDENTIALS"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    planilha = gspread.authorize(cred).open_by_key(planilha_id)
    gravar_aba(planilha, "venda", linhas_venda)
    gravar_aba(planilha, "situacao", linhas_sit)
    sobre = [
        ["Venda por produto do BIP360 — extração avulsa para a curva ABC da precificação"],
        ["Relatório: Financeiro › Venda de Produto por Operador (finRelVendaOperadorPDV2.jsf)"],
        ["Quantidade: Índice 'Quantidade'; valor: Índice 'Valor' (R$), somados por loja, mês e produto (todos os operadores)"],
        ["Período: 01/11/2025 a 24/11/2026, mês a mês (novembro/2026 parcial, até o dia 24)"],
        ["Aba 'situacao': ok = lido; sem_venda = o BIP respondeu 'nenhum registro' no mês"],
        ["Conferência: SJP 1, outubro/2026 — R$ 156.162,14 aqui e no Metas (indicadores_diarios)"],
    ]
    gravar_aba(planilha, "sobre", sobre)
    try:
        planilha.del_worksheet(planilha.worksheet("Página1"))
    except gspread.WorksheetNotFound:
        pass
    try:
        planilha.del_worksheet(planilha.worksheet("Sheet1"))
    except gspread.WorksheetNotFound:
        pass
    print(f"Planilha atualizada: {len(corpo)} linha(s) de venda, {len(situacao) - 1} loja-mês.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
