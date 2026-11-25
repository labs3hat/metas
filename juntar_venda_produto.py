"""
Junta os CSVs da extração avulsa de venda por produto (extracao_venda_produto.py)
num arquivo só: CSV e XLSX (abas "venda" e "situacao"). Não lê nem grava banco.

Uso: python juntar_venda_produto.py <pasta com os CSVs> <prefixo de saída>
"""

import csv
import glob
import os
import sys

import openpyxl


def ler(padrao: str) -> tuple[list[str], list[list[str]]]:
    cabecalho: list[str] = []
    linhas: list[list[str]] = []
    for caminho in sorted(glob.glob(padrao, recursive=True)):
        with open(caminho, newline="", encoding="utf-8") as f:
            r = csv.reader(f)
            cab = next(r, None)
            if cab is None:
                continue
            if cabecalho and cab != cabecalho:
                raise SystemExit(f"Cabeçalho diferente em {caminho}: {cab} != {cabecalho}")
            cabecalho = cab
            linhas.extend(r)
    return cabecalho, linhas


def main() -> int:
    pasta, prefixo = sys.argv[1], sys.argv[2]
    # "**" com recursive=True pega a pasta e as subpastas (um artifact por loja).
    cab_v, vendas = ler(os.path.join(pasta, "**", "venda_produto_*.csv"))
    cab_s, situacao = ler(os.path.join(pasta, "**", "situacao_*.csv"))
    if not cab_v or not cab_s:
        raise SystemExit(f"Nenhum CSV da extração em {pasta}.")
    vendas.sort(key=lambda l: (l[0], l[1], -float(l[4] or 0)))
    situacao.sort(key=lambda l: (l[0], l[1]))

    with open(f"{prefixo}.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(cab_v)
        w.writerows(vendas)
    with open(f"{prefixo}_situacao.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(cab_s)
        w.writerows(situacao)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "venda"
    ws.append(cab_v)
    for l in vendas:
        ws.append([l[0], l[1], l[2], float(l[3]), float(l[4])])
    ws2 = wb.create_sheet("situacao")
    ws2.append(cab_s)
    for l in situacao:
        ws2.append([l[0], l[1], l[2], int(l[3]), l[4]])
    wb.save(f"{prefixo}.xlsx")

    erros = [l for l in situacao if l[2] not in ("ok", "sem_venda")]
    print(f"{len(vendas)} linha(s) de venda; {len(situacao)} loja-mês; {len(erros)} com erro")
    for l in erros:
        print(f"  ✗ {l[0]} {l[1]}: {l[2]} — {l[4]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
