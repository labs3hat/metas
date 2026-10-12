"""Correções no "Resumo Distribuições Lucros" (.xlsx sem macro), aprovadas por André em 11/10/2026:

  1. Payback: linha de 2026 (16) com fórmulas no mesmo padrão da de 2025, pagamentos de
     empréstimos de 2026 (banco + interno, planilha de empréstimos de 30/09/2026), fórmulas de
     ROE/ROI estendidas até a linha 16 e "Emprést. Aberto" atualizado (Jockey e Outlet).
  2. Linha "Total c/ Rend" (não é mais usada): esvaziada no único bloco recente que ainda a tem
     (dez/25). Os blocos de 2023-2025 alimentam outras fórmulas e ficam como histórico.
  3. Nomes de loja padronizados (texto compartilhado, vale para todos os blocos).
  4. Nota na aba Payback sobre o valor fixo (Aya Amintas até jun/26, Eurogarden desde jul/26).

Mesmo princípio das planilhas gerenciais: só as partes XML tocadas são reescritas, e
`verificar` acusa qualquer parte alterada fora do esperado.
"""
from __future__ import annotations

import html
import re
import zipfile

from corrige_planilha import celula, estilo, mapa_abas, shared_strings, substitui_celula, valor

ABA_PAYBACK = "Payback"
ABA_BLOCOS = "Distribuição 2023-24-25-26"
LINHA_2026 = 16

# "Dist. Lucros" de 2026: total anual da aba de cada loja (padrão: 2024 → 2025 → 2026 avança 10 colunas)
DIST_2026 = {
    "C": "'3HAT SJP - 2018 a 2026'!BE86",
    "F": "'3HAT ESTACAO - 2022 a 2026'!AM98",
    "I": "'3HAT JOCKEY - 2023 a 2026'!AH98",
    "L": "'3HAT OUTLET - 2023 a 2026'!AH98",
    "O": "'3HAT PALLADIUM - 2025 a 2026'!X98",
    "S": "'3HT AVENIDA - 2025 a 2026'!O98",
    "V": "'3HT CIDADE - 2025 a 2026'!O98",
}
# "Pgto Emprést." de 2026 = parcelas pagas a bancos + devoluções de empréstimos internos no ano,
# conforme "2026.09.30 - PLANILHA EMPRÉSTIMOS TOMADOS" (abas Empréstimos Bancos e Capital Próprio).
PGTO_EMPRESTIMOS_2026 = {
    "D": "0",                      # SJP 1
    "G": "0",                      # CTBA 3
    "J": "64725.809+185131.54",    # CTBA 5: Pronampes antecipados + devoluções internas
    "M": "43150.54+52568.28",      # CL 2
    "P": "0",                      # CTBA 7
    "T": "0",                      # MGA 3
    "W": "0",                      # MGA 5
}
EMPRESTIMO_ABERTO = {"I24": "31560.77", "L24": "0"}  # Jockey e Outlet em 30/09/2026 (antes: 279.769,94 e 146.795,43)
CABECALHOS = {
    "C2": "SJP 1 - São José dos Pinhais", "F2": "CTBA 3 - Estação", "I2": "CTBA 5 - Jockey",
    "L2": "CL 2 - Outlet", "O2": "CTBA 7 - Quiosque Palladium", "S2": "MGA 3 - Avenida", "V2": "MGA 5 - Cidade",
}
NOTA_B35 = ("Nota (11/10/2026): o valor fixo mensal por loja é o investimento Aya Amintas (5 lojas, até jun/26) "
            "e Eurogarden (a partir de jul/26) e conta como lucro distribuído. Pgto Emprést. 2026 = bancos + "
            "empréstimos internos, conforme planilha de empréstimos de 30/09/2026.")
NOMES = {
    "SJP": "SJP 1 - São José dos Pinhais", "L SJP": "SJP 1 - São José dos Pinhais",
    "Estação": "CTBA 3 - Estação", "L Estação": "CTBA 3 - Estação",
    "Jockey": "CTBA 5 - Jockey", "L Jockey": "CTBA 5 - Jockey",
    "Outlet": "CL 2 - Outlet", "Q Outlet": "CL 2 - Outlet",
    "Palladium": "CTBA 7 - Quiosque Palladium", "Q Palladium": "CTBA 7 - Quiosque Palladium",
    "L Palladium": "CTBA 11 - Loja Palladium",
    "Cidade": "MGA 5 - Cidade", "L Cidade": "MGA 5 - Cidade",
    "Avenida": "MGA 3 - Avenida", "L Avenida": "MGA 3 - Avenida",
    "Q Avenida": "MGA 7 - Quiosque Avenida",
    "Q Estação/Havan": "MGA 8 - Havan",
    "Q Estação": "CTBA 9 - Quiosque Estação (hoje MGA 8 Havan)",
    "L Bonifácio Mall": "MGA 9 - Bonifácio Mall",
}
TOTAL_C_REND = ("N504", "O504", "P504", "Q504", "R504", "S504", "T504")


def _garantir_linha(xml: str, n: int) -> str:
    if re.search(rf'<row r="{n}"[^>]*>', xml) or re.search(rf'<row r="{n}"[^>]*/>', xml):
        return re.sub(rf'<row r="{n}"([^>]*)/>', rf'<row r="{n}"\1></row>', xml)
    for rm in re.finditer(r'<row r="(\d+)"', xml):
        if int(rm.group(1)) > n:
            return xml[: rm.start()] + f'<row r="{n}"></row>' + xml[rm.start():]
    if "</sheetData>" in xml:  # nenhuma linha depois de n: entra no fim
        return xml.replace("</sheetData>", f'<row r="{n}"></row></sheetData>', 1)
    raise ValueError(f"não achei onde inserir a linha {n}")


def _f(ref: str, formula: str, s: str) -> str:
    return f'<c r="{ref}"{s}><f>{html.escape(formula, quote=False)}</f></c>'


def _n(ref: str, numero: str, s: str) -> str:
    return f'<c r="{ref}"{s}><v>{numero}</v></c>'


def _texto(ref: str, texto: str, s: str) -> str:
    return (f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">'
            f'{html.escape(texto, quote=False)}</t></is></c>')


def _estende_ate_16(formula: str) -> str:
    return re.sub(r"([A-Z]{1,2})(4|8):\1(15)\b", lambda m: f"{m.group(1)}{m.group(2)}:{m.group(1)}16", formula)


def corrigir_payback(xml: str, ss: list[str]) -> tuple[str, list[str]]:
    mud: list[str] = []
    xml = _garantir_linha(xml, LINHA_2026)
    s15 = lambda col: estilo(celula(xml, f"{col}15"))
    if valor(celula(xml, f"B{LINHA_2026}"), ss) not in (None, "", "2026"):
        raise ValueError(f"B{LINHA_2026} já tem {valor(celula(xml, f'B{LINHA_2026}'), ss)!r}; não sobrescrevo")
    xml = substitui_celula(xml, f"B{LINHA_2026}", _n(f"B{LINHA_2026}", "2026", s15("B")))
    for col, ref_loja in DIST_2026.items():
        xml = substitui_celula(xml, f"{col}{LINHA_2026}", _f(f"{col}{LINHA_2026}", ref_loja, s15(col)))
    for col, expr in PGTO_EMPRESTIMOS_2026.items():
        ref = f"{col}{LINHA_2026}"
        xml = substitui_celula(xml, ref, _f(ref, expr, s15(col)) if "+" in expr else _n(ref, expr, s15(col)))
    for dist_col, pct_col in zip("CFILOSV", "EHKNQUX"):
        pg_col = chr(ord(dist_col) + 1)
        ref = f"{pct_col}{LINHA_2026}"
        xml = substitui_celula(xml, ref, _f(ref, f"SUM({dist_col}$4:{pg_col}{LINHA_2026})/{dist_col}$19", s15(pct_col)))
    mud.append("Payback linha 16: 2026 com Dist. Lucros (abas das lojas), Pgto Emprést. 2026 e % acumulado")
    # ROE (20), ROE % (21, matricial) e ROI (22): até a linha 16
    n_ext = 0

    def estende_celula(m: re.Match) -> str:
        nonlocal n_ext
        inner = m.group(3)
        fm = re.search(r"<f([^>]*)>(.*?)</f>", inner, re.S)
        if not fm:
            return m.group(0)
        antiga = html.unescape(fm.group(2))
        nova = _estende_ate_16(antiga)
        if nova == antiga:
            return m.group(0)
        novo_inner = inner[: fm.start()] + f"<f{fm.group(1)}>{html.escape(nova, quote=False)}</f>" + inner[fm.end():]
        novo_inner = re.sub(r"<v>.*?</v>", "", novo_inner, flags=re.S)
        n_ext += 1
        return f'<c r="{m.group(1)}"{m.group(2)}>{novo_inner}</c>'

    for lin in (20, 21, 22):
        # só células abertas (<c ...>…</c>); as autofechadas (<c .../>) não têm fórmula e não podem
        # engolir a célula seguinte. Substituição por callback: nada de offsets velhos após cada troca.
        xml = re.sub(rf'<c r="([A-Z]{{1,2}}{lin})"([^>]*?)(?<!/)>(.*?)</c>', estende_celula, xml, flags=re.S)
    if n_ext != 7 + 7 + 14:
        raise ValueError(f"esperava estender 28 fórmulas nas linhas 20-22, estendi {n_ext}")
    mud.append("Payback linhas 20-22: ROE, ROE % e ROI passam a somar até 2026")
    for ref, num in EMPRESTIMO_ABERTO.items():
        antes = valor(celula(xml, ref), ss)
        xml = substitui_celula(xml, ref, _n(ref, num, estilo(celula(xml, ref))))
        mud.append(f"Payback {ref} (Emprést. Aberto): {antes} → {num}")
    for ref, nome in CABECALHOS.items():
        atual = valor(celula(xml, ref), ss) or ""
        sufixo = atual.split("\n", 1)[1] if "\n" in atual else ""
        novo = nome + ("\n" + sufixo if sufixo else "")
        if atual != novo:
            xml = substitui_celula(xml, ref, _texto(ref, novo, estilo(celula(xml, ref))))
    mud.append("Payback linha 2: nomes das lojas padronizados")
    xml = _garantir_linha(xml, 35)
    xml = substitui_celula(xml, "B35", _texto("B35", NOTA_B35, estilo(celula(xml, "B31"))))
    mud.append("Payback B35: nota sobre valor fixo (Aya Amintas / Eurogarden) e Pgto Emprést. 2026")
    return xml, mud


def corrigir_blocos(xml: str, ss: list[str]) -> tuple[str, list[str]]:
    rotulo = valor(celula(xml, TOTAL_C_REND[0]), ss)
    if rotulo != "Total c/ Rend":
        raise ValueError(f"{TOTAL_C_REND[0]} deveria ser 'Total c/ Rend', é {rotulo!r}")
    for ref in TOTAL_C_REND:
        m = celula(xml, ref)
        if m:
            xml = substitui_celula(xml, ref, f'<c r="{ref}"{estilo(m)}/>')
    return xml, ["Blocos: linha 'Total c/ Rend' do bloco dez/25 (N504:T504) esvaziada"]


def corrigir_nomes(ssxml: str) -> tuple[str, list[str]]:
    trocados = []
    def troca(m: re.Match) -> str:
        texto = html.unescape(re.sub(r"<[^>]+>", "", m.group(1)))
        if texto in NOMES:
            trocados.append(f"{texto!r} → {NOMES[texto]!r}")
            return f"<si><t>{html.escape(NOMES[texto], quote=False)}</t></si>"
        return m.group(0)
    novo = re.sub(r"<si>(.*?)</si>", troca, ssxml, flags=re.S)
    faltam = set(NOMES) - {t.split(" → ")[0].strip("'") for t in trocados}
    if faltam:
        raise ValueError(f"nomes não encontrados no texto compartilhado: {sorted(faltam)}")
    return novo, [f"Nomes padronizados ({len(trocados)} textos): " + "; ".join(trocados)]


def corrige(origem: str, destino: str) -> list[str]:
    z = zipfile.ZipFile(origem)
    abas = mapa_abas(z); ss = shared_strings(z)
    partes: dict[str, bytes] = {}
    mud: list[str] = []
    pb, m1 = corrigir_payback(z.read(abas[ABA_PAYBACK]).decode("utf-8"), ss)
    partes[abas[ABA_PAYBACK]] = pb.encode("utf-8"); mud += m1
    bl, m2 = corrigir_blocos(z.read(abas[ABA_BLOCOS]).decode("utf-8"), ss)
    partes[abas[ABA_BLOCOS]] = bl.encode("utf-8"); mud += m2
    sx, m3 = corrigir_nomes(z.read("xl/sharedStrings.xml").decode("utf-8"))
    partes["xl/sharedStrings.xml"] = sx.encode("utf-8"); mud += m3
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r"<calcPr[^>]*/>", wbx)
    if m and "fullCalcOnLoad" not in m.group(0):
        wbx = wbx.replace(m.group(0), m.group(0)[:-2] + ' fullCalcOnLoad="1"/>')
        partes["xl/workbook.xml"] = wbx.encode("utf-8")
    with zipfile.ZipFile(destino, "w") as out:
        for info in z.infolist():
            out.writestr(info, partes.get(info.filename, z.read(info.filename)), compress_type=info.compress_type)
    return mud


def verificar(origem: str, destino: str) -> None:
    import openpyxl

    a, b = zipfile.ZipFile(origem), zipfile.ZipFile(destino)
    if b.testzip() is not None:
        raise ValueError(f"zip corrompido: {b.testzip()}")
    if [i.filename for i in a.infolist()] != [i.filename for i in b.infolist()]:
        raise ValueError("lista de partes do zip mudou")
    abas = mapa_abas(a)
    permitidas = {"xl/workbook.xml", "xl/sharedStrings.xml", abas[ABA_PAYBACK], abas[ABA_BLOCOS]}
    fora = [i.filename for i, j in zip(a.infolist(), b.infolist()) if i.CRC != j.CRC and i.filename not in permitidas]
    if fora:
        raise ValueError(f"partes alteradas fora do esperado: {fora}")
    wb = openpyxl.load_workbook(destino)
    pb = wb[ABA_PAYBACK]
    if pb["B16"].value != 2026:
        raise ValueError(f"B16 ≠ 2026: {pb['B16'].value!r}")
    for col, ref_loja in DIST_2026.items():
        if str(pb[f"{col}16"].value) != "=" + ref_loja:
            raise ValueError(f"{col}16 = {pb[f'{col}16'].value!r}, esperava ={ref_loja}")
    if "C16" not in str(pb["C20"].value) or "C16" not in str(pb["C22"].value):
        raise ValueError(f"ROE/ROI não somam até 2026: C20={pb['C20'].value!r} C22={pb['C22'].value!r}")
    if float(pb["I24"].value) != 31560.77 or float(pb["L24"].value) != 0:
        raise ValueError(f"Emprést. Aberto não atualizado: I24={pb['I24'].value!r} L24={pb['L24'].value!r}")
    if not str(pb["B35"].value or "").startswith("Nota (11/10/2026)"):
        raise ValueError("nota B35 ausente")
    bl = wb[ABA_BLOCOS]
    if any(bl[r].value is not None for r in TOTAL_C_REND):
        raise ValueError("N504:T504 não ficou vazio")
    if bl["A630"].value != "SJP 1 - São José dos Pinhais" or bl["A639"].value != "MGA 8 - Havan":
        raise ValueError(f"nomes não padronizados: A630={bl['A630'].value!r} A639={bl['A639'].value!r}")
    # nada mudou nas abas das lojas: mesma quantidade de fórmulas e mesmos totais anuais
    wa = openpyxl.load_workbook(origem, data_only=True); wbv = openpyxl.load_workbook(destino, data_only=True)
    for sh, ref in [(r.split("!")[0].strip("'"), r.split("!")[1]) for r in DIST_2026.values()]:
        if wa[sh][ref].value != wbv[sh][ref].value:
            raise ValueError(f"{sh}!{ref} mudou: {wa[sh][ref].value} → {wbv[sh][ref].value}")
