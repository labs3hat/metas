"""
Bebida láctea — leitura do relatório "Saída de Composição" do BIP360 (só a
lógica: sem navegador, sem rede, sem banco). Quem navega é lactea_scraper.py.

O relatório (Estoque → Relatórios → Saída de Composição, Insumo = Sim) lista o
consumo TEÓRICO dos insumos pelas vendas — o "APTE" da planilha de desperdício.
Cada insumo aparece até 2 vezes, uma por Tipo de Saída: "Venda" e "Criação de
Produtos" (loja que faz Chiquinho Pote). As duas SOMAM (Rose, 14/11/2026; é
também o que o projeto de suprimentos faz).

Regras de leitura (falhar fechado — número plausível e errado é o pior caso):
  - o cabeçalho tem de ser exatamente Descrição | Unidade de Medida | Tipo de
    Saída | Qtd. Venda;
  - Tipo de Saída desconhecido aborta (devolução, perda… não podem somar calados);
  - quantidade em branco ou não numérica aborta;
  - o mesmo produto e tipo duas vezes aborta (export corrompido somaria em dobro);
  - as linhas de um sabor têm de ter a mesma unidade, senão a soma não existe.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

CABECALHO = ("Descrição", "Unidade de Medida", "Tipo de Saída", "Qtd. Venda")
TIPOS_SAIDA = ("Venda", "Criação de Produtos")
SABORES = ("baunilha", "chocolate")


class RelatorioInvalido(Exception):
    """O arquivo não é o relatório esperado ou veio corrompido."""


@dataclass(frozen=True)
class Linha:
    numero: int
    descricao: str
    unidade: str
    tipo: str
    quantidade: float


@dataclass
class SaborDoDia:
    sabor: str
    litros: float
    unidade: str
    linhas: list[dict] = field(default_factory=list)


def normalizar(texto: str) -> str:
    sem_acento = unicodedata.normalize("NFD", texto.replace(" ", " "))
    sem_acento = "".join(c for c in sem_acento if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sem_acento).strip().upper()


def numero(valor: object, onde: str) -> float:
    """Célula de quantidade: número, ou texto no formato brasileiro (1.234,5)."""
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    texto = str(valor if valor is not None else "").strip()
    if not texto:
        raise RelatorioInvalido(f"{onde}: quantidade em branco.")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError as erro:
        raise RelatorioInvalido(f"{onde}: quantidade não numérica {valor!r}.") from erro


def codigo_da_franquia(nome: str) -> str:
    """'MARINGÁ 09 - BONIFÁCIO MALL (PR)' e 'Maringá 09 (PR)' → 'MARINGA 09'.

    O BIP360 não usa o mesmo nome em todo lugar: na troca de loja a franquia
    vem com o shopping; no filtro Franquia da Saída de Composição, às vezes só
    com cidade e número (Bonifácio, 14/11/2026). Cidade + número identificam a
    unidade sem ambiguidade."""
    base = normalizar(nome).split(" - ")[0]
    return re.sub(r"\s*\([A-Z]{2}\)$", "", base).strip()


def ler_linhas(folha) -> list[Linha]:
    """Lê a planilha exportada. `folha` tem .nrows, .ncols e .cell_value(r, c)
    (o contrato de scraper.open_workbook_smart)."""
    if folha.nrows == 0:
        raise RelatorioInvalido("planilha vazia.")
    cab = [str(folha.cell_value(0, c)).strip() for c in range(folha.ncols)]
    while cab and cab[-1] == "":
        cab.pop()
    if tuple(cab) != CABECALHO:
        raise RelatorioInvalido(
            f"cabeçalho inesperado: {' | '.join(cab) or '(vazio)'} "
            f"(esperado {' | '.join(CABECALHO)}). É o relatório Saída de Composição?"
        )
    linhas: list[Linha] = []
    vistos: dict[tuple[str, str], int] = {}
    for r in range(1, folha.nrows):
        celulas = [folha.cell_value(r, c) for c in range(folha.ncols)]
        if all(str(v).strip() == "" for v in celulas):
            continue
        if any(str(v).strip() != "" for v in celulas[len(CABECALHO):]):
            raise RelatorioInvalido(f"linha {r + 1}: dado fora das colunas do cabeçalho.")
        descricao = str(celulas[0]).strip()
        if not descricao:
            raise RelatorioInvalido(f"linha {r + 1}: dado sem descrição do produto.")
        tipo = str(celulas[2]).strip()
        if tipo not in TIPOS_SAIDA:
            raise RelatorioInvalido(
                f"linha {r + 1} ({descricao}): Tipo de Saída "
                f"{repr(tipo) if tipo else 'em branco'} (aceitos: {', '.join(TIPOS_SAIDA)})."
            )
        chave = (normalizar(descricao), tipo)
        if chave in vistos:
            raise RelatorioInvalido(
                f"linha {r + 1}: {descricao} / {tipo} repete a linha {vistos[chave]}."
            )
        vistos[chave] = r + 1
        linhas.append(Linha(
            numero=r + 1,
            descricao=descricao,
            unidade=str(celulas[1]).strip(),
            tipo=tipo,
            quantidade=numero(celulas[3], f"linha {r + 1} ({descricao})"),
        ))
    return linhas


def sabor_da_bebida_lactea(descricao: str) -> str | None:
    """'Bebida Láctea Uht Baunilha' → 'baunilha'. Outro produto → None.
    Bebida láctea sem sabor reconhecido aborta: é produto novo que alguém
    precisa olhar, não algo para ignorar calado."""
    nome = normalizar(descricao)
    if "BEBIDA LACTEA" not in nome:
        return None
    achados = [s for s in SABORES if s.upper() in nome]
    if len(achados) != 1:
        raise RelatorioInvalido(f"bebida láctea com sabor não reconhecido: {descricao!r}.")
    return achados[0]


def somar_por_sabor(linhas: list[Linha]) -> list[SaborDoDia]:
    """Soma Venda + Criação de Produtos de cada sabor. Sabor sem linha fica de
    fora (o banco grava zero quando o relatório tem outras linhas)."""
    por_sabor: dict[str, list[Linha]] = {}
    for linha in linhas:
        sabor = sabor_da_bebida_lactea(linha.descricao)
        if sabor:
            por_sabor.setdefault(sabor, []).append(linha)
    resultado: list[SaborDoDia] = []
    for sabor in SABORES:
        do_sabor = por_sabor.get(sabor)
        if not do_sabor:
            continue
        unidades = {normalizar(l.unidade) for l in do_sabor}
        if len(unidades) != 1:
            raise RelatorioInvalido(
                f"{sabor}: linhas com unidades diferentes ({', '.join(sorted(unidades))}) — a soma não existe."
            )
        resultado.append(SaborDoDia(
            sabor=sabor,
            litros=round(sum(l.quantidade for l in do_sabor), 4),
            unidade=do_sabor[0].unidade,
            linhas=[{"descricao": l.descricao, "tipo": l.tipo, "quantidade": l.quantidade}
                    for l in do_sabor],
        ))
    return resultado
