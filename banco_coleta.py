"""
Escrita dupla — manda a coleta também para o banco do sistema novo.

Os scrapers continuam gravando na planilha exatamente como antes. Depois,
este módulo manda o MESMO dado para a função `ingerir-coleta` do Supabase
(repositório labs3hat/metas-41944e89, docs/05 Fase 2).

Regra de ouro: a escrita no banco NUNCA derruba a escrita na planilha.
Nenhum método daqui levanta exceção — falha vira log de aviso e o scraper
segue. Sem a variável COLETA_TOKEN, tudo aqui fica desligado.

Idempotente: o banco substitui o dia inteiro da loja a cada envio, então
reenviar é seguro.

Variáveis de ambiente:
  COLETA_TOKEN                token dos coletores (secret do GitHub Actions)
  COLETA_URL                  opcional — URL da função (padrão: produção)
  COLETA_BACKFILL_MAX_POR_LOJA  opcional — quantos dias que faltam SÓ no banco
                              cada loja recoleta por execução (padrão 2), para
                              não estourar o tempo do workflow
"""

from __future__ import annotations

import json
import logging
import os
import time
import urllib.error
import urllib.request
from dataclasses import dataclass

log = logging.getLogger("banco_coleta")

URL_PADRAO = "https://bfwybphkxfffazkcobmm.supabase.co/functions/v1/ingerir-coleta"

# As categorias que o banco conhece (tabela `categorias`). O parse do XLS
# também devolve "total" por operador, que não é categoria.
CATEGORIAS = ("shake", "chantilly", "agua", "milk", "canecake")

# O scraper de indicadores usa as posições antigas da planilha como chave.
# No banco a chave é a da loja: CTBA 7 é "ctba7q" e a linha "ctba9" é MGA 8
# (docs/01 seção 2 — nunca usar ctba9 para MGA 8 no sistema novo).
CHAVE_NO_BANCO = {"ctba7": "ctba7q", "ctba9": "mga8"}


def chave_no_banco(chave_do_scraper: str) -> str:
    return CHAVE_NO_BANCO.get(chave_do_scraper, chave_do_scraper)


def data_iso(data_br: str) -> str:
    """'22/09/2026' → '2026-09-22'."""
    dia, mes, ano = data_br.split("/")
    return f"{ano}-{mes}-{dia}"


@dataclass(frozen=True)
class ClienteColeta:
    url: str
    token: str
    tentativas: int = 3
    timeout_s: int = 20
    backfill_max_por_loja: int = 2

    @classmethod
    def do_ambiente(cls) -> "ClienteColeta | None":
        token = os.environ.get("COLETA_TOKEN", "").strip()
        if not token:
            log.info("COLETA_TOKEN ausente: escrita no banco desligada (só planilha).")
            return None
        try:
            backfill = int(os.environ.get("COLETA_BACKFILL_MAX_POR_LOJA", "2"))
        except ValueError:
            log.warning("COLETA_BACKFILL_MAX_POR_LOJA inválido; usando 2.")
            backfill = 2
        return cls(
            url=os.environ.get("COLETA_URL", "").strip() or URL_PADRAO,
            token=token,
            backfill_max_por_loja=max(0, backfill),
        )

    # ── Transporte ────────────────────────────────────────────────────────
    def _enviar(self, corpo: dict) -> dict | None:
        """POST com timeout e retry exponencial (2s, 4s) em falha de rede ou 5xx.
        4xx não repete: é dado recusado pelo banco, e a mensagem diz por quê.
        Nunca levanta exceção; devolve None na falha."""
        dados = json.dumps(corpo, ensure_ascii=False).encode("utf-8")
        acao = corpo.get("acao")
        for tentativa in range(1, self.tentativas + 1):
            pedido = urllib.request.Request(
                self.url,
                data=dados,
                method="POST",
                headers={"Content-Type": "application/json", "x-coleta-token": self.token},
            )
            try:
                with urllib.request.urlopen(pedido, timeout=self.timeout_s) as resposta:
                    return json.loads(resposta.read().decode("utf-8") or "{}")
            except urllib.error.HTTPError as erro:
                detalhe = erro.read().decode("utf-8", "replace")[:300]
                if 400 <= erro.code < 500:
                    log.warning("  [banco] %s recusado (HTTP %d): %s", acao, erro.code, detalhe)
                    return None
                log.warning("  [banco] %s: HTTP %d na tentativa %d: %s",
                            acao, erro.code, tentativa, detalhe)
            except Exception as erro:  # rede, timeout, JSON — nunca derruba o scraper
                log.warning("  [banco] %s: falha na tentativa %d: %s", acao, tentativa, erro)
            if tentativa < self.tentativas:
                time.sleep(2 ** tentativa)
        log.warning("  [banco] %s: desistindo após %d tentativas.", acao, self.tentativas)
        return None

    # ── Ações ─────────────────────────────────────────────────────────────
    def enviar_vendas(self, loja: str, data_br: str, operadores: dict[str, dict],
                      linhas: list[dict] | None = None) -> bool:
        """Vendas de UM dia de UMA loja, por operador. Dia sem venda: operadores={}.
        `linhas`: as linhas cruas do relatório ({operador, produto, quantidade}),
        que o banco classifica pelas regras cadastradas; [] = dia sem venda."""
        lista = [
            {
                "nome": nome,
                "valores": {
                    cat: float(valores.get(cat, 0) or 0)
                    for cat in CATEGORIAS
                    if float(valores.get(cat, 0) or 0) > 0
                },
            }
            for nome, valores in operadores.items()
        ]
        corpo = {
            "acao": "vendas",
            "loja": chave_no_banco(loja),
            "data": data_iso(data_br),
            "operadores": lista,
        }
        if linhas is not None:
            corpo["linhas"] = linhas
        resposta = self._enviar(corpo)
        if resposta is not None:
            log.info("  [banco] vendas %s %s: %s linha(s), %s linha(s) de produto", loja, data_br,
                     resposta.get("linhas"), "sem" if linhas is None else len(linhas))
        return resposta is not None

    def enviar_indicadores(self, loja: str, data_br: str, receita: float,
                           ticket_medio: float, pessoas: int) -> bool:
        resposta = self._enviar({
            "acao": "indicadores",
            "loja": chave_no_banco(loja),
            "data": data_iso(data_br),
            "receita": round(float(receita), 2),
            "ticket_medio": round(float(ticket_medio), 2),
            "pessoas": int(pessoas),
        })
        if resposta is not None:
            log.info("  [banco] indicadores %s %s gravados", loja, data_br)
        return resposta is not None

    def enviar_totem(self, loja: str, data_br: str, pessoas_totem: int, pessoas_loja: int) -> bool:
        """Clientes atendidos pelo Totem e pela loja em UM dia (Relatório de
        Venda, Qtd. Pessoas Atendidas). Base do % de uso do Totem."""
        resposta = self._enviar({
            "acao": "totem",
            "loja": chave_no_banco(loja),
            "data": data_iso(data_br),
            "pessoas_totem": int(pessoas_totem),
            "pessoas_loja": int(pessoas_loja),
        })
        if resposta is not None:
            log.info("  [banco] Totem %s %s: %d de %d cliente(s)", loja, data_br,
                     pessoas_totem, pessoas_loja)
        return resposta is not None

    def enviar_atendimentos(self, loja: str, data_br: str, atendimentos: dict[str, int]) -> bool:
        """Clientes atendidos por operador em UM dia de UMA loja (relatório de
        ranking, índice Pessoas Atendidas). Base do % de uso do Totem. Dia sem
        venda: atendimentos={}."""
        resposta = self._enviar({
            "acao": "atendimentos",
            "loja": chave_no_banco(loja),
            "data": data_iso(data_br),
            "operadores": [
                {"nome": nome, "pessoas": int(pessoas)} for nome, pessoas in atendimentos.items()
            ],
        })
        if resposta is not None:
            total = sum(atendimentos.values())
            log.info("  [banco] atendimentos %s %s: %d operador(es), %d cliente(s)",
                     loja, data_br, len(atendimentos), total)
        return resposta is not None

    def dias_presentes(self, tipo: str, datas_br: list[str]) -> set[tuple[str, str]] | None:
        """(loja NA CHAVE DO BANCO, 'DD/MM/AAAA') já gravados no banco — compare
        com chave_no_banco(sua_chave). None se não deu para perguntar: aí
        ninguém é recoletado por causa do banco."""
        if not datas_br:
            return set()
        isos = sorted(data_iso(d) for d in datas_br)
        resposta = self._enviar({"acao": "presentes", "tipo": tipo, "desde": isos[0], "ate": isos[-1]})
        if resposta is None:
            return None
        presentes: set[tuple[str, str]] = set()
        for linha in resposta.get("dias") or []:
            ano, mes, dia = str(linha["data"]).split("-")
            presentes.add((str(linha["loja"]), f"{dia}/{mes}/{ano}"))
        return presentes

    def registrar_execucao(self, tipo: str, iniciado_em: str, status: str,
                           lojas_ok: int, lojas_erro: int,
                           pendencias: list[dict], resumo: str) -> None:
        self._enviar({
            "acao": "execucao",
            "tipo": tipo,
            "iniciado_em": iniciado_em,
            "status": status,
            "lojas_ok": lojas_ok,
            "lojas_erro": lojas_erro,
            "pendencias": pendencias,
            "log": resumo,
        })
