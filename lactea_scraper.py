"""
Coleta diária do APTE de bebida láctea (gestão, 14/11/2026).

Lê no BIP360 o relatório Estoque → Relatórios → Saída de Composição (Insumo =
Sim) de cada loja, soma Venda + Criação de Produtos por sabor (lactea.py) e
manda para o banco do Metas (banco_coleta.enviar_lactea). É ACOMPANHAMENTO:
o desperdício do mês e a premiação seguem pela planilha da gestão.

Módulo NOVO, ao lado do scraper.py — reaproveita dele o login e a troca de
loja, sem mudar nada lá (CLAUDE.md do sistema de metas, regra 5).

Cada conferência que falha ABORTA a loja com o motivo (o arquivo errado não dá
erro sozinho: entrariam números plausíveis de outra loja ou de outro dia):
  - a página tem de ser a "Saída de Composição";
  - as datas lidas de volta da tela têm de ser as pedidas;
  - filtros no padrão (Insumo = Sim, Tipo de Saída = Todos, nenhum Produto);
  - o filtro Franquia tem de marcar SÓ a loja-alvo;
  - "Nenhum registro" vira `sem_registros`, nunca consumo zero.

Modos:
  python lactea_scraper.py                       # ontem (D-1, São Paulo), grava
  python lactea_scraper.py --de 01/11/2026 --ate 13/11/2026   # dia a dia, grava
  python lactea_scraper.py --janela --de 05/10/2026 --ate 11/10/2026
        # validação: UM período somado por loja, só imprime (não grava)
  --lojas sjp1,ctba3   # restringe as lojas
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import tempfile
import time
from datetime import datetime, timedelta

from playwright.sync_api import sync_playwright

import scraper as s
from banco_coleta import ClienteColeta
from lactea import RelatorioInvalido, codigo_da_franquia, ler_linhas, somar_por_sabor

log = logging.getLogger("lactea")

URL_RELATORIO = "https://itfgestor.com.br/ITFGestor/estRelSaidaComposicao.jsf"
F = "formRelMovimentacaoEstoque"
# Os padrões da tela (medidos pelo projeto de suprimentos em 09/10/2026). O
# robô não muda filtro nenhum: confere e aborta se divergirem.
FILTROS_ESPERADOS = {
    f"{F}:comboTipoSaida_input": "",  # Todos = Venda + Criação de Produtos
    f"{F}:comboInsumo_input": "S",
}


class FalhaLoja(Exception):
    """A loja não pôde ser lida com segurança — motivo vai para o log."""


def ir_para_relatorio(page) -> None:
    page.goto(URL_RELATORIO, wait_until="load", timeout=45000)
    s.wait_bip_idle(page, timeout=45)
    corpo = page.locator("body").inner_text(timeout=10000)
    if "publico/login.jsf" in page.url:
        raise FalhaLoja("a sessão caiu (voltou para o login)")
    if "Saída de Composição" not in corpo:
        raise FalhaLoja("a página do relatório não abriu (o endereço da Saída de Composição mudou?)")


def preencher_datas(page, inicio: str, fim: str) -> None:
    page.evaluate(
        """([pares]) => {
            for (const [id, v] of pares) {
                const el = document.getElementById(id);
                if (!el) throw new Error('campo ' + id + ' não encontrado');
                el.focus(); el.value = v;
                el.dispatchEvent(new Event('input', {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
                el.blur();
            }
        }""",
        [[[f"{F}:calendarioDataInicial_input", inicio], [f"{F}:calendarioDataFinal_input", fim]]],
    )
    s.wait_bip_idle(page, timeout=30)
    lidas = page.evaluate(
        """([a, b]) => [document.getElementById(a).value, document.getElementById(b).value]""",
        [f"{F}:calendarioDataInicial_input", f"{F}:calendarioDataFinal_input"],
    )
    if lidas != [inicio, fim]:
        raise FalhaLoja(f"datas na tela {lidas[0]}–{lidas[1]} ≠ pedidas {inicio}–{fim}")


def conferir_filtros(page, alvo: str) -> str:
    valores = page.evaluate(
        """(ids) => Object.fromEntries(ids.map(id => [id, document.getElementById(id)?.value ?? null]))""",
        list(FILTROS_ESPERADOS),
    )
    errados = [f"{k.split(':')[1]}={valores[k]!r} (esperado {v!r})"
               for k, v in FILTROS_ESPERADOS.items() if valores.get(k) != v]
    if errados:
        raise FalhaLoja(f"filtros fora do padrão: {', '.join(errados)}")
    lido = page.evaluate(
        """(f) => {
            const rotulo = (c) => document.querySelector(`label[for="${CSS.escape(c.id)}"]`)?.innerText.trim()
                                  ?? `(sem rótulo, value=${c.value})`;
            const marcadas = Array.from(document.querySelectorAll(`input[type="checkbox"][name="${f}:comboFranquia"]`))
                .filter(c => c.checked).map(rotulo);
            const produtos = Array.from(document.querySelectorAll(`input[type="checkbox"][name="${f}:comboProduto"]`))
                .filter(c => c.checked).length;
            return {marcadas, produtos};
        }""",
        F,
    )
    if lido["produtos"]:
        raise FalhaLoja(f"o filtro Produto tem {lido['produtos']} marcado(s) — o relatório sairia parcial")
    if len(lido["marcadas"]) != 1 or codigo_da_franquia(lido["marcadas"][0]) != codigo_da_franquia(alvo):
        raise FalhaLoja(f"o filtro Franquia marca {lido['marcadas'] or 'nenhuma'}, esperado só {alvo!r}")
    return lido["marcadas"][0]


def pesquisar(page) -> bool:
    """Pesquisa e espera a tabela do resultado ser RECRIADA. Devolve True se
    o BIP360 disse "nenhum registro"."""
    page.evaluate("""() => document.querySelectorAll('table, .ui-datatable').forEach(el => el.__stale = true)""")
    page.locator(f'[id="{F}:botaoPesquisar"]').click(timeout=15000)
    fim = time.time() + 60
    while True:
        time.sleep(0.5)
        nova = page.evaluate(
            """(f) => { const t = document.getElementById(f + ':tabelaRelatorio'); return !!t && !t.__stale; }""", F)
        if nova:
            break
        if time.time() > fim:
            raise FalhaLoja("a tabela do resultado não foi recriada depois de Pesquisar")
    s.wait_bip_idle(page, timeout=45)
    r = page.evaluate(
        """(f) => {
            const growl = document.querySelector('#growl')?.innerText.trim() || '';
            const t = document.getElementById(f + ':tabelaRelatorio');
            return {growl, vazia: /nenhum registro encontrado/i.test(t?.innerText ?? '')};
        }""",
        F,
    )
    if r["growl"]:
        raise FalhaLoja(f"o BIP360 avisou: {r['growl']!r}")
    return bool(r["vazia"])


def baixar_e_ler(page, pasta: str, rotulo: str) -> list:
    download = s.click_xls_export(page)
    caminho = os.path.join(pasta, f"lactea-{rotulo}.xls")
    download.save_as(caminho)
    wb = s.open_workbook_smart(caminho)
    return ler_linhas(wb.sheets()[0])


def ler_periodo(page, loja: dict, inicio: str, fim: str, pasta: str) -> dict:
    """Lê um período (dd/mm/aaaa) de uma loja. Devolve
    {"sem_registros": bool, "itens": [SaborDoDia...], "franquia": str}."""
    ir_para_relatorio(page)
    s.select_store(page, loja["bip_name"])
    ir_para_relatorio(page)  # a troca de loja recarrega a página
    preencher_datas(page, inicio, fim)
    franquia = conferir_filtros(page, loja["bip_name"])
    vazia = pesquisar(page)
    if vazia:
        return {"sem_registros": True, "itens": [], "franquia": franquia}
    linhas = baixar_e_ler(page, pasta, f"{loja['key']}-{inicio.replace('/', '')}-{fim.replace('/', '')}")
    return {"sem_registros": False, "itens": somar_por_sabor(linhas), "franquia": franquia}


def dias_entre(de: str, ate: str) -> list[str]:
    d0 = datetime.strptime(de, "%d/%m/%Y")
    d1 = datetime.strptime(ate, "%d/%m/%Y")
    if d1 < d0:
        raise SystemExit(f"--ate {ate} antes de --de {de}")
    return [(d0 + timedelta(days=i)).strftime("%d/%m/%Y") for i in range((d1 - d0).days + 1)]


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s [%(name)s] %(message)s")
    ap = argparse.ArgumentParser()
    ap.add_argument("--de")
    ap.add_argument("--ate")
    ap.add_argument("--janela", action="store_true",
                    help="soma o período inteiro por loja e só imprime (validação; não grava)")
    ap.add_argument("--lojas", default="", help="chaves separadas por vírgula (padrão: todas)")
    args = ap.parse_args()

    ontem = (s.br_now() - timedelta(days=1)).strftime("%d/%m/%Y")
    de = args.de or ontem
    ate = args.ate or de
    filtro = {k.strip() for k in args.lojas.split(",") if k.strip()}
    lojas = [l for l in s.LOJAS if (not filtro or l["key"] in filtro)]
    cliente = None if args.janela else ClienteColeta.do_ambiente()
    iniciado_em = s.br_now().isoformat()
    log.info("início: %s, %s a %s, %d loja(s)%s", "janela" if args.janela else "diário", de, ate,
             len(lojas), "" if args.janela else ("" if cliente else " — SEM token: não grava"))

    ok, falhas, resumo = [], [], []
    pasta = tempfile.mkdtemp(prefix="lactea-")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1600, "height": 1000})
        page = context.new_page()
        try:
            s.login(page)
            for loja in lojas:
                periodos = [(de, ate)] if args.janela else [(d, d) for d in dias_entre(de, ate)]
                for inicio, fim in periodos:
                    if not s.loja_aberta_em(loja, fim):
                        log.info("%s: ainda não abria em %s — pulada", loja["key"], fim)
                        continue
                    t0 = time.time()
                    try:
                        r = ler_periodo(page, loja, inicio, fim, pasta)
                    except (FalhaLoja, RelatorioInvalido) as erro:
                        log.warning("%s %s–%s: %s", loja["key"], inicio, fim, erro)
                        s.save_debug(page, f"lactea_{loja['key']}_{inicio.replace('/', '')}")
                        falhas.append({"loja": loja["key"], "periodo": f"{inicio}–{fim}", "motivo": str(erro)[:200]})
                        continue
                    except Exception as erro:  # inesperado: registra e segue para a próxima loja
                        log.exception("%s %s–%s: falha inesperada", loja["key"], inicio, fim)
                        s.save_debug(page, f"lactea_{loja['key']}_{inicio.replace('/', '')}_erro")
                        falhas.append({"loja": loja["key"], "periodo": f"{inicio}–{fim}",
                                       "motivo": f"inesperado: {erro}"[:200]})
                        continue
                    texto = "sem registros" if r["sem_registros"] else ", ".join(
                        f"{i.sabor} {i.litros:.4f} {i.unidade}" for i in r["itens"]) or "nenhum sabor"
                    log.info("%s %s–%s (%s): %s [%.0fs]", loja["key"], inicio, fim, r["franquia"], texto,
                             time.time() - t0)
                    resumo.append({"loja": loja["key"], "de": inicio, "ate": fim, "sem_registros": r["sem_registros"],
                                   **{i.sabor: i.litros for i in r["itens"]},
                                   "unidade": r["itens"][0].unidade if r["itens"] else None})
                    if cliente is not None:
                        if cliente.enviar_lactea(loja["key"], fim, r["itens"], r["sem_registros"]):
                            ok.append(loja["key"])
                        else:
                            falhas.append({"loja": loja["key"], "periodo": fim, "motivo": "banco recusou o envio"})
                    else:
                        ok.append(loja["key"])
        finally:
            context.close()
            browser.close()

    print("RESUMO_JSON " + json.dumps(resumo, ensure_ascii=False))
    if cliente is not None:
        status = "sucesso" if not falhas else ("parcial" if ok else "falha")
        cliente.registrar_execucao("lactea", iniciado_em, status, len(set(ok)), len({f["loja"] for f in falhas}),
                                   falhas, f"{de} a {ate}: {len(resumo)} leitura(s), {len(falhas)} falha(s)")
    log.info("fim: %d leitura(s), %d falha(s)", len(resumo), len(falhas))
    if falhas:
        raise SystemExit(f"{len(falhas)} leitura(s) com falha: {json.dumps(falhas, ensure_ascii=False)}")


if __name__ == "__main__":
    main()
