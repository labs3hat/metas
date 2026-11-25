"""
Extração AVULSA de venda por produto do BIP360 (André, 25/11/2026), para a
curva ABC da precificação (sessão "CMV - Precificação lojas").

NÃO é coleta: não grava no Supabase nem na planilha, não muda o robô diário.
Só lê o relatório "Venda de Produto por Operador" (finRelVendaOperadorPDV2.jsf)
mês a mês, com o Índice "Quantidade" e depois "Valor", e escreve CSV.

Reaproveita do scraper.py o login, a troca de loja, a navegação, as datas e a
exportação — o que já foi validado em produção (docs do coletor). Nenhuma
função do scraper.py é alterada.

Saída (pasta SAIDA, padrão "saida/"):
  venda_produto_<loja>.csv    loja, mes, produto, quantidade, valor
  situacao_<loja>.csv         loja, mes, situacao, linhas, detalhe
Uma linha por loja + mês + produto: os operadores são somados.

Variáveis:
  LOJAS   chaves separadas por vírgula (cl2,ctba3,...) ou "todas"
  DE      primeiro mês, AAAA-MM (padrão 2025-11)
  ATE     último dia, AAAA-MM-DD (padrão: ontem em São Paulo)
  SAIDA   pasta de saída
  BIP_USER / BIP_PASS (segredos do repositório; nunca vão para o log)

Falha: cada mês tem 2 tentativas, a segunda com sessão nova. Mês que falhar
nas duas vira "erro" no arquivo de situação (não some calado, não vira zero);
mês sem venda ("nenhum registro") vira "sem_venda". O processo termina com
código 1 se algum mês ficou com erro.
"""

import csv
import os
import re
import sys
import tempfile
import time
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

# O scraper.py lê estas variáveis ao ser importado; a extração não usa planilha.
os.environ.setdefault("GOOGLE_CREDENTIALS", "{}")

from playwright.sync_api import sync_playwright  # noqa: E402

import scraper as s  # noqa: E402

TZ = ZoneInfo("America/Sao_Paulo")


def log(msg: str) -> None:
    print(f"{datetime.now(TZ).strftime('%Y-%m-%d %H:%M:%S')} INFO [extracao] {msg}", flush=True)


def meses(de: str, ate: date) -> list[tuple[str, date, date]]:
    """[(AAAA-MM, primeiro dia, último dia)] de `de` até o mês de `ate`, o último cortado em `ate`."""
    ano, mes = (int(x) for x in de.split("-"))
    lista = []
    while (ano, mes) <= (ate.year, ate.month):
        ini = date(ano, mes, 1)
        prox = date(ano + (mes == 12), 1 if mes == 12 else mes + 1, 1)
        fim = min(prox - timedelta(days=1), ate)
        lista.append((f"{ano:04d}-{mes:02d}", ini, fim))
        ano, mes = (ano + 1, 1) if mes == 12 else (ano, mes + 1)
    return lista


def exportar(page, loja: dict, pasta: str, ini: date, fim: date, indice: str) -> str | None:
    """Exporta o mês com o Índice pedido. None = "nenhum registro"."""
    s.go_to_report(page)
    if not s.active_store_matches(page, loja["bip_name"]):
        raise Exception(f"Loja ativa errada no relatório: {s.get_active_store_text(page)}")
    # O Índice ANTES das datas: trocá-lo re-renderiza o formulário.
    s.selecionar_indice(page, indice)
    d_ini, d_fim = ini.strftime("%d/%m/%Y 00:00:00"), fim.strftime("%d/%m/%Y 23:59:59")
    s.set_report_dates(page, d_ini, d_fim)
    time.sleep(0.5)
    s.click_pesquisar(page)  # obrigatório depois de trocar o Índice
    s.wait_bip_idle(page, timeout=90)

    prazo = time.time() + 150
    estado: dict = {}
    while time.time() < prazo:
        estado = page.evaluate("""() => {
            const body = document.body.innerText || '';
            const valor = (id) => { const el = document.getElementById(id); return el ? el.value : null; };
            const linhas = Array.from(document.querySelectorAll('table tbody tr'))
                .map(r => (r.innerText || '').trim())
                .filter(t => t.length > 0 && !/nenhum registro/i.test(t)).length;
            return {
                linhas,
                hasXls: !!document.querySelector('img[src*="xls" i], img[src*="excel" i]'),
                hasNoRecords: body.toLowerCase().includes('nenhum registro'),
                inicio: valor('form:inputDataInicial_input'),
                fim: valor('form:inputDataFinal_input'),
            };
        }""")
        if estado.get("hasXls") and estado.get("linhas", 0) > 0:
            break
        if estado.get("hasNoRecords") and estado.get("linhas", 0) == 0:
            break
        time.sleep(1)
    log(f"  {indice}: {estado.get('inicio')!r} a {estado.get('fim')!r} | linhas na tela="
        f"{estado.get('linhas')} xls={estado.get('hasXls')} nenhum_registro={estado.get('hasNoRecords')}")
    if estado.get("inicio") is not None and str(estado["inicio"])[:10] != d_ini[:10]:
        raise Exception(f"A busca saiu com a data {estado['inicio']!r}, esperado {d_ini[:10]}.")
    if not (estado.get("hasXls") and estado.get("linhas", 0) > 0):
        if estado.get("hasNoRecords"):
            return None
        raise Exception(f"Tabela de {indice} não ficou pronta: {estado}")

    download = s.click_xls_export(page)
    caminho = os.path.join(pasta, f"{loja['key']}_{ini:%Y%m}_{s.slug(indice)}.xls")
    download.save_as(caminho)
    if not os.path.exists(caminho) or os.path.getsize(caminho) == 0:
        raise Exception(f"Arquivo exportado vazio: {caminho}")
    return caminho


def ler_por_produto(caminho: str, coluna: str) -> dict[str, float]:
    """Produto → soma da coluna ("quantidade" ou "valor") de todos os operadores.

    Linha idêntica repetida (mesmo operador, produto e número) entra uma vez,
    como no robô. Cabeçalho sem a coluna esperada é erro: é o sintoma de
    exportar com o Índice errado.
    """
    wb = s.open_workbook_smart(caminho)
    ws = wb.sheets()[0]
    cab = c_op = c_prod = c_num = None
    alvos = ("quantidade", "qtd") if coluna == "quantidade" else ("valor",)
    for r in range(min(20, ws.nrows)):
        vals = [s.normalize(str(ws.cell_value(r, c))).strip() for c in range(ws.ncols)]
        if "operador" in vals and "produto" in vals and any(a in vals for a in alvos):
            cab, c_op, c_prod = r, vals.index("operador"), vals.index("produto")
            c_num = next(vals.index(a) for a in alvos if a in vals)
            break
    if cab is None:
        raise Exception(f'Cabeçalho Operador/Produto/{coluna} não encontrado em {caminho}.')

    por_produto: dict[str, float] = {}
    vistas = set()
    repetidas, soma_repetidas = 0, 0.0
    for r in range(cab + 1, ws.nrows):
        produto = re.sub(r"\s+", " ", str(ws.cell_value(r, c_prod)).strip())
        if not produto or s.normalize(produto) in ("produto", "total", "totais"):
            continue
        operador = s.normalize_operator_name(ws.cell_value(r, c_op))
        if re.match(r"^tota(l|is)\b", s.normalize(operador)):
            continue
        n = s.parse_qty_cell(ws.cell_value(r, c_num))
        if n != n or n < 0:
            raise Exception(f'{coluna} inválido em {caminho}, linha {r + 1}: {ws.cell_value(r, c_num)!r}')
        chave = (operador.lower(), produto.lower(), round(n, 4))
        if chave in vistas:
            repetidas += 1
            soma_repetidas += n
            if os.environ.get("SOMAR_REPETIDAS") != "1":
                continue
        vistas.add(chave)
        por_produto[produto] = por_produto.get(produto, 0.0) + n
    if repetidas:
        log(f"  {coluna}: {repetidas} linha(s) repetida(s) (mesmo operador, produto e número), "
            f"somando {soma_repetidas:.2f}; "
            f"{'somadas' if os.environ.get('SOMAR_REPETIDAS') == '1' else 'descartadas'}")
    return por_produto


def extrair_mes(page, loja: dict, pasta: str, ini: date, fim: date):
    """(situacao, linhas) do mês: linhas = [(produto, quantidade, valor)]."""
    arq_q = exportar(page, loja, pasta, ini, fim, "Quantidade")
    arq_v = exportar(page, loja, pasta, ini, fim, "Valor")
    if arq_q is None and arq_v is None:
        return "sem_venda", []
    if (arq_q is None) != (arq_v is None):
        raise Exception(f"Um Índice veio vazio e o outro não (quantidade={arq_q}, valor={arq_v}).")
    qtd = ler_por_produto(arq_q, "quantidade")
    val = ler_por_produto(arq_v, "valor")
    produtos = sorted(set(qtd) | set(val))
    return "ok", [(p, round(qtd.get(p, 0.0), 3), round(val.get(p, 0.0), 2)) for p in produtos]


def main() -> int:
    hoje = datetime.now(TZ).date()
    ate = date.fromisoformat(os.environ.get("ATE") or (hoje - timedelta(days=1)).isoformat())
    de = os.environ.get("DE") or "2025-11"
    saida = os.environ.get("SAIDA") or "saida"
    pedidas = os.environ.get("LOJAS", "todas").strip().lower()
    # Loja ainda sem abertura (MGA 10) fica fora, como no robô.
    lojas = [l for l in s.LOJAS
             if ("abertura" not in l or l["abertura"])
             and (pedidas == "todas" or l["key"] in pedidas.split(","))]
    if not lojas:
        raise SystemExit(f"Nenhuma loja reconhecida em LOJAS={pedidas!r}.")
    os.makedirs(saida, exist_ok=True)
    periodo = meses(de, ate)
    log(f"início: {len(lojas)} loja(s), {len(periodo)} mês(es), {periodo[0][1]} a {periodo[-1][2]}")

    erros = 0
    inicio = time.time()
    with tempfile.TemporaryDirectory() as pasta, sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        for loja in lojas:
            log(f"--- {loja['key']}: {loja['bip_name']} ---")
            situacao, vendas = [], []
            contexto = page = None

            def sessao_nova():
                nonlocal contexto, page
                if contexto:
                    contexto.close()
                contexto = browser.new_context(accept_downloads=True,
                                               viewport={"width": 1600, "height": 1000})
                page = contexto.new_page()
                page.set_default_timeout(30000)
                s.login(page)
                s.select_store(page, loja["bip_name"])

            try:
                sessao_nova()
            except Exception as e:
                log(f"  ✗ loja indisponível no BIP: {e}")
                situacao.extend((loja["key"], m, "loja_indisponivel", 0, str(e)[:200]) for m, _, _ in periodo)
                erros += 1
            else:
                for mes, ini, fim in periodo:
                    for tentativa in (1, 2):
                        try:
                            if tentativa == 2:
                                log(f"  ↻ {mes}: tentando de novo com sessão nova")
                                sessao_nova()
                            estado, linhas = extrair_mes(page, loja, pasta, ini, fim)
                            vendas.extend((loja["key"], mes, p, q, v) for p, q, v in linhas)
                            situacao.append((loja["key"], mes, estado, len(linhas), ""))
                            log(f"  {mes}: {estado}, {len(linhas)} produto(s), "
                                f"qtd {sum(q for _, q, _ in linhas):.0f}, R$ {sum(v for *_, v in linhas):.2f}")
                            break
                        except Exception as e:
                            log(f"  ⚠️ {mes} (tentativa {tentativa}): {e}")
                            if tentativa == 2:
                                situacao.append((loja["key"], mes, "erro", 0, str(e)[:200]))
                                erros += 1
                                try:
                                    sessao_nova()
                                except Exception as e2:
                                    log(f"  ⚠️ sessão nova falhou: {e2}")
            finally:
                if contexto:
                    contexto.close()

            with open(os.path.join(saida, f"venda_produto_{loja['key']}.csv"), "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["loja", "mes", "produto", "quantidade", "valor"])
                w.writerows(vendas)
            with open(os.path.join(saida, f"situacao_{loja['key']}.csv"), "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["loja", "mes", "situacao", "linhas", "detalhe"])
                w.writerows(situacao)
        browser.close()

    log(f"fim em {time.time() - inicio:.0f} s: {erros} mês(es) ou loja(s) com erro")
    return 1 if erros else 0


if __name__ == "__main__":
    sys.exit(main())
