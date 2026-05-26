"""
BIP360 → Google Sheets — Indicadores Operacionais Diários
=========================================================
Coleta do relatório "Ticket Médio Diário" (finRelTicketMedioDiarioPDV2.jsf):
  - Receita Total Líquida (R$)
  - Ticket Médio Líquido (R$)
  - Pessoas Atendidas

Escreve na aba "Indicadores de Venda YYYY" do ano corrente de D-1:
  - 2026 → aba "Indicadores de Venda 2026"
  - 2027 → aba "Indicadores de Venda 2027" (crie a aba no Sheets antes de jan/2027)
  A troca de aba é automática — o código deriva o nome do ano de D-1.

Roda via GitHub Actions todo dia às 05:00 BRT (08:00 UTC).

Estrutura da planilha por mês (confirmada lendo a planilha real)
─────────────────────────────────────────────────────────────────
Cada bloco de mês tem 3 seções empilhadas verticalmente:
  Seção A — Receitas Totais Líquidas (R$)   base jan/26 = linha 4
  Seção B — Ticket Médio Líquido (R$)       base jan/26 = linha 15
  Seção C — Pessoas Atendidas               base jan/26 = linha 26
Delta entre meses: 36 linhas. Coluna C = dia 1.
"""

import os
import re
import time
import json
import logging
import tempfile
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from dataclasses import dataclass

import xlrd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
import gspread
from google.oauth2.service_account import Credentials

# ── LOGGING ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ── CONFIG ────────────────────────────────────────────────────────────────────
BIP_LOGIN_URL = "https://itfgestor.com.br/ITFGestor/publico/login.jsf"
TM_REPORT_URL = "https://itfgestor.com.br/ITFGestor/finRelTicketMedioDiarioPDV2.jsf"
DASHBOARD_URL = "https://itfgestor.com.br/ITFGestor/dashboardPDV2.jsf"

BIP_USER = os.environ["BIP_USER"]
BIP_PASS = os.environ["BIP_PASS"]

SHEET_ID          = "1qU8Ny_OqoF4VrI0IU4JuuRvoNnmBOMSf9JkFs1h4PRY"
# Nome da aba derivado automaticamente do ano de D-1.
# Não precisa alterar este código na virada do ano —
# basta criar a nova aba no Sheets com o nome correspondente.
SHEET_TAB_PREFIX  = "Indicadores de Venda"


def get_sheet_tab(year: int) -> str:
    """Retorna o nome da aba do ano informado, ex: 'Indicadores de Venda 2026'."""
    return f"{SHEET_TAB_PREFIX} {year}"

GOOGLE_CREDS = json.loads(os.environ["GOOGLE_CREDENTIALS"])
DEBUG_DIR    = os.environ.get("DEBUG_DIR", "debug")

TZ_BR = ZoneInfo("America/Sao_Paulo")

# ── LOJAS ─────────────────────────────────────────────────────────────────────
# key: identificador interno usado para mapear linhas na planilha.
# bip_name: nome exato como aparece no dropdown do BIP360.
LOJAS = [
    {"key": "sjp1",   "bip_name": "SÃO JOSÉ DOS PINHAIS 01 - SHOPPING SÃO JOSÉ (PR)"},
    {"key": "ctba3",  "bip_name": "CURITIBA 03 - SHOPPING ESTAÇÃO CURITIBA (PR)"},
    {"key": "ctba5",  "bip_name": "CURITIBA 05 - JOCKEY PLAZA SHOPPING (PR)"},
    {"key": "cl2",    "bip_name": "CAMPO LARGO 02 - CITY CENTER OUTLET PREMIUM (PR)"},
    {"key": "ctba7",  "bip_name": "CURITIBA 07 (PR)"},
    {"key": "ctba9",  "bip_name": "MARINGÁ 08 - HAVAN (PR)"},  # posição ctba9 na planilha → loja MGA8 no BIP360
    {"key": "ctba11", "bip_name": "CURITIBA 11 - SHOPPING PALLADIUM (PR)"},
    {"key": "mga3",   "bip_name": "MARINGÁ 03 - SHOPPING AVENIDA CENTER MARINGÁ (PR)"},
    {"key": "mga5",   "bip_name": "MARINGÁ 05 - SHOPPING CIDADE MARINGÁ (PR)"},
    {"key": "mga7",   "bip_name": "MARINGÁ 07 - SHOPPING AVENIDA CENTER (PR)"},
]

# ── MAPA DE LINHAS NA PLANILHA ────────────────────────────────────────────────
# Estrutura confirmada lendo a planilha real (maio/2026):
#
#   Jan/26  cabeçalho=3   SJP1 receita=4   SJP1 TM=15  SJP1 pessoas=26
#   Fev/26  cabeçalho=39  SJP1 receita=40  SJP1 TM=51  SJP1 pessoas=62
#   Mar/26  cabeçalho=75  SJP1 receita=76  SJP1 TM=87  SJP1 pessoas=98
#   Delta entre meses: 36 linhas
#
# Ordem das lojas dentro de cada seção (linhas consecutivas):
#   +0 SJP1 | +1 CTBA3 | +2 CTBA5 | +3 CL2 | +4 CTBA7 | +5 CTBA9
#   +6 CTBA11 | +7 MGA3 | +8 MGA5 | +9 MGA7 | +10 TOTAL (não escrevemos)
#
# Coluna C (3, 1-indexed) = dia 1. Dia N = coluna (2 + N).

ROW_BASE_RECEITA = 4   # SJP1 receita em janeiro/2026 (1-indexed)
ROW_BASE_TM      = 15  # SJP1 ticket médio em janeiro/2026
ROW_BASE_PA      = 26  # SJP1 pessoas atendidas em janeiro/2026

LINHAS_POR_BLOCO_MES = 36  # delta jan→fev→mar confirmado na planilha

# Ordem exata das lojas dentro de cada bloco (0-indexed offset)
LOJA_ORDER = ["sjp1", "ctba3", "ctba5", "cl2", "ctba7", "ctba9", "ctba11", "mga3", "mga5", "mga7"]

# Coluna C (1-indexed = 3) = dia 1 do mês
COL_DIA_1 = 3  # coluna C


# ── DATA / PERÍODO ────────────────────────────────────────────────────────────
def br_now() -> datetime:
    return datetime.now(TZ_BR)


def get_d1_info():
    """Retorna informações sobre o dia D-1 no fuso de São Paulo."""
    now_br = br_now()
    d1 = now_br - timedelta(days=1)
    return {
        "date": d1,
        "day": d1.day,
        "month": d1.month,
        "year": d1.year,
        "start": d1.strftime("%d/%m/%Y 00:00"),
        "end":   d1.strftime("%d/%m/%Y 23:59"),
    }


# ── DATACLASS DE RESULTADO ────────────────────────────────────────────────────
@dataclass
class StoreMetrics:
    store_key: str
    receita: float       # Receitas Totais Líquidas (R$)
    ticket_medio: float  # Ticket Médio Líquido (R$)
    pessoas: int         # Pessoas Atendidas


# ── DEBUG ─────────────────────────────────────────────────────────────────────
def _ensure_debug_dir():
    os.makedirs(DEBUG_DIR, exist_ok=True)


def _save_debug(page, label: str):
    try:
        _ensure_debug_dir()
        base = os.path.join(DEBUG_DIR, _slug(label))
        page.screenshot(path=base + ".png", full_page=True, timeout=15000)
        with open(base + ".html", "w", encoding="utf-8") as f:
            f.write(page.content())
        log.info("Debug saved: %s.png / .html", base)
    except Exception as exc:
        log.warning("Could not save debug for %s: %s", label, exc)


def _print_snapshot(page, label: str):
    try:
        body = page.locator("body").inner_text(timeout=5000)
        body = re.sub(r"\s+", " ", body).strip()
        log.info("Snapshot [%s] url=%s", label, page.url)
        log.info("Body [%s] %s", label, body[:500])
    except Exception as exc:
        log.warning("Could not print snapshot %s: %s", label, exc)


def _slug(text: str) -> str:
    table = str.maketrans(
        "ÁÀÂÃÉÊÍÓÔÕÚÇáàâãéêíóôõúç",
        "AAAAEEIOOOUCaaaaeeiooouc",
    )
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", str(text).translate(table)).strip("_").lower()


# ── HELPERS BIP360 ────────────────────────────────────────────────────────────
def _wait_bip_idle(page, timeout: int = 45) -> bool:
    """Aguarda o loader do PrimeFaces/BIP360 sumir e o AJAX terminar."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            state = page.evaluate("""() => {
                const loader = document.querySelector('#j_idt17, .itf-load, .ui-dialog.itf-load');
                let loaderVisible = false;
                if (loader) {
                    const s = window.getComputedStyle(loader);
                    const r = loader.getBoundingClientRect();
                    loaderVisible = s.display !== 'none' && s.visibility !== 'hidden'
                                    && r.width > 0 && r.height > 0
                                    && loader.getAttribute('aria-hidden') !== 'true';
                }
                const jqActive = window.jQuery ? window.jQuery.active : 0;
                return { ready: document.readyState, loaderVisible, jqActive };
            }""")
            if (
                state.get("ready") == "complete"
                and not state.get("loaderVisible")
                and int(state.get("jqActive") or 0) == 0
            ):
                return True
        except Exception:
            pass
        time.sleep(0.5)
    log.warning("_wait_bip_idle timeout after %ds", timeout)
    return False


def _normalize(text: str) -> str:
    table = str.maketrans(
        "ÁÀÂÃÉÊÍÓÔÕÚÇáàâãéêíóôõúç",
        "AAAAEEIOOOUCaaaaeeiooouc",
    )
    return re.sub(r"\s+", " ", str(text or "").translate(table)).strip().upper()


def _store_code(name: str) -> str:
    n = _normalize(name)
    if "-" in n:
        return n.split("-")[0].strip()
    if "(" in n:
        return n.split("(")[0].strip()
    return n.strip()


def _get_active_store(page) -> str:
    """Retorna o nome da loja ativa.
    No layout do BIP360 ha dois span.name na topbar:
      [0] = grupo franqueador (ex: 'Grupo CHQ')
      [1] = loja ativa (ex: 'Curitiba 07 (PR)')
    Ignoramos o span do grupo e retornamos o da loja.
    """
    js = """() => {
        const spans = Array.from(document.querySelectorAll('.widgets-item span.name'));
        for (let i = spans.length - 1; i >= 0; i--) {
            const txt = (spans[i].innerText || spans[i].textContent || '').replace(/\\s+/g,' ').trim();
            if (txt && !txt.toLowerCase().includes('grupo')) return txt;
        }
        if (spans.length >= 2) return (spans[1].innerText || '').trim();
        if (spans.length === 1) return (spans[0].innerText || '').trim();
        return '';
    }"""
    try:
        return page.evaluate(js) or ""
    except Exception:
        return ""


def _active_store_matches(page, bip_name: str) -> bool:
    expected_code = _store_code(bip_name)
    active = _normalize(_get_active_store(page))
    if not active:
        return False
    if expected_code and expected_code in active:
        return True
    active_code = _store_code(active)
    exp_tokens = expected_code.split()
    act_tokens = active_code.split()
    if len(exp_tokens) >= 2 and len(act_tokens) >= 2 and exp_tokens[:2] == act_tokens[:2]:
        return True
    return False


# ── LOGIN ─────────────────────────────────────────────────────────────────────
def login(page):
    log.info("Logging in to BIP360...")
    page.goto(BIP_LOGIN_URL, wait_until="load", timeout=45000)
    page.wait_for_load_state("networkidle", timeout=20000)

    page.wait_for_selector('input[type="text"], input[type="email"]', timeout=20000)
    inputs = page.query_selector_all('input[type="text"], input[type="email"]')
    if not inputs:
        raise RuntimeError("Login input not found.")
    inputs[0].fill(BIP_USER)
    page.fill('input[type="password"]', BIP_PASS)

    try:
        page.locator("text=ENTRAR").first.click(timeout=15000)
    except Exception:
        page.click('button:has-text("ENTRAR"), button:has-text("Entrar")')

    deadline = time.time() + 45
    while time.time() < deadline:
        try:
            url = page.url
            body = page.locator("body").inner_text(timeout=3000)
            if "publico/login.jsf" not in url and "Página Inicial" in body:
                break
        except Exception:
            pass
        time.sleep(1)

    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    if "publico/login.jsf" in page.url:
        _save_debug(page, "login_failed")
        raise RuntimeError("Login failed: still on login page.")
    log.info("Logged in. URL: %s", page.url)


# ── SELEÇÃO DE LOJA ───────────────────────────────────────────────────────────
def _open_store_dropdown(page) -> str:
    """Retorna o nome da loja ativa ignorando o span do grupo franqueador."""
    js = """() => {
        const spans = Array.from(document.querySelectorAll('.widgets-item span.name'));
        for (let i = spans.length - 1; i >= 0; i--) {
            const txt = (spans[i].innerText || spans[i].textContent || '').replace(/\\s+/g,' ').trim();
            if (txt && !txt.toLowerCase().includes('grupo')) return txt;
        }
        if (spans.length >= 2) return (spans[1].innerText || '').trim();
        return spans.length ? (spans[0].innerText || '').trim() : 'not found';
    }"""
    return page.evaluate(js)


def _click_store_item(page, bip_name: str) -> bool:
    expected_code = _store_code(bip_name)
    js = """([targetName, targetCode]) => {
        const norm = s => String(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
                          .replace(/\\s+/g,' ').trim().toUpperCase();
        const wanted = norm(targetName);
        const code   = norm(targetCode);
        const lists = Array.from(document.querySelectorAll('.lista-empresas'));
        let anchors = lists.flatMap(l => Array.from(l.querySelectorAll('li a')));
        if (!anchors.length) anchors = Array.from(document.querySelectorAll('a.ui-commandlink, a'));
        const matches = [];
        for (const a of anchors) {
            const label = a.querySelector('label');
            const raw = (label ? label.innerText : a.innerText || a.textContent || '').replace(/\\s+/g,' ').trim();
            const txt = norm(raw);
            if (!txt) continue;
            if (txt === wanted || txt.includes(wanted) || txt.includes(code))
                matches.push({id: a.id || '', rawText: raw, len: raw.length});
        }
        matches.sort((a,b) => a.len - b.len);
        if (!matches.length) return {ok: false, reason: 'not_found'};
        const chosen = matches[0];
        if (!chosen.id) return {ok: false, reason: 'no_id', chosen};
        if (window.PrimeFaces && PrimeFaces.addSubmitParam) {
            const payload = {}; payload[chosen.id] = chosen.id;
            PrimeFaces.addSubmitParam('topBar', payload).submit('topBar');
            return {ok: true, method: 'PrimeFaces', chosen};
        }
        const el = document.getElementById(chosen.id);
        if (el) { el.click(); return {ok: true, method: 'click', chosen}; }
        return {ok: false, reason: 'el_not_found', chosen};
    }"""
    result = page.evaluate(js, [bip_name, expected_code])
    log.info("  store_click result: %s", result)
    if not result or not result.get("ok"):
        return False
    try:
        page.wait_for_load_state("load", timeout=25000)
    except Exception:
        pass
    _wait_bip_idle(page, timeout=45)
    deadline = time.time() + 45
    while time.time() < deadline:
        if _active_store_matches(page, bip_name):
            return True
        time.sleep(1)
    return False


def select_store(page, bip_name: str):
    page.wait_for_load_state("domcontentloaded", timeout=15000)
    _wait_bip_idle(page, timeout=45)
    time.sleep(1)

    if _active_store_matches(page, bip_name):
        log.info("  Store already active: %s", bip_name)
        return

    # No novo layout do BIP360 a lista de lojas ja esta no DOM — vai direto ao click.
    active_now = _open_store_dropdown(page)
    log.info("  Loja ativa atual: %s | Selecionando: %s", active_now, bip_name)

    last_error: str | None = None

    for attempt in range(1, 4):
        log.info("  Tentativa %d/3 de selecionar loja...", attempt)

        if _click_store_item(page, bip_name):
            # Aguarda o BIP360 recarregar a pagina apos o PrimeFaces submit
            try:
                page.wait_for_load_state("load", timeout=30000)
            except Exception:
                pass
            _wait_bip_idle(page, timeout=45)
            time.sleep(2)

            after = _open_store_dropdown(page)
            log.info("  Loja ativa apos selecao: %s", after)

            if _active_store_matches(page, bip_name):
                log.info("  Store confirmed: %s", bip_name)
                return

            last_error = f"Store selection unconfirmed after attempt {attempt}. Expected={bip_name}, active={after}"
        else:
            last_error = f"Could not click store '{bip_name}'"

        _save_debug(page, f"select_attempt_{attempt}_{_slug(bip_name)}")
        time.sleep(2)

    _save_debug(page, f"could_not_select_{_slug(bip_name)}")
    raise RuntimeError(last_error or f"Could not select store '{bip_name}'")


# ── RELATÓRIO TICKET MÉDIO DIÁRIO ─────────────────────────────────────────────
def go_to_tm_report(page):
    """Abre o relatório Ticket Médio Diário."""
    log.info("  Opening Ticket Médio Diário report...")
    try:
        page.goto(TM_REPORT_URL, wait_until="load", timeout=30000)
        time.sleep(4)
        _wait_bip_idle(page, timeout=45)
        body = page.locator("body").inner_text(timeout=5000)
        if "Ticket Médio" in body or "Data Inicial" in body:
            log.info("  TM report opened by direct URL ✔")
            return
        log.info("  Direct URL fallback to menu...")
    except Exception as exc:
        log.warning("  TM direct URL failed: %s", exc)

    page.goto(DASHBOARD_URL, wait_until="load", timeout=30000)
    _wait_bip_idle(page, timeout=45)
    time.sleep(2)
    page.locator("text=Financeiro").first.click(timeout=10000)
    time.sleep(1.2)
    page.locator("text=Relatórios").first.click(timeout=10000)
    time.sleep(1.2)
    page.locator("text=Ticket Médio Diário").first.click(timeout=10000)
    time.sleep(4)
    _wait_bip_idle(page, timeout=45)
    log.info("  TM report opened by menu ✔")


def _set_dates(page, start_date: str, end_date: str):
    """Preenche Data Inicial e Data Final nos inputs visíveis."""
    result = page.evaluate("""([s, e]) => {
        const isVisible = el => {
            const st = window.getComputedStyle(el);
            const r = el.getBoundingClientRect();
            return st.visibility !== 'hidden' && st.display !== 'none'
                   && r.width > 0 && r.height > 0 && !el.disabled && !el.readOnly;
        };
        const inputs = Array.from(document.querySelectorAll('input[type="text"]')).filter(isVisible);
        const dateLike = inputs.filter(i => {
            const blob = `${i.value} ${i.id} ${i.name} ${i.getAttribute('placeholder')||''}`.toLowerCase();
            return blob.includes('data') || blob.includes('date') || /\\d{2}\\/\\d{2}\\/\\d{4}/.test(i.value)
                   || inputs.indexOf(i) <= 2;
        });
        const targets = dateLike.length >= 2 ? dateLike.slice(0,2) : inputs.slice(0,2);
        if (targets.length < 2) return {ok: false, total: inputs.length};
        const set = (el, v) => {
            el.focus(); el.value = v;
            el.dispatchEvent(new Event('input',{bubbles:true}));
            el.dispatchEvent(new Event('change',{bubbles:true}));
            el.blur();
        };
        set(targets[0], s); set(targets[1], e);
        return {ok: true, filled: targets.map(i => ({id: i.id, value: i.value}))};
    }""", [start_date, end_date])
    log.info("  Date fill result: %s", result)
    if not result or not result.get("ok"):
        raise RuntimeError(f"Could not fill date fields: {result}")


def _click_pesquisar(page):
    try:
        page.locator("text=Pesquisar").first.click(timeout=15000)
        return
    except Exception:
        pass
    clicked = page.evaluate("""() => {
        const els = Array.from(document.querySelectorAll('button, a, span'));
        const visible = el => {
            const s = window.getComputedStyle(el); const r = el.getBoundingClientRect();
            return s.visibility !== 'hidden' && s.display !== 'none' && r.width > 0 && r.height > 0;
        };
        const found = els.find(el => visible(el) && (el.innerText||'').trim().includes('Pesquisar'));
        if (found) { found.click(); return found.innerText.trim(); }
        return null;
    }""")
    if not clicked:
        raise RuntimeError("Pesquisar button not found.")


def _wait_tm_table(page) -> bool:
    _wait_bip_idle(page, timeout=60)
    deadline = time.time() + 60
    while time.time() < deadline:
        try:
            state = page.evaluate("""() => {
                const body = document.body.innerText || '';
                return {
                    hasTotais:  body.includes('Totais'),
                    hasTicket:  body.includes('Ticket Médio Líquido') || body.includes('Ticket Medio Liquido'),
                    hasPessoas: body.includes('Pessoas Atendidas'),
                    hasXls:     !!document.querySelector('img[src*="xls" i], img[src*="excel" i]'),
                };
            }""")
            if state.get("hasTotais") and (state.get("hasTicket") or state.get("hasPessoas")):
                return True
        except Exception:
            pass
        time.sleep(1)
    return False


def _click_xls(page):
    """Clica no botão de exportação XLS e retorna o objeto download."""
    _wait_bip_idle(page, timeout=45)
    try:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1)
    except Exception:
        pass

    selectors = ['a:has(img[src*="xls"])', 'a:has(img[src*="excel"])',
                 'img[src*="xls"]', 'img[src*="excel"]']
    last_error = None

    for sel in selectors:
        try:
            loc = page.locator(sel).last
            if loc.count() == 0:
                continue
            with page.expect_download(timeout=90000) as dl:
                loc.click(timeout=15000, force=True)
            return dl.value
        except Exception as exc:
            last_error = exc

    # Fallback JS
    try:
        with page.expect_download(timeout=90000) as dl:
            result = page.evaluate("""() => {
                const imgs = Array.from(document.querySelectorAll('img[src*="xls" i], img[src*="excel" i]'));
                for (const img of imgs) {
                    const a = img.closest('a');
                    if (!a) continue;
                    a.scrollIntoView(); a.click();
                    return {ok: true, src: img.getAttribute('src')};
                }
                return {ok: false};
            }""")
            if not result or not result.get("ok"):
                raise RuntimeError(f"XLS button not found via JS: {result}")
        return dl.value
    except Exception as exc:
        last_error = exc

    raise RuntimeError(f"XLS export failed. Last error: {last_error}")


# ── PARSE DO XLS TICKET MÉDIO DIÁRIO ─────────────────────────────────────────
def parse_tm_xls(file_path: str) -> dict:
    """
    Lê o XLS do relatório Ticket Médio Diário e extrai da linha 'Totais':
      - receita:       Receitas Totais Líquidas (R$)
      - ticket_medio:  Ticket Médio Líquido (R$)
      - pessoas:       Pessoas Atendidas

    Colunas esperadas no relatório:
    Data | Pessoas Atendidas | Ticket Médio Líquido (R$) | Venda em Itens |
    Ticket Médio por Produto | Receitas Totais Líquidas (R$) | ...

    A ordem real é localizada pelo cabeçalho para resistir a mudanças futuras.
    """
    wb = xlrd.open_workbook(file_path)
    ws = wb.sheets()[0]

    # Localizar linha de cabeçalho e colunas
    col_receita = None
    col_tm      = None
    col_pessoas = None
    header_row  = None

    for r in range(min(15, ws.nrows)):
        vals = [str(ws.cell_value(r, c)).strip().lower() for c in range(ws.ncols)]
        for c, v in enumerate(vals):
            vn = _remove_accents(v)
            if "receita" in vn and "total" in vn and "liquida" in vn:
                col_receita = c
            if "ticket" in vn and "medio" in vn and "liquido" in vn:
                col_tm = c
            if "pessoa" in vn and "atendida" in vn:
                col_pessoas = c
        if col_receita is not None and col_tm is not None and col_pessoas is not None:
            header_row = r
            break

    # Fallbacks posicionais baseados no layout observado no scraper atual.
    # parse_ticket_medio_xls usa col 2 como fallback para Ticket Médio Líquido.
    # Ajuste se o relatório tiver colunas diferentes na sua versão.
    if col_tm is None:
        col_tm = 2
    if col_pessoas is None:
        col_pessoas = 1
    if col_receita is None:
        col_receita = 5  # posição típica de Receitas Totais Líquidas

    log.info(
        "  XLS headers: header_row=%s col_tm=%s col_pessoas=%s col_receita=%s",
        header_row, col_tm, col_pessoas, col_receita,
    )

    # Localizar linha 'Totais' (última ocorrência = totais finais do período)
    total_row = None
    for r in range(ws.nrows):
        row_text = " ".join(str(ws.cell_value(r, c)) for c in range(ws.ncols)).lower()
        if "totais" in row_text or "total" in row_text:
            total_row = r  # não dá break: queremos o último

    if total_row is None:
        # Fallback: última linha não vazia
        for r in range(ws.nrows - 1, -1, -1):
            if any(str(ws.cell_value(r, c)).strip() for c in range(ws.ncols)):
                total_row = r
                break

    if total_row is None:
        raise RuntimeError("Linha 'Totais' não encontrada no XLS.")

    def _parse_cell(row: int, col: int) -> float:
        raw = ws.cell_value(row, col)
        if isinstance(raw, (int, float)):
            return float(raw)
        s = re.sub(r"[^0-9,.\-]", "", str(raw or "").strip())
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s or 0)

    receita      = _parse_cell(total_row, col_receita)
    ticket_medio = _parse_cell(total_row, col_tm)
    pessoas      = int(_parse_cell(total_row, col_pessoas))

    log.info(
        "  Parsed totals: receita=%.2f ticket_medio=%.2f pessoas=%d",
        receita, ticket_medio, pessoas,
    )
    return {"receita": receita, "ticket_medio": ticket_medio, "pessoas": pessoas}


def _remove_accents(text: str) -> str:
    import unicodedata
    return "".join(
        c for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )


# ── COLETA POR LOJA ───────────────────────────────────────────────────────────
def collect_store(page, store: dict, tmpdir: str, d1: dict) -> StoreMetrics:
    """
    Para a loja já selecionada:
    1. Abre relatório Ticket Médio Diário
    2. Preenche data D-1
    3. Pesquisa
    4. Exporta XLS
    5. Faz parse das 3 métricas
    """
    start_date = d1["start"]
    end_date   = d1["end"]
    key        = store["key"]

    go_to_tm_report(page)

    # Verifica que a loja ativa continua correta após navegação
    if not _active_store_matches(page, store["bip_name"]):
        raise RuntimeError(
            f"Loja errada no relatório. Esperado={store['bip_name']}, "
            f"Ativo={_get_active_store(page)}"
        )

    log.info("  Setting TM dates: %s → %s", start_date, end_date)
    _set_dates(page, start_date, end_date)
    time.sleep(0.8)
    _click_pesquisar(page)

    table_ready = _wait_tm_table(page)
    if not table_ready:
        log.warning("  TM table not fully confirmed; trying XLS anyway.")

    _print_snapshot(page, f"after_tm_search_{key}")

    try:
        download = _click_xls(page)
    except Exception:
        _save_debug(page, f"xls_not_found_{key}")
        raise

    file_path = os.path.join(tmpdir, f"{key}_tm.xls")
    download.save_as(file_path)
    log.info("  Downloaded: %s (%d bytes)", file_path, os.path.getsize(file_path))

    metrics = parse_tm_xls(file_path)
    return StoreMetrics(
        store_key=key,
        receita=metrics["receita"],
        ticket_medio=metrics["ticket_medio"],
        pessoas=metrics["pessoas"],
    )


# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
def _get_worksheet(year: int) -> gspread.Worksheet:
    """Abre a aba do ano correspondente a D-1.
    Se a aba não existir, o gspread lança WorksheetNotFound com mensagem clara.
    """
    tab_name = get_sheet_tab(year)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(GOOGLE_CREDS, scopes=scopes)
    gc = gspread.authorize(creds)
    try:
        ws = gc.open_by_key(SHEET_ID).worksheet(tab_name)
        log.info("Aba conectada: '%s'", tab_name)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        raise RuntimeError(
            f"Aba '{tab_name}' não encontrada na planilha. "
            f"Crie a aba no Google Sheets antes de rodar o scraper para {year}."
        )


def _col_for_day(day: int) -> int:
    """Retorna o índice de coluna 1-indexed para o dia do mês.
    Coluna C (3) = dia 1, coluna D (4) = dia 2, etc.
    """
    return COL_DIA_1 + (day - 1)


def _row_offset_for_month(month: int, year: int) -> int:
    """
    Calcula o deslocamento de linha para o bloco do mês correto.
    Janeiro/2026 = offset 0 (linha base).
    Fevereiro/2026 = offset 36, Março/2026 = offset 72, etc.
    """
    REFERENCE_MONTH = 1
    REFERENCE_YEAR  = 2026
    total_months = (year - REFERENCE_YEAR) * 12 + (month - REFERENCE_MONTH)
    return total_months * LINHAS_POR_BLOCO_MES


def _loja_row_index(store_key: str) -> int:
    """Posição 0-indexed da loja em LOJA_ORDER."""
    if store_key not in LOJA_ORDER:
        raise ValueError(f"Loja desconhecida: {store_key}")
    return LOJA_ORDER.index(store_key)


def write_metrics_to_sheet(ws: gspread.Worksheet, metrics: StoreMetrics, d1: dict):
    """
    Escreve as 3 métricas na coluna do dia D-1 do mês correto.

    Linhas na planilha (1-indexed):
      Receita     = ROW_BASE_RECEITA + offset_mes + loja_idx
      Ticket Méd  = ROW_BASE_TM      + offset_mes + loja_idx
      Pessoas     = ROW_BASE_PA      + offset_mes + loja_idx
    """
    month = d1["month"]
    year  = d1["year"]
    day   = d1["day"]
    key   = metrics.store_key

    offset    = _row_offset_for_month(month, year)
    loja_idx  = _loja_row_index(key)
    col       = _col_for_day(day)

    row_receita = ROW_BASE_RECEITA + offset + loja_idx
    row_tm      = ROW_BASE_TM      + offset + loja_idx
    row_pa      = ROW_BASE_PA      + offset + loja_idx

    # Enviamos 3 updates individuais para evitar conflito de range
    # (cada métrica fica em seções não contíguas).
    # Receita: arredondada em 2 casas como número
    ws.update_cell(row_receita, col, round(metrics.receita, 2))
    ws.update_cell(row_tm,      col, round(metrics.ticket_medio, 2))
    ws.update_cell(row_pa,      col, metrics.pessoas)

    log.info(
        "  → %s dia=%02d: receita=%.2f TM=%.2f pessoas=%d "
        "(rows %d/%d/%d col %d)",
        key, day, metrics.receita, metrics.ticket_medio, metrics.pessoas,
        row_receita, row_tm, row_pa, col,
    )


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    d1 = get_d1_info()
    log.info("=== BIP360 Indicadores Scraper — D-1: %s/%s/%s ===",
             d1["day"], d1["month"], d1["year"])
    log.info("Período: %s → %s", d1["start"], d1["end"])

    ws = _get_worksheet(d1["year"])

    success_count = 0
    error_count   = 0
    all_metrics: list[StoreMetrics] = []

    with tempfile.TemporaryDirectory() as tmpdir:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)

            for store in LOJAS:
                log.info("\n--- %s : %s ---", store["key"].upper(), store["bip_name"])
                context = None
                page    = None

                try:
                    context = browser.new_context(
                        accept_downloads=True,
                        viewport={"width": 1600, "height": 1000},
                    )
                    page = context.new_page()
                    page.set_default_timeout(30000)

                    login(page)
                    select_store(page, store["bip_name"])

                    metrics = collect_store(page, store, tmpdir, d1)
                    all_metrics.append(metrics)
                    log.info(
                        "  ✓ %s: receita=%.2f TM=%.2f pessoas=%d",
                        store["key"], metrics.receita, metrics.ticket_medio, metrics.pessoas,
                    )
                    success_count += 1

                except PlaywrightTimeout as exc:
                    error_count += 1
                    log.error("  Timeout %s: %s", store["key"], exc)
                    if page:
                        _save_debug(page, f"timeout_{store['key']}")

                except Exception as exc:
                    error_count += 1
                    log.error("  Erro %s: %s", store["key"], exc)
                    if page:
                        _save_debug(page, f"error_{store['key']}")

                finally:
                    if context:
                        context.close()

            browser.close()

    # Escreve todos os resultados no Sheets em sequência (após fechar o browser)
    log.info("\n--- Escrevendo no Google Sheets ---")
    sheet_errors = 0
    for m in all_metrics:
        try:
            write_metrics_to_sheet(ws, m, d1)
        except Exception as exc:
            sheet_errors += 1
            log.error("  Erro ao gravar %s: %s", m.store_key, exc)

    # Resumo final
    log.info(
        "\n=== Concluído: browser=%d ok / %d erros | sheets=%d erros ===",
        success_count, error_count, sheet_errors,
    )

    if success_count == 0:
        raise RuntimeError(
            "Nenhuma loja processada com sucesso. Verifique os logs e artifacts de debug."
        )


if __name__ == "__main__":
    main()
