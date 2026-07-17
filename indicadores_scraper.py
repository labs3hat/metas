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

Roda via GitHub Actions diariamente (ver .github/workflows/indicadores_daily.yml).

Comportamento auto-corretivo (backfill)
───────────────────────────────────────
A cada execução o scraper lê a planilha e verifica os últimos BACKFILL_DAYS
dias (padrão 7). Coleta APENAS as combinações loja+dia com célula vazia:
  - execução normal: só D-1 das 10 lojas;
  - se algum dia anterior falhou, ele é recuperado automaticamente;
  - células já preenchidas nunca são recoletadas (idempotente).
Se ao final restarem pendências, o processo termina com exit code 1 —
o workflow fica vermelho e o watchdog dispara o reprocessamento.

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

# Janela de auto-backfill: verifica os últimos N dias na planilha e coleta
# apenas as combinações loja+dia que estiverem vazias (auto-corretivo).
BACKFILL_DAYS = int(os.environ.get("BACKFILL_DAYS", "7"))


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


def _date_info(d: datetime) -> dict:
    """Empacota uma data no formato usado pelo restante do código."""
    return {
        "date": d,
        "day": d.day,
        "month": d.month,
        "year": d.year,
        "start": d.strftime("%d/%m/%Y 00:00"),
        "end":   d.strftime("%d/%m/%Y 23:59"),
        "label": d.strftime("%d/%m/%Y"),
    }


def build_dates_window(days: int) -> list[dict]:
    """Retorna a janela D-1..D-N (mais recente primeiro) no fuso de São Paulo."""
    now_br = br_now()
    return [_date_info(now_br - timedelta(days=i)) for i in range(1, days + 1)]


# ── DATACLASS DE RESULTADO ────────────────────────────────────────────────────
@dataclass
class StoreMetrics:
    store_key: str
    date: dict           # date_info do dia coletado
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
    # Marca as tabelas atuais como "stale". Quando o PrimeFaces renderizar o
    # novo resultado, a tabela é substituída e a marca desaparece — assim
    # _wait_tm_table distingue o resultado novo do anterior (essencial ao
    # pesquisar várias datas na mesma sessão).
    try:
        page.evaluate("""() => {
            document.querySelectorAll('table, .ui-datatable').forEach(el => { el.__stale = true; });
        }""")
    except Exception:
        pass
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
                const tables = Array.from(document.querySelectorAll('table, .ui-datatable'));
                return {
                    hasTotais:  body.includes('Totais'),
                    hasTicket:  body.includes('Ticket Médio Líquido') || body.includes('Ticket Medio Liquido'),
                    hasPessoas: body.includes('Pessoas Atendidas'),
                    hasNenhum:  body.includes('Nenhum registro'),
                    // Alguma tabela sem a marca __stale = resultado novo renderizado
                    hasFresh:   tables.length === 0 || tables.some(el => !el.__stale),
                };
            }""")
            fresh = state.get("hasFresh", True)
            if fresh and state.get("hasNenhum"):
                return True  # dia sem vendas: resultado novo, vazio
            if fresh and state.get("hasTotais") and (state.get("hasTicket") or state.get("hasPessoas")):
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
def collect_store_days(
    page, store: dict, tmpdir: str, dates: list[dict]
) -> tuple[list[StoreMetrics], list[dict]]:
    """
    Para a loja já logada e selecionada, coleta N datas na mesma sessão:
    abre o relatório uma vez e, para cada data, preenche o período,
    pesquisa, exporta o XLS e faz parse.

    Retorna (métricas coletadas, datas que falharam).
    """
    key = store["key"]

    go_to_tm_report(page)

    # Verifica que a loja ativa continua correta após navegação
    if not _active_store_matches(page, store["bip_name"]):
        raise RuntimeError(
            f"Loja errada no relatório. Esperado={store['bip_name']}, "
            f"Ativo={_get_active_store(page)}"
        )

    results: list[StoreMetrics] = []
    failed_dates: list[dict] = []

    for d in dates:
        try:
            log.info("  [%s] Coletando %s...", key, d["label"])
            _set_dates(page, d["start"], d["end"])
            time.sleep(0.8)
            _click_pesquisar(page)

            if not _wait_tm_table(page):
                log.warning("  TM table not fully confirmed; trying XLS anyway.")

            try:
                download = _click_xls(page)
            except Exception:
                _save_debug(page, f"xls_not_found_{key}_{d['year']}{d['month']:02d}{d['day']:02d}")
                raise

            file_path = os.path.join(
                tmpdir, f"{key}_tm_{d['year']}{d['month']:02d}{d['day']:02d}.xls"
            )
            download.save_as(file_path)
            log.info("  Downloaded: %s (%d bytes)", file_path, os.path.getsize(file_path))

            metrics = parse_tm_xls(file_path)
            results.append(StoreMetrics(
                store_key=key,
                date=d,
                receita=metrics["receita"],
                ticket_medio=metrics["ticket_medio"],
                pessoas=metrics["pessoas"],
            ))
            log.info(
                "  ✓ %s %s: receita=%.2f TM=%.2f pessoas=%d",
                key, d["label"], metrics["receita"], metrics["ticket_medio"], metrics["pessoas"],
            )
        except Exception as exc:
            failed_dates.append(d)
            log.error("  ✗ %s %s: %s", key, d["label"], exc)

    return results, failed_dates


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


def _col_letter(n: int) -> str:
    """Converte índice de coluna 1-indexed em letra A1 (3 → 'C')."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def read_missing_combos(dates: list[dict], ws_cache: dict) -> dict[str, list[dict]]:
    """
    Lê a planilha e retorna {store_key: [date_info, ...]} das combinações
    loja+dia em que qualquer uma das 3 métricas está vazia.

    Faz no máximo 1 batch_get (3 ranges) por mês envolvido na janela —
    tipicamente 1-2 chamadas de API no total.
    """
    from collections import defaultdict

    by_month: dict[tuple[int, int], list[dict]] = defaultdict(list)
    for d in dates:
        by_month[(d["year"], d["month"])].append(d)

    missing: dict[str, list[dict]] = defaultdict(list)

    for (year, month), dlist in sorted(by_month.items()):
        if year not in ws_cache:
            ws_cache[year] = _get_worksheet(year)
        ws = ws_cache[year]

        offset = _row_offset_for_month(month, year)
        days   = sorted(d["day"] for d in dlist)
        c1, c2 = _col_for_day(days[0]), _col_for_day(days[-1])

        ranges = []
        for base in (ROW_BASE_RECEITA, ROW_BASE_TM, ROW_BASE_PA):
            r1 = base + offset
            r2 = r1 + len(LOJA_ORDER) - 1
            ranges.append(f"{_col_letter(c1)}{r1}:{_col_letter(c2)}{r2}")

        blocks = ws.batch_get(ranges)

        for d in dlist:
            day_off = _col_for_day(d["day"]) - c1
            for idx, key in enumerate(LOJA_ORDER):
                vals = []
                for block in blocks:
                    row = block[idx] if idx < len(block) else []
                    v = row[day_off] if day_off < len(row) else ""
                    vals.append(str(v).strip())
                if any(v == "" for v in vals):
                    missing[key].append(d)

    return dict(missing)


def write_all_metrics(all_metrics: list[StoreMetrics], ws_cache: dict) -> int:
    """
    Escreve todas as métricas coletadas em lote — uma única chamada
    update_cells por aba de ano. Retorna o número de erros de escrita.
    """
    from collections import defaultdict
    from gspread.cell import Cell

    by_year: dict[int, list[StoreMetrics]] = defaultdict(list)
    for m in all_metrics:
        by_year[m.date["year"]].append(m)

    errors = 0
    for year, items in sorted(by_year.items()):
        try:
            if year not in ws_cache:
                ws_cache[year] = _get_worksheet(year)
            ws = ws_cache[year]

            cells: list[Cell] = []
            for m in items:
                offset   = _row_offset_for_month(m.date["month"], year)
                loja_idx = _loja_row_index(m.store_key)
                col      = _col_for_day(m.date["day"])
                cells.append(Cell(ROW_BASE_RECEITA + offset + loja_idx, col, round(m.receita, 2)))
                cells.append(Cell(ROW_BASE_TM      + offset + loja_idx, col, round(m.ticket_medio, 2)))
                cells.append(Cell(ROW_BASE_PA      + offset + loja_idx, col, m.pessoas))
                log.info(
                    "  → %s %s: receita=%.2f TM=%.2f pessoas=%d",
                    m.store_key, m.date["label"], m.receita, m.ticket_medio, m.pessoas,
                )

            ws.update_cells(cells, value_input_option="USER_ENTERED")
            log.info("  Gravadas %d células na aba de %d em 1 chamada.", len(cells), year)
        except Exception as exc:
            errors += 1
            log.error("  Erro ao gravar lote do ano %d: %s", year, exc)

    return errors


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    window = build_dates_window(BACKFILL_DAYS)
    log.info(
        "=== BIP360 Indicadores Scraper — janela %s..%s (%d dias) ===",
        window[-1]["label"], window[0]["label"], BACKFILL_DAYS,
    )

    # 1. Lê a planilha e detecta o que falta (idempotente / auto-corretivo)
    ws_cache: dict[int, gspread.Worksheet] = {}
    missing = read_missing_combos(window, ws_cache)

    total_pend = sum(len(v) for v in missing.values())
    if total_pend == 0:
        log.info("Planilha completa nos últimos %d dias. Nada a coletar.", BACKFILL_DAYS)
        return

    for key, dlist in missing.items():
        log.info("Pendências %s: %s", key, ", ".join(d["label"] for d in dlist))
    log.info("Total de pendências: %d combinações loja+dia", total_pend)

    # 2. Coleta apenas o que falta — 1 login por loja, N datas por sessão,
    #    com uma segunda tentativa (sessão nova) para as datas que falharem.
    all_metrics: list[StoreMetrics] = []
    unresolved: list[tuple[str, str, str]] = []  # (store_key, date_label, motivo)

    with tempfile.TemporaryDirectory() as tmpdir:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)

            for store in LOJAS:
                pending_dates = missing.get(store["key"])
                if not pending_dates:
                    continue

                log.info("\n--- %s : %s (%d dia(s) pendente(s)) ---",
                         store["key"].upper(), store["bip_name"], len(pending_dates))

                for attempt in (1, 2):
                    if not pending_dates:
                        break
                    if attempt == 2:
                        log.info("  Retry (sessão nova) para %d data(s) de %s...",
                                 len(pending_dates), store["key"])

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

                        results, failed = collect_store_days(
                            page, store, tmpdir, pending_dates
                        )
                        all_metrics.extend(results)
                        pending_dates = failed

                    except PlaywrightTimeout as exc:
                        log.error("  Timeout %s: %s", store["key"], exc)
                        if page:
                            _save_debug(page, f"timeout_{store['key']}")
                    except Exception as exc:
                        log.error("  Erro %s: %s", store["key"], exc)
                        if page:
                            _save_debug(page, f"error_{store['key']}")
                    finally:
                        if context:
                            context.close()

                for d in pending_dates:
                    unresolved.append((store["key"], d["label"], "falha na coleta após retry"))

            browser.close()

    # 3. Escrita em lote (1 chamada de API por aba de ano)
    log.info("\n--- Escrevendo no Google Sheets ---")
    sheet_errors = 0
    if all_metrics:
        sheet_errors = write_all_metrics(all_metrics, ws_cache)

    # 4. Resumo e exit code — pendência restante = falha do workflow,
    #    para que o watchdog dispare o reprocessamento automático.
    log.info(
        "\n=== Concluído: coletadas=%d | pendentes=%d | sheets=%d erro(s) de lote ===",
        len(all_metrics), len(unresolved), sheet_errors,
    )

    if unresolved:
        for key, label, motivo in unresolved:
            log.error("  PENDENTE: %s %s (%s)", key, label, motivo)

    if unresolved or sheet_errors:
        raise RuntimeError(
            f"{len(unresolved)} combinação(ões) loja+dia sem dados e "
            f"{sheet_errors} erro(s) de escrita. O watchdog irá reprocessar; "
            f"as células já preenchidas não serão recoletadas."
        )


if __name__ == "__main__":
    main()
