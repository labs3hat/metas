"""
BIP360 → Google Sheets scraper
Roda via GitHub Actions.
Fluxo confirmado no BIP360:
Login → selecionar loja → Relatório Venda de Produto por Operador → pesquisar → exportar XLS → atualizar Google Sheets.
"""

import os
import re
import time
import json
import tempfile
import shutil
import csv
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import xlrd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
import gspread
from google.oauth2.service_account import Credentials

# ── CONFIG ────────────────────────────────────────────────────────────────────
BIP_LOGIN_URL = "https://itfgestor.com.br/ITFGestor/publico/login.jsf"
REPORT_URL    = "https://itfgestor.com.br/ITFGestor/finRelVendaOperadorPDV2.jsf"
TM_REPORT_URL = "https://itfgestor.com.br/ITFGestor/finRelTicketMedioDiarioPDV2.jsf"

BIP_USER     = os.environ["BIP_USER"]
BIP_PASS     = os.environ["BIP_PASS"]
SHEET_ID     = "1qU8Ny_OqoF4VrI0IU4JuuRvoNnmBOMSf9JkFs1h4PRY"
SHEET_TAB    = "Metas Operacionais"
SHEET_TAB_TM = "Metas TM suporte"
SHEET_TAB_OPERADOR = "Metas Operador suporte"
SHEET_TAB_HIST     = "Historico Vendas"

# ── DATA / PERÍODO ─────────────────────────────────────────────────────────────
# GitHub Actions roda normalmente em UTC. Para D-1 operacional do BIP360,
# a referência correta é o horário do Brasil.
TZ_BR = ZoneInfo("America/Sao_Paulo")


def br_now() -> datetime:
    return datetime.now(TZ_BR)


def get_report_period():
    """Return business D-1 period based on America/Sao_Paulo.

    Exemplo:
    - se for 24/04 23:55 no Brasil, busca até 23/04 23:59:59
    - se for 25/04 00:03 no Brasil, busca até 24/04 23:59:59
    - se for 01/05, busca o mês anterior: 01/04 até 30/04
    """
    now_br = br_now()
    end_dt = now_br - timedelta(days=1)
    start_dt = end_dt.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    return now_br, start_dt, end_dt


def get_report_month_info():
    """Month/year used to write the consolidated D-1 result."""
    _now_br, start_dt, end_dt = get_report_period()
    return end_dt.month, end_dt.year


def get_report_date_strings(include_seconds: bool = True):
    _now_br, start_dt, end_dt = get_report_period()
    if include_seconds:
        start_date = start_dt.strftime("%d/%m/%Y 00:00:00")
        end_date = end_dt.strftime("%d/%m/%Y 23:59:59")
    else:
        start_date = start_dt.strftime("%d/%m/%Y 00:00")
        end_date = end_dt.strftime("%d/%m/%Y 23:59")
    return start_date, end_date

GOOGLE_CREDS = json.loads(os.environ["GOOGLE_CREDENTIALS"])
DEBUG_DIR    = os.environ.get("DEBUG_DIR", "debug")

# ── DEBUG / BLINDAGEM DE DOWNLOAD ─────────────────────────────────────────────
DEBUG_STORES = {
    s.strip().lower()
    for s in os.environ.get("DEBUG_STORES", "mga7,ctba3").split(",")
    if s.strip()
}
DEBUG_ALL = os.environ.get("DEBUG_ALL", "0") == "1"


def is_debug_store(store_key: str) -> bool:
    return DEBUG_ALL or str(store_key or "").lower() in DEBUG_STORES


def ensure_debug_dir():
    os.makedirs(DEBUG_DIR, exist_ok=True)


def clean_xls_files(folder: str):
    """Remove XLS/XLSX files from a folder before a fresh BIP360 export."""
    if not folder or not os.path.isdir(folder):
        return
    for name in os.listdir(folder):
        if name.lower().endswith((".xls", ".xlsx", ".crdownload", ".tmp")):
            try:
                os.remove(os.path.join(folder, name))
            except Exception:
                pass


def copy_xls_debug(file_path: str, store_key: str, suffix: str = "operadores"):
    """Keep a copy of the exact XLS processed by the scraper."""
    try:
        ensure_debug_dir()
        dest = os.path.join(DEBUG_DIR, f"{store_key}_{suffix}.xls")
        shutil.copyfile(file_path, dest)
        print(f"  🧾 Debug XLS saved: {dest}")
    except Exception as e:
        print(f"  ⚠️ Could not save debug XLS for {store_key}: {e}")


def parse_qty_cell(value) -> float:
    """Parse numeric cells from xlrd safely."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    s = s.replace("'", "").replace("’", "").replace("`", "").replace("´", "")
    s = re.sub(r"[^0-9,.\-]", "", s)

    if not s:
        return 0.0

    # Brazilian decimal style: 1.071,00
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    return float(s or 0)


def write_operator_parse_debug(store_key: str, rows: list[dict]):
    """Save row-by-row parse diagnostic CSV and print a compact summary."""
    if not is_debug_store(store_key):
        return

    ensure_debug_dir()
    path = os.path.join(DEBUG_DIR, f"{store_key}_parse_linhas.csv")

    try:
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["row", "operador", "produto", "quantidade", "categoria", "used"],
            )
            writer.writeheader()
            writer.writerows(rows)

        print(f"  🧾 Debug parse CSV saved: {path}")
    except Exception as e:
        print(f"  ⚠️ Could not save parse debug CSV for {store_key}: {e}")

    # Print row-level details in GitHub logs for the target debug stores.
    print(f"  🔎 Parse detail for {store_key}:")
    for r in rows:
        if r.get("used"):
            print(
                f"    row={r['row']} | {r['operador']} | {r['produto']} | "
                f"qtd={r['quantidade']} | cat={r['categoria']}"
            )


def print_operator_category_summary(store_key: str, operators: dict):
    """Print compact per-operator summary for validation."""
    print(f"  🔎 Operator summary for {store_key}:")
    for op, vals in sorted(operators.items(), key=lambda kv: kv[1].get("total", 0), reverse=True):
        print(
            f"    {op}: agua={int(vals.get('agua', 0))}, "
            f"chantilly={int(vals.get('chantilly', 0))}, "
            f"shake={int(vals.get('shake', 0))}, "
            f"milk={int(vals.get('milk', 0))}, "
            f"total={int(vals.get('total', 0))}"
        )



LOJAS = [
    {"bip_name": "CAMPO LARGO 02 - CITY CENTER OUTLET PREMIUM (PR)",  "key": "cl2"},
    {"bip_name": "CURITIBA 03 - SHOPPING ESTAÇÃO CURITIBA (PR)",       "key": "ctba3"},
    {"bip_name": "CURITIBA 05 - JOCKEY PLAZA SHOPPING (PR)",           "key": "ctba5"},
    {"bip_name": "CURITIBA 07 (PR)",                                    "key": "ctba7q"},
    {"bip_name": "CURITIBA 11 - SHOPPING PALLADIUM (PR)",              "key": "ctba10"},
    {"bip_name": "MARINGÁ 03 - SHOPPING AVENIDA CENTER MARINGÁ (PR)", "key": "mga3"},
    {"bip_name": "MARINGÁ 05 - SHOPPING CIDADE MARINGÁ (PR)",         "key": "mga5"},
    {"bip_name": "MARINGÁ 07 - SHOPPING AVENIDA CENTER (PR)",         "key": "mga7"},
    {"bip_name": "MARINGÁ 08 - HAVAN (PR)",                           "key": "mga8"},
    {"bip_name": "SÃO JOSÉ DOS PINHAIS 01 - SHOPPING SÃO JOSÉ (PR)",  "key": "sjp1"},
]

MESES_PT = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"
}

STORE_SHEET_ROW = {
    "sjp1": 3,
    "ctba3": 7,
    "ctba5": 11,
    "cl2": 15,
    "ctba7q": 19,
    "ctba10": 23,
    "mga3": 27,
    "mga5": 31,
    "mga7": 35,
    "mga8": 39,
}

# Aba "Metas TM suporte"
# Estrutura: cada mês ocupa 2 colunas: Meta TM / Realizado TM.
# Janeiro: B/C, Fevereiro: D/E, Março: F/G, Abril: H/I...
# Aqui usamos coluna inicial do mês como a coluna da META; o REALIZADO é +1.
TM_STORE_SHEET_ROW = {
    "sjp1": 3,
    "ctba3": 4,
    "ctba5": 5,
    "cl2": 6,
    "ctba7q": 7,
    "ctba10": 8,
    "mga3": 9,
    "mga5": 10,
    "mga7": 11,
    "mga8": 12,
}


CAT_OFFSET = {"shake": 0, "chantilly": 1, "agua": 2, "milk": 3, "canecake": 4}


# ── HELPERS ───────────────────────────────────────────────────────────────────
def safe_name(text: str) -> str:
    text = str(text)
    text = text.normalize("NFD") if hasattr(text, "normalize") else text
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", str(text)).strip("_").lower()


def slug(text: str) -> str:
    text = str(text)
    text = (
        text.replace("Á", "A").replace("À", "A").replace("Â", "A").replace("Ã", "A")
            .replace("É", "E").replace("Ê", "E")
            .replace("Í", "I")
            .replace("Ó", "O").replace("Ô", "O").replace("Õ", "O")
            .replace("Ú", "U")
            .replace("Ç", "C")
            .replace("á", "a").replace("à", "a").replace("â", "a").replace("ã", "a")
            .replace("é", "e").replace("ê", "e")
            .replace("í", "i")
            .replace("ó", "o").replace("ô", "o").replace("õ", "o")
            .replace("ú", "u")
            .replace("ç", "c")
    )
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", text).strip("_").lower()


def save_debug(page, label: str):
    try:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        base = os.path.join(DEBUG_DIR, slug(label))
        page.screenshot(path=base + ".png", full_page=True, timeout=15000)
        with open(base + ".html", "w", encoding="utf-8") as f:
            f.write(page.content())
        print(f"  🧪 Debug saved: {base}.png and {base}.html")
    except Exception as e:
        print(f"  ⚠️ Could not save debug for {label}: {e}")


def print_page_snapshot(page, label: str):
    try:
        title = page.title()
        url = page.url
        body_text = page.locator("body").inner_text(timeout=5000)
        body_text = re.sub(r"\s+", " ", body_text).strip()
        print(f"  🧭 Snapshot [{label}] title={title!r} url={url!r}")
        print(f"  🧭 Body [{label}] {body_text[:700]}")
    except Exception as e:
        print(f"  ⚠️ Could not print snapshot {label}: {e}")


def wait_bip_idle(page, timeout: int = 45):
    """Wait until PrimeFaces/BIP360 loader disappears and AJAX finishes."""
    deadline = time.time() + timeout
    last_state = None

    while time.time() < deadline:
        try:
            state = page.evaluate(
                """() => {
                    const loader = document.querySelector('#j_idt17, .itf-load, .ui-dialog.itf-load');
                    let loaderVisible = false;

                    if (loader) {
                        const style = window.getComputedStyle(loader);
                        const rect = loader.getBoundingClientRect();
                        loaderVisible =
                            style.display !== 'none' &&
                            style.visibility !== 'hidden' &&
                            rect.width > 0 &&
                            rect.height > 0 &&
                            loader.getAttribute('aria-hidden') !== 'true';
                    }

                    const jqActive = window.jQuery ? window.jQuery.active : 0;

                    return {
                        ready: document.readyState,
                        loaderVisible,
                        jqActive,
                        url: window.location.href
                    };
                }"""
            )
            last_state = state

            if state.get("ready") == "complete" and not state.get("loaderVisible") and int(state.get("jqActive") or 0) == 0:
                return True

        except Exception:
            pass

        time.sleep(0.5)

    print(f"  ⚠️ wait_bip_idle timeout. Last state: {last_state}")
    return False


def wait_active_store(page, bip_name: str, timeout: int = 45) -> bool:
    """Wait until the active store in the top bar matches expected."""
    deadline = time.time() + timeout

    while time.time() < deadline:
        try:
            if active_store_matches(page, bip_name):
                return True
        except Exception:
            pass

        time.sleep(1)

    return False


def fill_text_input(element, value: str):
    """Preenche input PrimeFaces/JSF com eventos de teclado reais."""
    element.click()
    time.sleep(0.2)
    element.press("Control+A")
    time.sleep(0.1)
    element.press("Backspace")
    time.sleep(0.1)
    element.type(value, delay=10)


def normalize_js_text_for_python(s: str) -> str:
    s = str(s or "")
    table = str.maketrans("ÁÀÂÃÉÊÍÓÔÕÚÇáàâãéêíóôõúç", "AAAAEEIOOOUCaaaaeeiooouc")
    return re.sub(r"\s+", " ", s.translate(table)).strip().upper()


def store_code_from_name(name: str) -> str:
    """Return stable store code prefix from BIP name, e.g. CURITIBA 03, MARINGA 08."""
    n = normalize_js_text_for_python(name)
    if "-" in n:
        return n.split("-")[0].strip()
    if "(" in n:
        return n.split("(")[0].strip()
    return n.strip()


def get_active_store_text(page) -> str:
    """Read only the active store shown in the top blue bar.

    Important: body text is not reliable because when dropdown is open it contains all stores.
    This function looks at visible elements near the top bar and returns the shortest text that
    looks like a store name.
    """
    js = """
    () => {
        const visible = (el) => {
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style &&
                   style.visibility !== 'hidden' &&
                   style.display !== 'none' &&
                   rect.width > 0 &&
                   rect.height > 0;
        };

        const els = Array.from(document.querySelectorAll('a, span, div, button'));
        const candidates = [];

        for (const el of els) {
            if (!visible(el)) continue;
            const rect = el.getBoundingClientRect();
            const txt = (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim();
            if (!txt) continue;

            // Top blue bar is roughly top 0-170 px.
            if (rect.top > 180) continue;

            const looksStore =
                txt.includes('SHOPPING') ||
                txt.includes('Shopping') ||
                txt.includes('SÃO JOSÉ') ||
                txt.includes('SAO JOSE') ||
                txt.includes('CURITIBA') ||
                txt.includes('MARING') ||
                txt.includes('CAMPO LARGO');

            if (!looksStore) continue;
            if (txt.includes('MINHAS FRANQUIAS')) continue;
            if (txt.includes('Página Inicial')) continue;
            if (txt.includes('Faturamento')) continue;

            candidates.push({
                text: txt,
                top: rect.top,
                left: rect.left,
                width: rect.width,
                len: txt.length
            });
        }

        candidates.sort((a, b) => {
            // Prefer top bar and shorter exact label.
            if (Math.abs(a.top - b.top) > 20) return a.top - b.top;
            return a.len - b.len;
        });

        return candidates.length ? candidates[0].text : '';
    }
    """
    try:
        return page.evaluate(js) or ""
    except Exception:
        return ""


def active_store_matches(page, bip_name: str) -> bool:
    expected_full = normalize_js_text_for_python(bip_name)
    expected_code = store_code_from_name(bip_name)
    active_raw = get_active_store_text(page)
    active = normalize_js_text_for_python(active_raw)

    if not active:
        return False

    # Normal cases.
    if expected_code and expected_code in active:
        return True

    if expected_full and expected_full in active:
        return True

    # Some BIP360 labels are shortened, e.g. "CURITIBA 07 (PR)".
    active_code = store_code_from_name(active_raw)
    if active_code and expected_code and active_code == expected_code:
        return True

    # Last fallback: compare first two tokens + number.
    exp_tokens = expected_code.split()
    act_tokens = active_code.split()
    if len(exp_tokens) >= 2 and len(act_tokens) >= 2 and exp_tokens[:2] == act_tokens[:2]:
        return True

    return False


# ── CLASSIFICAÇÃO PRODUTOS ────────────────────────────────────────────────────
import unicodedata
import re

def normalize(text):
    text = text.lower()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    return text

CANECAKE_SEASONS = ["inverno", "outono", "primavera", "verao", "verão"]
CANECAKE_FLAVORS = ["floresta negra", "merengue limao", "merengue morango", "smores", "s mores", "torta de morango"]

def classify(produto: str) -> str | None:
    p = str(produto or "").strip()
    n = normalize(p)

    # Água
    if "agua" in n:
        return "agua"

    # Chantilly — somente adicionais (Ad Chantilly)
    if "chantilly" in n and re.search(r"\bad\b", n):
        return "chantilly"

    # Shake
    if ("shake mix" in n or "ovomaltine" in n) and ("200" in n or "300" in n):
        return "shake"

    # Milk
    if ("milk shake" in n or "cafe shake" in n) and "500" in n:
        return "milk"

    # Canecake 2026
    # Regras:
    # 1. Deve conter "canecake 2026"
    # 2. Deve conter uma estação do ano
    # 3. Se contiver "ad " no início ou " ad " = é adicional → EXCLUIR
    if "canecake 2026" in n:
        # Excluir adicionais: "ad " no nome
        if re.search(r"\bad\b", n):
            return None
        # Deve ter estação do ano
        has_season = any(s in n for s in CANECAKE_SEASONS)
        if not has_season:
            return None
        return "canecake"

    return None


# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
def get_month_start_col(month: int) -> int:
    # 1-indexed: Jan=C=3, Fev=G=7, Mar=K=11...
    # 5 colunas por mês: Shake, Chantilly, Água, MS500, Canecake
    return 3 + (month - 1) * 5


def get_tm_month_meta_col(month: int) -> int:
    """1-indexed column of Meta TM in 'Metas TM suporte'.
    Jan=B=2, Fev=D=4, Mar=F=6, Abr=H=8...
    Realizado TM is meta_col + 1.
    """
    return 2 + (month - 1) * 2


def get_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(GOOGLE_CREDS, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID).worksheet(SHEET_TAB)


def get_sheet_tm():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(GOOGLE_CREDS, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID).worksheet(SHEET_TAB_TM)


def get_sheet_operador():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(GOOGLE_CREDS, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID).worksheet(SHEET_TAB_OPERADOR)

def get_sheet_hist():
    """Connect to Historico Vendas sheet."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(GOOGLE_CREDS, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID).worksheet(SHEET_TAB_HIST)


def get_existing_hist_dates(ws_hist) -> set:
    """Return set of (store_key, date_str) already in Historico Vendas.
    date_str format: DD/MM/YYYY
    """
    existing = ws_hist.get_all_values()
    result = set()
    for row in existing[1:]:  # skip header
        if len(row) >= 2 and row[0] and row[1]:
            result.add((row[0].strip(), row[1].strip()))
    return result


def update_historico_day(ws_hist, store_key: str, date_str: str, totals: dict, existing_dates: set):
    """Append one day's totals for one store to Historico Vendas.
    Skips if already exists. date_str = DD/MM/YYYY
    """
    if (store_key, date_str) in existing_dates:
        print(f"  ↷ Histórico {store_key} {date_str} já existe — pulando")
        return False

    new_row = [
        store_key,
        date_str,
        int(totals.get("shake", 0) or 0),
        int(totals.get("chantilly", 0) or 0),
        int(totals.get("agua", 0) or 0),
        int(totals.get("milk", 0) or 0),
        int(totals.get("canecake", 0) or 0),
    ]
    ws_hist.append_row(new_row, value_input_option="RAW")
    existing_dates.add((store_key, date_str))
    print(f"  → Histórico {store_key} {date_str}: shake={new_row[2]} chant={new_row[3]} agua={new_row[4]} milk={new_row[5]} canecake={new_row[6]}")
    return True


def prune_old_hist(ws_hist):
    """Keep only last 7 days in Historico Vendas. Called once at end of run."""
    from datetime import timedelta
    cutoff = datetime.now() - timedelta(days=7)

    existing = ws_hist.get_all_values()
    if not existing:
        return

    header = existing[0]
    data_rows = existing[1:]

    def parse_date(row):
        try:
            d, m, y = row[1].split("/")
            return datetime(int(y), int(m), int(d))
        except Exception:
            return datetime.now()

    kept = [r for r in data_rows if parse_date(r) >= cutoff]

    if len(kept) == len(data_rows):
        print(f"  Histórico: {len(kept)} linhas, nenhuma removida")
        return

    final_values = [header] + kept
    ws_hist.clear()
    ws_hist.update("A1", final_values, value_input_option="RAW")
    removed = len(data_rows) - len(kept)
    print(f"  Histórico: {removed} linhas antigas removidas, {len(kept)} mantidas")




def update_realizado(ws_gsheet, store_key: str, totals: dict):
    report_month, report_year = get_report_month_info()
    month_col = get_month_start_col(report_month)
    store_row = STORE_SHEET_ROW[store_key]
    real_row = store_row + 2

    for cat, offset in CAT_OFFSET.items():
        col = month_col + offset
        value = int(totals.get(cat, 0) or 0)
        ws_gsheet.update_cell(real_row, col, value)
        print(f"  → {store_key} {cat}: {value} written to row={real_row} col={col}")


def update_ticket_medio(ws_tm, store_key: str, ticket_medio: float):
    """Write Ticket Médio Realizado in the 'Metas TM suporte' tab."""
    report_month, report_year = get_report_month_info()
    meta_col = get_tm_month_meta_col(report_month)
    realizado_col = meta_col + 1
    row = TM_STORE_SHEET_ROW[store_key]

    # Google Sheets aceita número com ponto como decimal quando enviado pela API.
    value = round(float(ticket_medio or 0), 2)
    ws_tm.update_cell(row, realizado_col, value)
    print(f"  → {store_key} Ticket Médio Realizado: {value:.2f} written to row={row} col={realizado_col}")


# ── BIP360 ────────────────────────────────────────────────────────────────────
def login(page):
    print("Logging in to BIP360...")
    page.goto(BIP_LOGIN_URL, wait_until="load", timeout=45000)
    page.wait_for_load_state("networkidle", timeout=20000)

    page.wait_for_selector('input[type="text"], input[type="email"]', timeout=20000)
    inputs = page.query_selector_all('input[type="text"], input[type="email"]')
    if not inputs:
        raise Exception("Login input not found.")
    inputs[0].fill(BIP_USER)

    page.fill('input[type="password"]', BIP_PASS)

    try:
        page.locator("text=ENTRAR").first.click(timeout=15000)
    except Exception:
        page.click('button:has-text("ENTRAR"), input[value*="ENTRAR"], button:has-text("Entrar")')

    # JSF/PrimeFaces can navigate more than once after login.
    deadline = time.time() + 45
    last_url = ""
    while time.time() < deadline:
        try:
            last_url = page.url
            body = page.locator("body").inner_text(timeout=3000)
            if "publico/login.jsf" not in last_url and "Página Inicial" in body and "Financeiro" in body:
                break
        except Exception:
            pass
        time.sleep(1)

    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    if "publico/login.jsf" in page.url:
        save_debug(page, "login_failed_still_on_login")
        raise Exception("Login failed: still on login page.")

    print("Logged in successfully.")
    print_page_snapshot(page, "after_login")


def open_store_dropdown(page) -> str:
    """Open the store dropdown in the top bar.

    Uses the real HTML structure:
    div.widgets-item.active contains span.name with current franchise and div.lista-empresas.
    """
    js = """
    () => {
        const visible = (el) => {
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style &&
                   style.visibility !== 'hidden' &&
                   style.display !== 'none' &&
                   rect.width > 0 &&
                   rect.height > 0;
        };

        // Most reliable target: current franchise widget in the top bar.
        const widgets = Array.from(document.querySelectorAll('#topBar .widgets-item.active, .widgets-item.active'));
        for (const w of widgets) {
            const list = w.querySelector('.lista-empresas');
            const name = w.querySelector('span.name');
            if (list && name && visible(name)) {
                name.scrollIntoView({block: 'center', inline: 'center'});
                name.click();
                return name.innerText.trim();
            }
        }

        // Fallback: any visible span.name that looks like a store.
        const names = Array.from(document.querySelectorAll('span.name'));
        for (const el of names) {
            const txt = (el.innerText || '').trim();
            if (!visible(el) || !txt) continue;
            if (
                txt.includes('Shopping') ||
                txt.includes('SHOPPING') ||
                txt.includes('São José') ||
                txt.includes('SÃO JOSÉ') ||
                txt.includes('Curitiba') ||
                txt.includes('CURITIBA') ||
                txt.includes('Maring') ||
                txt.includes('MARING') ||
                txt.includes('Campo Largo') ||
                txt.includes('CAMPO LARGO')
            ) {
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
                return txt;
            }
        }

        return 'not found';
    }
    """
    return page.evaluate(js)


def click_store_item(page, bip_name: str) -> str | None:
    """Select store using the real JSF commandLink from .lista-empresas.

    The HTML shows each store as:
    .lista-empresas li > a.ui-commandlink[id=...] > label

    Instead of clicking a visual container, this calls the same PrimeFaces submit
    that the anchor's onclick uses. This avoids the 5/5 partial selection issue.
    """
    print(f"  Active before selection: {get_active_store_text(page)}")

    expected_code = store_code_from_name(bip_name)

    js = """
    ([targetName, targetCode]) => {
        const normalize = (s) => String(s || '')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim()
            .toUpperCase();

        const wanted = normalize(targetName);
        const code = normalize(targetCode);

        const lists = Array.from(document.querySelectorAll('.lista-empresas'));
        let anchors = [];

        for (const list of lists) {
            anchors = anchors.concat(Array.from(list.querySelectorAll('li a.ui-commandlink, li a')));
        }

        if (!anchors.length) {
            anchors = Array.from(document.querySelectorAll('a.ui-commandlink, a'));
        }

        const matches = [];

        for (const a of anchors) {
            const label = a.querySelector('label');
            const rawText = (label ? label.innerText : a.innerText || a.textContent || '').replace(/\s+/g, ' ').trim();
            const text = normalize(rawText);

            if (!text) continue;

            if (text === wanted || text.includes(wanted) || text.includes(code)) {
                matches.push({
                    id: a.id || '',
                    rawText,
                    text,
                    onclick: a.getAttribute('onclick') || '',
                    len: rawText.length
                });
            }
        }

        matches.sort((a, b) => a.len - b.len);

        if (!matches.length) {
            return { ok: false, reason: 'not_found', targetName, targetCode };
        }

        const chosen = matches[0];

        if (!chosen.id) {
            return { ok: false, reason: 'no_id', chosen };
        }

        // Most reliable path: reproduce PrimeFaces commandLink submission.
        if (window.PrimeFaces && PrimeFaces.addSubmitParam) {
            const payload = {};
            payload[chosen.id] = chosen.id;
            PrimeFaces.addSubmitParam('topBar', payload).submit('topBar');
            return { ok: true, method: 'PrimeFaces.addSubmitParam', chosen };
        }

        // Fallback: execute the onclick code exactly as defined.
        const el = document.getElementById(chosen.id);
        if (el) {
            el.click();
            return { ok: true, method: 'native_click', chosen };
        }

        return { ok: false, reason: 'element_not_found_after_match', chosen };
    }
    """

    try:
        result = page.evaluate(js, [bip_name, expected_code])
        print(f"  Store submit JS result: {result}")

        if not result or not result.get("ok"):
            return None

        # JSF submit causes a full page reload to dashboard.
        try:
            page.wait_for_load_state("load", timeout=25000)
        except Exception:
            pass

        wait_bip_idle(page, timeout=45)

        if wait_active_store(page, bip_name, timeout=45):
            chosen = result.get("chosen", {})
            return f"jsf_submit: {chosen.get('rawText')}"

        print(f"  Store submit did not confirm change. active={get_active_store_text(page)} expected={bip_name}")
        return None

    except Exception as e:
        print(f"  Store click by JSF submit failed: {e}")
        return None


def select_store(page, bip_name: str):
    page.wait_for_load_state("domcontentloaded", timeout=15000)
    wait_bip_idle(page, timeout=45)
    time.sleep(1)

    before = get_active_store_text(page)
    print(f"  Active store before: {before}")

    if active_store_matches(page, bip_name):
        print(f"  Store already active: {bip_name}")
        return

    last_error = None

    for attempt in range(1, 4):
        print(f"  Opening store dropdown... attempt {attempt}/3")
        result = open_store_dropdown(page)
        print(f"  Dropdown click result: {result}")
        time.sleep(1.5)

        if result == "not found":
            last_error = "Could not open store dropdown."
            save_debug(page, f"store_dropdown_not_found_attempt_{attempt}")
            continue

        if attempt == 1:
            print_page_snapshot(page, "after_dropdown_click")

        clicked = click_store_item(page, bip_name)

        if clicked:
            print(f"  Store click result: {clicked}")

            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass

            time.sleep(2)
            after = get_active_store_text(page)
            print(f"  Active store after: {after}")

            if active_store_matches(page, bip_name):
                print_page_snapshot(page, f"after_select_{slug(bip_name)}")
                print(f"  Store selected and confirmed: {bip_name}")
                return

            last_error = f"Store selection not confirmed. Expected '{bip_name}', active top bar is '{after}'"
        else:
            last_error = f"Could not click store '{bip_name}'"

        time.sleep(2)

    save_debug(page, f"could_not_select_store_{slug(bip_name)}")
    print_page_snapshot(page, f"could_not_select_store_{slug(bip_name)}")
    raise Exception(last_error or f"Could not select store '{bip_name}'")


def go_to_report(page):
    """Vai para o relatório correto.

    Caminho confirmado manualmente:
    Financeiro → Relatórios → Venda de Produto por Operador

    No BIP360/JSF, a URL direta só funciona se for no mesmo domínio da sessão.
    Por isso tentamos primeiro a URL exata sem www; se não abrir, usamos menu.
    """
    print("  Opening report page...")

    # Tentativa 1: URL direta no mesmo domínio da sessão, conforme tela manual.
    try:
        page.goto(REPORT_URL, wait_until="load", timeout=30000)
        time.sleep(4)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass

        body = page.locator("body").inner_text(timeout=5000)
        if "ENTRAR" not in body and ("Relatório de Ranking dos Operadores" in body or "Data Inicial" in body):
            print("  Report page opened by direct URL ✔")
            print_page_snapshot(page, "after_report_direct_url")
            return

        print("  Direct URL did not open report, trying menu...")
    except Exception as e:
        print(f"  Direct URL failed, trying menu: {e}")

    # Tentativa 2: menu real.
    try:
        page.goto("https://itfgestor.com.br/ITFGestor/dashboardPDV2.jsf", wait_until="load", timeout=30000)
        time.sleep(3)

        page.locator("text=Financeiro").first.click(timeout=10000)
        time.sleep(1.5)

        page.locator("text=Relatórios").first.click(timeout=10000)
        time.sleep(1.5)

        page.locator("text=Venda de Produto por Operador").first.click(timeout=10000)
        time.sleep(4)

        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass

        print("  Report page opened by menu ✔")
        print_page_snapshot(page, "after_report_menu")
        return

    except Exception as e:
        save_debug(page, "go_to_report_failed")
        print_page_snapshot(page, "go_to_report_failed")
        raise Exception(f"Erro ao navegar até relatório: {e}")


def set_report_dates(page, start_date: str, end_date: str):
    """Set Data Inicial and Data Final using only visible text inputs.

    PrimeFaces creates hidden/helper inputs for calendars. The old version tried to click
    the first input[type=text], which could be hidden. This version filters visible inputs
    and sets the first two visible date fields.
    """
    print("  Filling visible date fields...")

    result = page.evaluate(
        """([startDate, endDate]) => {
            const isVisible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style &&
                       style.visibility !== 'hidden' &&
                       style.display !== 'none' &&
                       rect.width > 0 &&
                       rect.height > 0 &&
                       !el.disabled &&
                       !el.readOnly;
            };

            const inputs = Array.from(document.querySelectorAll('input[type="text"]'))
                .filter(isVisible);

            const dateLike = inputs.filter(input => {
                const value = input.value || '';
                const placeholder = input.getAttribute('placeholder') || '';
                const aria = input.getAttribute('aria-label') || '';
                const id = input.id || '';
                const name = input.name || '';
                const blob = `${value} ${placeholder} ${aria} ${id} ${name}`.toLowerCase();

                return blob.includes('data') ||
                       blob.includes('date') ||
                       /\d{2}\/\d{2}\/\d{4}/.test(value) ||
                       inputs.indexOf(input) <= 2;
            });

            const targets = dateLike.length >= 2 ? dateLike.slice(0, 2) : inputs.slice(0, 2);

            if (targets.length < 2) {
                return {
                    ok: false,
                    totalVisible: inputs.length,
                    totalDateLike: dateLike.length,
                    values: inputs.map(i => i.value || '')
                };
            }

            const setValue = (el, value) => {
                el.focus();
                el.value = value;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.blur();
            };

            setValue(targets[0], startDate);
            setValue(targets[1], endDate);

            return {
                ok: true,
                totalVisible: inputs.length,
                totalDateLike: dateLike.length,
                selected: targets.map(i => ({
                    id: i.id || '',
                    name: i.name || '',
                    value: i.value || '',
                    placeholder: i.getAttribute('placeholder') || ''
                }))
            };
        }""",
        [start_date, end_date]
    )

    print(f"  Date fill result: {result}")

    if not result or not result.get("ok"):
        raise Exception(f"Could not fill date fields. Result: {result}")


def click_pesquisar(page):
    print("  Clicking Pesquisar...")

    # Botão azul "Pesquisar" da tela.
    try:
        page.locator("text=Pesquisar").first.click(timeout=15000)
        return
    except Exception:
        pass

    clicked = page.evaluate(
        """() => {
            const els = Array.from(document.querySelectorAll('button, a, span, div'));
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' &&
                       style.display !== 'none' &&
                       rect.width > 0 && rect.height > 0;
            };
            const found = els.find(el => visible(el) && (el.innerText || '').trim().includes('Pesquisar'));
            if (found) {
                found.click();
                return found.innerText.trim();
            }
            return null;
        }"""
    )
    if not clicked:
        raise Exception("Pesquisar button not found.")


def click_xls_export(page):
    print("  Looking for XLS export button...")

    wait_bip_idle(page, timeout=45)

    # Make sure the bottom-right XLS icon is in the viewport.
    try:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1)
    except Exception:
        pass

    # Main reliable path: click the parent <a> of the xls.png image.
    selectors = [
        'a:has(img[src*="xls"])',
        'a:has(img[src*="excel"])',
        'img[src*="xls"]',
        'img[src*="excel"]',
    ]

    last_error = None

    for selector in selectors:
        try:
            loc = page.locator(selector).last
            if loc.count() == 0:
                continue

            with page.expect_download(timeout=90000) as dl_info:
                loc.click(timeout=15000, force=True)

            download = dl_info.value
            print(f"  XLS downloaded using selector: {selector}")
            return download

        except Exception as e:
            last_error = e
            print(f"  XLS selector failed ({selector}): {e}")

    # JS fallback based on real HTML.
    try:
        with page.expect_download(timeout=90000) as dl_info:
            clicked = page.evaluate(
                """() => {
                    const imgs = Array.from(document.querySelectorAll('img[src*="xls" i], img[src*="excel" i]'));

                    for (const img of imgs) {
                        const a = img.closest('a');
                        if (!a) continue;

                        a.scrollIntoView({block: 'center', inline: 'center'});
                        img.scrollIntoView({block: 'center', inline: 'center'});

                        if (typeof a.click === 'function') {
                            a.click();
                            return {
                                ok: true,
                                method: 'a.click',
                                href: a.getAttribute('href') || '',
                                onclick: a.getAttribute('onclick') || '',
                                img: img.getAttribute('src') || ''
                            };
                        }
                    }

                    return {ok: false, reason: 'xls_img_anchor_not_found'};
                }"""
            )

            print(f"  XLS JS fallback result: {clicked}")

            if not clicked or not clicked.get("ok"):
                raise Exception(f"XLS export button not found. Result: {clicked}")

        return dl_info.value

    except Exception as e:
        last_error = e

    raise Exception(f"XLS export button/download failed. Last error: {last_error}")


def download_xls(page, store: dict, download_dir: str) -> str | None:
    start_date, end_date = get_report_date_strings(include_seconds=True)

    clean_xls_files(download_dir)

    go_to_report(page)

    active_before_report = get_active_store_text(page)
    print(f"  Active store on report page: {active_before_report}")

    if not active_store_matches(page, store["bip_name"]):
        save_debug(page, f"wrong_store_before_report_{store['key']}")
        raise Exception(
            f"Wrong active store before report. Expected={store['bip_name']} Active={active_before_report}"
        )

    print(f"  Setting dates: {start_date} → {end_date}")
    try:
        set_report_dates(page, start_date, end_date)
    except Exception:
        save_debug(page, f"date_fields_not_found_{store['key']}")
        print_page_snapshot(page, f"date_fields_not_found_{store['key']}")
        raise

    time.sleep(0.8)
    click_pesquisar(page)

    # PrimeFaces updates the result table by AJAX. Wait until loader is gone.
    wait_bip_idle(page, timeout=60)

    # Wait for a real result area and XLS button after the search.
    deadline = time.time() + 75
    table_ready = False
    last_state = None

    while time.time() < deadline:
        try:
            state = page.evaluate(
                """() => {
                    const body = document.body.innerText || '';
                    const rows = Array.from(document.querySelectorAll('table tbody tr'));
                    const nonEmptyRows = rows.filter(r => (r.innerText || '').trim().length > 0).length;
                    const hasPagination = body.includes('Página:');
                    const hasNoRecords = body.toLowerCase().includes('nenhum registro');
                    const hasXls = !!document.querySelector('img[src*="xls" i], img[src*="excel" i]');
                    const hasReportTitle = body.includes('Relatório de Ranking dos Operadores');

                    return {nonEmptyRows, hasPagination, hasNoRecords, hasXls, hasReportTitle};
                }"""
            )
            last_state = state

            if state.get("hasReportTitle") and state.get("hasXls") and (state.get("nonEmptyRows", 0) > 0 or state.get("hasNoRecords")):
                table_ready = True
                break

        except Exception:
            pass

        time.sleep(1)

    print(f"  Table ready: {table_ready} | State: {last_state}")

    if not table_ready:
        save_debug(page, f"table_not_ready_before_xls_{store['key']}")
        print_page_snapshot(page, f"table_not_ready_before_xls_{store['key']}")
        raise Exception("Result table/XLS not confirmed before export.")

    wait_bip_idle(page, timeout=30)

    print_page_snapshot(page, f"after_search_{store['key']}")

    clean_xls_files(download_dir)

    try:
        download = click_xls_export(page)
    except Exception:
        save_debug(page, f"xls_button_not_found_{store['key']}")
        print_page_snapshot(page, f"xls_button_not_found_{store['key']}")
        raise

    file_path = os.path.join(download_dir, f"{store['key']}.xls")
    download.save_as(file_path)

    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        raise Exception(f"Downloaded XLS is missing or empty: {file_path}")

    print(f"  Downloaded: {file_path} ({os.path.getsize(file_path)} bytes)")
    copy_xls_debug(file_path, store["key"], "operadores")
    return file_path




def download_xls_daterange(page, store: dict, download_dir: str, day_start: str, day_end: str, suffix: str = "") -> str | None:
    """Download XLS for a specific date range (used for daily historico)."""
    go_to_report(page)

    print(f"  Setting dates: {day_start} → {day_end}")
    try:
        set_report_dates(page, day_start, day_end)
    except Exception:
        save_debug(page, f"hist_date_fields_{store['key']}{suffix}")
        raise

    time.sleep(0.5)
    click_pesquisar(page)

    wait_bip_idle(page, timeout=45)

    # Check if there are results
    state = page.evaluate("""() => {
        const body = document.body.innerText || '';
        const hasXls = !!document.querySelector('img[src*="xls" i]');
        const hasNoRecords = body.toLowerCase().includes('nenhum registro');
        return {hasXls, hasNoRecords};
    }""")

    if state.get("hasNoRecords") and not state.get("hasXls"):
        print(f"  ↷ Sem vendas neste dia")
        return None

    try:
        download = click_xls_export(page)
    except Exception:
        print(f"  ⚠️ XLS não encontrado para {day_start}")
        return None

    file_path = os.path.join(download_dir, f"{store['key']}{suffix}.xls")
    download.save_as(file_path)
    return file_path


def go_to_tm_report(page):
    """Open Ticket Médio Diário report."""
    print("  Opening Ticket Médio Diário report...")

    try:
        page.goto(TM_REPORT_URL, wait_until="load", timeout=30000)
        time.sleep(4)
        wait_bip_idle(page, timeout=45)

        body = page.locator("body").inner_text(timeout=5000)
        if "Relatório de Ticket Médio Diário" in body or "Ticket Médio Líquido" in body:
            print("  TM report page opened by direct URL ✔")
            print_page_snapshot(page, "after_tm_report_direct_url")
            return

        print("  TM direct URL did not open report, trying menu...")
    except Exception as e:
        print(f"  TM direct URL failed, trying menu: {e}")

    try:
        page.goto("https://itfgestor.com.br/ITFGestor/dashboardPDV2.jsf", wait_until="load", timeout=30000)
        wait_bip_idle(page, timeout=45)
        time.sleep(2)

        page.locator("text=Financeiro").first.click(timeout=10000)
        time.sleep(1.2)
        page.locator("text=Relatórios").first.click(timeout=10000)
        time.sleep(1.2)
        page.locator("text=Ticket Médio Diário").first.click(timeout=10000)
        time.sleep(4)

        wait_bip_idle(page, timeout=45)
        print("  TM report page opened by menu ✔")
        print_page_snapshot(page, "after_tm_report_menu")
        return

    except Exception as e:
        save_debug(page, "go_to_tm_report_failed")
        print_page_snapshot(page, "go_to_tm_report_failed")
        raise Exception(f"Erro ao navegar até relatório de Ticket Médio Diário: {e}")


def set_tm_report_dates(page, start_date: str, end_date: str):
    """Set TM report dates. This report usually uses dd/mm/yyyy HH:MM."""
    print("  Filling TM date fields...")

    result = page.evaluate(
        """([startDate, endDate]) => {
            const isVisible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style &&
                       style.visibility !== 'hidden' &&
                       style.display !== 'none' &&
                       rect.width > 0 &&
                       rect.height > 0 &&
                       !el.disabled &&
                       !el.readOnly;
            };

            const inputs = Array.from(document.querySelectorAll('input[type="text"]')).filter(isVisible);
            const dateLike = inputs.filter(input => {
                const value = input.value || '';
                const id = input.id || '';
                const name = input.name || '';
                const blob = `${value} ${id} ${name}`.toLowerCase();
                return blob.includes('data') || blob.includes('date') || /\\d{2}\\/\\d{2}\\/\\d{4}/.test(value);
            });

            const targets = dateLike.length >= 2 ? dateLike.slice(0, 2) : inputs.slice(0, 2);

            if (targets.length < 2) {
                return { ok:false, visible: inputs.length, dateLike: dateLike.length };
            }

            const setValue = (el, value) => {
                el.focus();
                el.value = value;
                el.dispatchEvent(new Event('input', { bubbles:true }));
                el.dispatchEvent(new Event('change', { bubbles:true }));
                el.blur();
            };

            setValue(targets[0], startDate);
            setValue(targets[1], endDate);

            return {
                ok:true,
                selected: targets.map(i => ({id:i.id || '', name:i.name || '', value:i.value || ''}))
            };
        }""",
        [start_date, end_date]
    )

    print(f"  TM date fill result: {result}")

    if not result or not result.get("ok"):
        raise Exception(f"Could not fill TM date fields. Result: {result}")


def wait_tm_table_ready(page):
    """Wait until the TM table/footer is loaded."""
    wait_bip_idle(page, timeout=60)

    deadline = time.time() + 60
    while time.time() < deadline:
        try:
            state = page.evaluate(
                """() => {
                    const body = document.body.innerText || '';
                    const hasTotals = body.includes('Totais');
                    const hasTicket = body.includes('Ticket Médio Líquido');
                    const hasXls = !!document.querySelector('img[src*="xls" i], img[src*="excel" i]');
                    return {hasTotals, hasTicket, hasXls};
                }"""
            )
            if state.get("hasTotals") and state.get("hasTicket"):
                return True
            if state.get("hasXls") and state.get("hasTicket"):
                return True
        except Exception:
            pass
        time.sleep(1)

    return False


def download_tm_xls(page, store: dict, download_dir: str) -> str:
    """Download Ticket Médio Diário XLS for the selected store."""
    # Ticket Médio Diário usa minuto, sem segundos, conforme tela do BIP360.
    start_date, end_date = get_report_date_strings(include_seconds=False)

    go_to_tm_report(page)

    print(f"  Setting TM dates: {start_date} → {end_date}")
    try:
        set_tm_report_dates(page, start_date, end_date)
    except Exception:
        save_debug(page, f"tm_date_fields_not_found_{store['key']}")
        print_page_snapshot(page, f"tm_date_fields_not_found_{store['key']}")
        raise

    time.sleep(0.8)
    click_pesquisar(page)

    table_ready = wait_tm_table_ready(page)
    if not table_ready:
        print("  ⚠️ TM table not fully confirmed; trying XLS anyway.")

    print_page_snapshot(page, f"after_tm_search_{store['key']}")

    try:
        download = click_xls_export(page)
    except Exception:
        save_debug(page, f"tm_xls_button_not_found_{store['key']}")
        print_page_snapshot(page, f"tm_xls_button_not_found_{store['key']}")
        raise

    file_path = os.path.join(download_dir, f"{store['key']}_tm.xls")
    download.save_as(file_path)
    print(f"  TM downloaded: {file_path}")
    return file_path


def parse_ticket_medio_xls(file_path: str) -> float:
    """Parse Ticket Médio Líquido from the Totais row of Ticket Médio Diário XLS.

    Expected columns in report:
    Data | Pessoas Atendidas | Ticket Médio Líquido (R$) | Venda em Itens | Ticket Médio por Produto | Receitas...
    We locate the column by header and then read the row containing 'Totais'.
    """
    wb = xlrd.open_workbook(file_path)
    ws = wb.sheets()[0]

    header_row = None
    col_tm = None

    for r in range(min(12, ws.nrows)):
        vals = [str(ws.cell_value(r, c)).strip().lower() for c in range(ws.ncols)]
        for c, val in enumerate(vals):
            if "ticket" in val and "médio" in val and "líquido" in val:
                col_tm = c
                header_row = r
                break
            # fallback without accents
            if "ticket" in val and "medio" in val and "liquido" in val:
                col_tm = c
                header_row = r
                break
        if col_tm is not None:
            break

    # Fallback from screen: Ticket Médio Líquido is the third data column, index 2.
    if col_tm is None:
        col_tm = 2

    total_row = None
    for r in range(ws.nrows):
        row_text = " ".join(str(ws.cell_value(r, c)).strip().lower() for c in range(ws.ncols))
        if "totais" in row_text or "total" in row_text:
            total_row = r

    if total_row is None:
        # fallback: last non-empty row
        for r in range(ws.nrows - 1, -1, -1):
            if any(str(ws.cell_value(r, c)).strip() for c in range(ws.ncols)):
                total_row = r
                break

    if total_row is None:
        raise Exception("Could not locate Totais row in TM XLS.")

    raw = ws.cell_value(total_row, col_tm)

    if isinstance(raw, (int, float)):
        return float(raw)

    s = str(raw or "").strip()
    s = re.sub(r"[^0-9,.-]", "", s)
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return float(s or 0)



def normalize_operator_name(name: str) -> str:
    name = str(name or "").strip()
    name = re.sub(r"\s+", " ", name)
    return name or "Operador"


def parse_xls_by_operator(file_path: str, store_key: str = "") -> tuple[dict, dict]:
    """Parse Venda de Produto por Operador XLS.

    Returns:
      totals: {shake, chantilly, agua, milk}
      operators: {
        "OPERADOR": {shake, chantilly, agua, milk, total}
      }

    Blindagem:
      - loga linha a linha para DEBUG_STORES;
      - salva CSV de diagnóstico;
      - evita linhas idênticas duplicadas dentro do mesmo XLS.
    """
    totals = {"shake": 0, "chantilly": 0, "agua": 0, "milk": 0, "canecake": 0}
    operators = {}
    debug_rows = []
    seen_rows = set()
    skipped_duplicates = 0
    parse_errors = 0

    wb = xlrd.open_workbook(file_path)
    ws = wb.sheets()[0]

    print(f"  XLS info: sheet={ws.name!r}, rows={ws.nrows}, cols={ws.ncols}")

    header_row = None
    col_operador = None
    col_produto = None
    col_quantidade = None

    # Locate headers flexibly.
    for r in range(min(20, ws.nrows)):
        row_vals = [str(ws.cell_value(r, c)).strip().lower() for c in range(ws.ncols)]

        for c, val in enumerate(row_vals):
            val_norm = (
                val.replace("á", "a").replace("à", "a").replace("ã", "a")
                   .replace("é", "e").replace("ê", "e")
                   .replace("í", "i")
                   .replace("ó", "o").replace("ô", "o").replace("õ", "o")
                   .replace("ú", "u")
                   .replace("ç", "c")
            )

            if val_norm == "operador" or "operador" in val_norm:
                col_operador = c
            if val_norm == "produto" or "produto" in val_norm:
                col_produto = c

            # Relatório confirmado: Rank | Operador | Produto | Quantidade.
            # Usa a coluna cujo cabeçalho é exatamente "Quantidade" ou "Qtd".
            if val_norm in ("quantidade", "qtd"):
                col_quantidade = c

        if col_operador is not None and col_produto is not None and col_quantidade is not None:
            header_row = r
            break

    # Fallback based on screen/report structure:
    # Rank | Operador | Produto | Quantidade
    if col_operador is None:
        col_operador = 1
    if col_produto is None:
        col_produto = 2
    if col_quantidade is None:
        col_quantidade = 3

    print(
        f"  XLS columns: header_row={header_row}, "
        f"operador={col_operador}, produto={col_produto}, quantidade={col_quantidade}"
    )

    start_row = (header_row + 1) if header_row is not None else 1

    for r in range(start_row, ws.nrows):
        try:
            operador = normalize_operator_name(ws.cell_value(r, col_operador))
            produto = str(ws.cell_value(r, col_produto)).strip()

            if not produto or produto.lower() in ("produto", "total", "totais"):
                continue

            raw_qty = ws.cell_value(r, col_quantidade)
            qty = parse_qty_cell(raw_qty)

            if qty <= 0:
                continue

            cat = classify(produto)

            used = bool(cat)
            debug_rows.append({
                "row": r + 1,
                "operador": operador,
                "produto": produto,
                "quantidade": qty,
                "categoria": cat or "",
                "used": used,
            })

            if not cat:
                continue

            # Evita contar a mesma linha do mesmo XLS duas vezes se o arquivo vier com duplicidade interna.
            # A chave inclui linha lógica completa.
            row_key = (
                operador.strip().lower(),
                produto.strip().lower(),
                round(float(qty), 4),
                cat,
            )

            if row_key in seen_rows:
                skipped_duplicates += 1
                print(f"  ⚠️ Duplicate XLS row skipped: {operador} | {produto} | {qty} | {cat}")
                continue

            seen_rows.add(row_key)

            if operador not in operators:
                operators[operador] = {"shake": 0, "chantilly": 0, "agua": 0, "milk": 0, "canecake": 0, "total": 0}

            operators[operador][cat] += qty
            operators[operador]["total"] += qty
            totals[cat] += qty

        except Exception as e:
            parse_errors += 1
            print(f"  ⚠️ Parse row error at Excel row {r + 1}: {e}")
            continue

    totals = {k: int(round(v)) for k, v in totals.items()}

    for op in operators:
        for cat in ["shake", "chantilly", "agua", "milk", "total"]:
            operators[op][cat] = int(round(operators[op].get(cat, 0) or 0))

    if skipped_duplicates:
        print(f"  ⚠️ Duplicate XLS rows skipped: {skipped_duplicates}")

    if parse_errors:
        print(f"  ⚠️ Parse row errors: {parse_errors}")

    write_operator_parse_debug(store_key, debug_rows)

    if is_debug_store(store_key):
        print_operator_category_summary(store_key, operators)

    return totals, operators



def clean_int_for_sheet(value):
    """Return clean integer for Google Sheets, avoiding apostrophe/text values."""
    try:
        if value is None:
            return 0
        if isinstance(value, str):
            value = value.strip().replace("'", "").replace(".", "").replace(",", ".")
        return int(round(float(value or 0)))
    except Exception:
        return 0



def canon_month_key(value: str) -> str:
    """Normalize month keys so 'abr-2026' and 'abr.-2026' are treated as the same."""
    s = str(value or "").strip().lower()
    s = (
        s.replace(".", "")
         .replace(" ", "")
         .replace("ç", "c")
         .replace("á", "a").replace("à", "a").replace("ã", "a")
         .replace("é", "e").replace("ê", "e")
         .replace("í", "i")
         .replace("ó", "o").replace("ô", "o").replace("õ", "o")
         .replace("ú", "u")
    )
    return s


def canon_store_key(value: str) -> str:
    """Normalize store keys for comparisons."""
    return str(value or "").strip().lower().replace(" ", "")


def update_operadores(ws_op, store_key: str, operators: dict):
    """Rewrite rows for current month + store in 'Metas Operador suporte'.

    Sheet columns:
      A Mês | B Loja | C Operador | D Água | E Chantilly | F Shake Mix | G Milk Shake | H Total
    """
    report_month, report_year = get_report_month_info()
    month_key = f"{MESES_PT[report_month]}-{report_year}"

    header = ["Mês", "Loja", "Operador", "Água", "Chantilly", "Shake Mix", "Milk Shake", "Total"]

    existing = ws_op.get_all_values()

    # Keep rows that are not the current month/store.
    kept = []
    if existing:
        data_rows = existing[1:] if existing[0] else existing
        for row in data_rows:
            row = row + [""] * (8 - len(row))
            if not row[0] and not row[1] and not row[2]:
                continue
            # Remove any existing rows for the same month + store before rewriting.
            # This also removes rows formatted by Google Sheets as "abr.-2026".
            if canon_month_key(row[0]) == canon_month_key(month_key) and canon_store_key(row[1]) == canon_store_key(store_key):
                continue
            # IMPORTANTE:
            # get_all_values() sempre retorna texto. Se regravarmos essas linhas como vieram,
            # as colunas numéricas D:H voltam como texto/apóstrofo no Google Sheets.
            # Por isso, ao preservar linhas de outras lojas, reconvertemos D:H para int.
            kept.append([
                str(row[0]).strip(),
                str(row[1]).strip(),
                str(row[2]).strip(),
                clean_int_for_sheet(row[3]),
                clean_int_for_sheet(row[4]),
                clean_int_for_sheet(row[5]),
                clean_int_for_sheet(row[6]),
                clean_int_for_sheet(row[7]),
            ])

    new_rows = []
    for operador, vals in sorted(operators.items(), key=lambda kv: kv[1].get("total", 0), reverse=True):
        new_rows.append([
            month_key,
            store_key,
            operador,
            clean_int_for_sheet(vals.get("agua", 0)),
            clean_int_for_sheet(vals.get("chantilly", 0)),
            clean_int_for_sheet(vals.get("shake", 0)),
            clean_int_for_sheet(vals.get("milk", 0)),
            clean_int_for_sheet(vals.get("total", 0)),
        ])

    final_values = [header] + kept + new_rows

    # Auditoria rápida de tipo: colunas D:H precisam chegar como int, não string.
    numeric_type_errors = []
    for idx, row in enumerate(final_values[1:], start=2):
        for col_idx in range(3, 8):
            if not isinstance(row[col_idx], int):
                numeric_type_errors.append((idx, col_idx + 1, row[col_idx], type(row[col_idx]).__name__))
    if numeric_type_errors:
        print(f"  ⚠️ Numeric type warning before Sheets update: {numeric_type_errors[:10]}")

    ws_op.clear()
    ws_op.update("A1", final_values, value_input_option="RAW")

    print(f"  → {store_key} operadores: {len(new_rows)} rows written to '{SHEET_TAB_OPERADOR}' for {month_key}")


def parse_xls(file_path: str) -> dict:
    totals = {"shake": 0, "chantilly": 0, "agua": 0, "milk": 0, "canecake": 0}

    try:
        wb = xlrd.open_workbook(file_path)
        ws = wb.sheets()[0]

        # Relatório esperado:
        # Operador | Produto | Quantidade (às vezes com colunas extras)
        # Vamos localizar colunas por cabeçalho quando possível.
        header_row = None
        col_produto = None
        col_quantidade = None

        for r in range(min(10, ws.nrows)):
            row_vals = [str(ws.cell_value(r, c)).strip().lower() for c in range(ws.ncols)]
            for c, val in enumerate(row_vals):
                if "produto" in val:
                    col_produto = c
                if "quantidade" in val or "qtd" in val:
                    col_quantidade = c
            if col_produto is not None and col_quantidade is not None:
                header_row = r
                break

        # Fallback baseado na versão original: produto col 2, quantidade col 3.
        if col_produto is None:
            col_produto = 2
        if col_quantidade is None:
            col_quantidade = 3
        start_row = (header_row + 1) if header_row is not None else 1

        for r in range(start_row, ws.nrows):
            produto = str(ws.cell_value(r, col_produto)).strip()
            if not produto:
                continue

            raw_qty = ws.cell_value(r, col_quantidade)
            try:
                qty = float(raw_qty or 0)
            except Exception:
                qty = float(str(raw_qty).replace(".", "").replace(",", ".") or 0)

            cat = classify(produto)
            if cat:
                totals[cat] += qty

    except Exception as e:
        print(f"  ⚠️  Error parsing {file_path}: {e}")

    return {k: int(v) for k, v in totals.items()}


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print(f"=== BIP360 Scraper — {br_now().strftime('%d/%m/%Y %H:%M')} BRT ===")

    ws_gsheet = get_sheet()
    ws_tm = get_sheet_tm()
    ws_op = get_sheet_operador()
    ws_hist = get_sheet_hist()
    report_month, report_year = get_report_month_info()
    start_date_dbg, end_date_dbg = get_report_date_strings(include_seconds=True)
    print(f"Google Sheets connected. Report month: {MESES_PT[report_month]}-{report_year}")
    print(f"Report period D-1 (BRT): {start_date_dbg} → {end_date_dbg}")

    # Load existing historico dates once (avoid repeated reads per store)
    existing_hist_dates = get_existing_hist_dates(ws_hist)
    print(f"Histórico: {len(existing_hist_dates)} entradas existentes")

    # Build list of last 7 days (D-1 to D-7) in DD/MM/YYYY
    from datetime import timedelta
    now_brt = datetime.now()
    hist_days = [(now_brt - timedelta(days=i)).strftime("%d/%m/%Y") for i in range(1, 8)]
    print(f"Dias a verificar: {hist_days}")

    success_count = 0
    error_count = 0

    with tempfile.TemporaryDirectory() as tmpdir:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)

            for store in LOJAS:
                print(f"\n--- {store['key'].upper()} : {store['bip_name']} ---")
                context = None
                page = None

                try:
                    context = browser.new_context(
                        accept_downloads=True,
                        viewport={"width": 1600, "height": 1000},
                    )
                    page = context.new_page()
                    page.set_default_timeout(30000)

                    login(page)
                    select_store(page, store["bip_name"])

                    file_path = download_xls(page, store, tmpdir)
                    if not file_path or not os.path.exists(file_path):
                        raise Exception("No XLS file downloaded.")

                    totals, operators = parse_xls_by_operator(file_path, store['key'])
                    print(f"  Totals: {totals}")
                    print(f"  Operators parsed: {len(operators)}")

                    update_realizado(ws_gsheet, store["key"], totals)
                    update_operadores(ws_op, store["key"], operators)

                    # ── Historico Vendas: download dia a dia ──────────────
                    for hist_date in hist_days:
                        if (store["key"], hist_date) in existing_hist_dates:
                            print(f"  ↷ {store['key']} {hist_date} já existe")
                            continue
                        # Parse hist_date to set exact day range
                        hd, hm, hy = hist_date.split("/")
                        day_start = f"{hd}/{hm}/{hy} 00:00:00"
                        day_end   = f"{hd}/{hm}/{hy} 23:59:59"
                        print(f"  📅 Baixando histórico {store['key']} {hist_date}...")
                        try:
                            hist_file = download_xls_daterange(page, store, tmpdir, day_start, day_end, suffix=f"_hist_{hd}{hm}")
                            if hist_file and os.path.exists(hist_file):
                                _, hist_ops = parse_xls_by_operator(hist_file)
                                hist_totals = {cat: sum(op.get(cat,0) for op in hist_ops.values()) for cat in ["shake","chantilly","agua","milk","canecake"]}
                                update_historico_day(ws_hist, store["key"], hist_date, hist_totals, existing_hist_dates)
                        except Exception as he:
                            print(f"  ⚠️ Histórico {hist_date} falhou: {he}")

                    # Ticket Médio Diário → aba Metas TM suporte
                    tm_file_path = download_tm_xls(page, store, tmpdir)
                    tm_value = parse_ticket_medio_xls(tm_file_path)
                    print(f"  Ticket Médio Realizado: {tm_value:.2f}")
                    update_ticket_medio(ws_tm, store["key"], tm_value)

                    success_count += 1

                except PlaywrightTimeout as e:
                    error_count += 1
                    print(f"  ⚠️  Timeout for {store['key']}: {e}")
                    if page:
                        save_debug(page, f"timeout_{store['key']}")
                        print_page_snapshot(page, f"timeout_{store['key']}")

                except Exception as e:
                    error_count += 1
                    print(f"  ⚠️  Error for {store['key']}: {e}")
                    if page:
                        save_debug(page, f"error_{store['key']}")
                        print_page_snapshot(page, f"error_{store['key']}")

                finally:
                    if context:
                        context.close()

            browser.close()

    # Prune historico to keep only last 7 days
    try:
        prune_old_hist(ws_hist)
    except Exception as pe:
        print(f"  ⚠️ Prune historico falhou: {pe}")

    print(f"\n=== Done! Success: {success_count} | Errors: {error_count} ===")

    if success_count == 0:
        raise RuntimeError("Nenhuma loja foi processada com sucesso. Verifique os logs acima.")


if __name__ == "__main__":
    main()
