"""Diagnóstico (só leitura): por que o Índice "Quantidade de Pessoas Atendidas"
volta "Nenhum Registro Encontrado" para o coletor, se na mão (04/10/2026, SJP 1,
03/10 de 00:00:00 a 23:59:59) ele traz Totem 156 e loja 310.

Roda variações do preenchimento da tela, cada uma numa página recém-aberta, e
imprime o que foi enviado no Pesquisar (campos do formulário, sem ViewState) e
o que voltou na tabela. Não grava nada: nem planilha, nem banco.
"""

import re
import time
from urllib.parse import parse_qsl

from playwright.sync_api import sync_playwright

import scraper as s

LOJA = next(l for l in s.LOJAS if l["key"] == "sjp1")
DIA_INI, DIA_FIM = "03/10/2026 00:00:00", "03/10/2026 23:59:59"


def resultado(page) -> str:
    return page.evaluate(
        """() => {
            const t = document.body.innerText || '';
            const i = t.indexOf('Resultado');
            return i < 0 ? '(sem bloco Resultado)' : t.slice(i, i + 600).replace(/\\s+/g, ' ');
        }"""
    )


def esperar_tabela(page, segundos: int = 40) -> str:
    fim = time.time() + segundos
    texto = ""
    while time.time() < fim:
        texto = resultado(page)
        if "Página" in texto and "Nenhum Registro" not in texto:
            break
        time.sleep(1)
    return texto


def ouvir_envios(page, envios: list) -> None:
    def ao_pedir(req):
        if req.method != "POST" or "finRelVendaOperadorPDV2" not in req.url:
            return
        campos = [
            (k, v) for k, v in parse_qsl(req.post_data or "", keep_blank_values=True)
            if "ViewState" not in k
        ]
        envios.append(campos)
    page.on("request", ao_pedir)


def digitar_datas(page) -> None:
    for campo, valor in (("form:inputDataInicial_input", DIA_INI), ("form:inputDataFinal_input", DIA_FIM)):
        el = page.locator(f'[id="{campo}"]')
        el.click()
        el.press("Control+a")
        el.type(valor, delay=20)
        el.press("Tab")
        s.wait_bip_idle(page, timeout=15)
    page.keyboard.press("Escape")


def variante(page, nome: str, passos) -> None:
    print(f"\n==== {nome} ====")
    envios: list = []
    ouvir_envios(page, envios)
    s.go_to_report(page)
    envios.clear()
    passos(page)
    texto = esperar_tabela(page)
    datas = page.evaluate(
        """() => [document.getElementById('form:inputDataInicial_input')?.value,
                  document.getElementById('form:inputDataFinal_input')?.value]"""
    )
    print(f"datas na tela: {datas}")
    for n, campos in enumerate(envios[-4:], 1):
        print(f"envio {n}: {campos}")
    print(f"resultado: {texto}")


def v1_como_coletor(page):
    s.selecionar_indice(page, "Pessoas Atendidas")
    s.set_report_dates(page, DIA_INI, DIA_FIM)
    time.sleep(0.5)
    s.click_pesquisar(page)
    s.wait_bip_idle(page, timeout=45)


def v2_datas_antes(page):
    s.set_report_dates(page, DIA_INI, DIA_FIM)
    s.selecionar_indice(page, "Pessoas Atendidas")
    s.click_pesquisar(page)
    s.wait_bip_idle(page, timeout=45)


def v3_digitando(page):
    s.selecionar_indice(page, "Pessoas Atendidas")
    digitar_datas(page)
    page.get_by_role("button", name=re.compile("Pesquisar")).click()
    s.wait_bip_idle(page, timeout=45)


def v4_quantidade(page):
    s.set_report_dates(page, DIA_INI, DIA_FIM)
    s.click_pesquisar(page)
    s.wait_bip_idle(page, timeout=45)


def main() -> None:
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1600, "height": 1000})
        page = context.new_page()
        page.set_default_timeout(30000)
        s.login(page)
        s.select_store(page, LOJA["bip_name"])
        for nome, passos in (
            ("v4: Quantidade (controle, como as vendas)", v4_quantidade),
            ("v1: Pessoas como o coletor fazia", v1_como_coletor),
            ("v2: Pessoas com as datas antes do Índice", v2_datas_antes),
            ("v3: Pessoas digitando as datas e clicando no botão", v3_digitando),
        ):
            try:
                variante(page, nome, passos)
            except Exception as e:
                print(f"falhou: {e}")
        browser.close()


if __name__ == "__main__":
    main()
