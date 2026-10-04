"""Diagnóstico (só leitura): roda coletar_operadores de verdade em SJP 1, dia
03/10/2026, com um cliente falso que imprime em vez de gravar. Esperado pela
tela (04/10): Totem 156, Eduarda 04 115, Ericson 06 31, MARIANA 5, Erika 05 3."""

import tempfile

from playwright.sync_api import sync_playwright

import scraper as s

LOJA = next(l for l in s.LOJAS if l["key"] == "sjp1")


class ClienteQueSoImprime:
    def enviar_operadores(self, loja: str, data_br: str, operadores: list[dict]) -> bool:
        print(f"ENVIARIA {loja} {data_br}:")
        for o in operadores:
            tm = o["receita"] / o["pessoas"] if o["pessoas"] else 0
            print(f"   {o['nome']}: {o['pessoas']} cliente(s), R$ {o['receita']:.2f}, TM {tm:.2f}")
        print(f"   total: {sum(o['pessoas'] for o in operadores)} cliente(s), "
              f"R$ {sum(o['receita'] for o in operadores):.2f}")
        return True


def main() -> None:
    with sync_playwright() as pw, tempfile.TemporaryDirectory() as pasta:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1600, "height": 1000})
        page = context.new_page()
        page.set_default_timeout(30000)
        s.login(page)
        s.select_store(page, LOJA["bip_name"])
        s.coletar_operadores(page, LOJA, ClienteQueSoImprime(), ["03/10/2026"], pasta)
        browser.close()


if __name__ == "__main__":
    main()
