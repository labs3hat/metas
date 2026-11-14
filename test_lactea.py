"""Testes da leitura da Saída de Composição (lactea.py). Rodar: python -m unittest test_lactea"""

import unittest

from lactea import RelatorioInvalido, codigo_da_franquia, ler_linhas, somar_por_sabor

CAB = ["Descrição", "Unidade de Medida", "Tipo de Saída", "Qtd. Venda"]


class Folha:
    def __init__(self, linhas):
        self.linhas = linhas
        self.nrows = len(linhas)
        self.ncols = max(len(l) for l in linhas)

    def cell_value(self, r, c):
        linha = self.linhas[r]
        return linha[c] if c < len(linha) else ""


class TestLactea(unittest.TestCase):
    def test_soma_venda_e_criacao_por_sabor(self):
        # Os números do print da Rose (01–14/11/2026).
        f = Folha([
            CAB,
            ["Bebida Láctea Uht Baunilha", "LT", "Criação de Produtos", 11.575],
            ["Bebida Láctea Uht Baunilha", "LT", "Venda", "440,9165"],
            ["Bebida Láctea Uht Chocolate", "LT", "Criação de Produtos", 2.455],
            ["Bebida Láctea Uht Chocolate", "LT", "Venda", 103.17],
            ["Açaí", "KG", "Venda", 12.0],
        ])
        r = {i.sabor: i for i in somar_por_sabor(ler_linhas(f))}
        self.assertEqual(r["baunilha"].litros, 452.4915)
        self.assertEqual(r["chocolate"].litros, 105.625)
        self.assertEqual(r["baunilha"].unidade, "LT")
        self.assertEqual(len(r["baunilha"].linhas), 2)

    def test_sabor_sem_linha_fica_de_fora(self):
        f = Folha([CAB, ["Bebida Láctea Uht Baunilha", "LT", "Venda", 30]])
        r = somar_por_sabor(ler_linhas(f))
        self.assertEqual([i.sabor for i in r], ["baunilha"])

    def test_cabecalho_errado_aborta(self):
        with self.assertRaises(RelatorioInvalido):
            ler_linhas(Folha([["Descrição", "Qtd"], ["x", 1]]))

    def test_tipo_de_saida_desconhecido_aborta(self):
        f = Folha([CAB, ["Bebida Láctea Uht Baunilha", "LT", "Devolução", 3]])
        with self.assertRaises(RelatorioInvalido):
            ler_linhas(f)

    def test_quantidade_em_branco_aborta(self):
        f = Folha([CAB, ["Bebida Láctea Uht Baunilha", "LT", "Venda", ""]])
        with self.assertRaises(RelatorioInvalido):
            ler_linhas(f)

    def test_produto_e_tipo_repetidos_abortam(self):
        f = Folha([CAB,
                   ["Bebida Láctea Uht Baunilha", "LT", "Venda", 3],
                   ["Bebida Lactea UHT Baunilha ", "LT", "Venda", 4]])
        with self.assertRaises(RelatorioInvalido):
            ler_linhas(f)

    def test_unidades_diferentes_no_mesmo_sabor_abortam(self):
        f = Folha([CAB,
                   ["Bebida Láctea Uht Baunilha", "LT", "Venda", 3],
                   ["Bebida Láctea Uht Baunilha", "CX", "Criação de Produtos", 1]])
        with self.assertRaises(RelatorioInvalido):
            somar_por_sabor(ler_linhas(f))

    def test_bebida_lactea_sem_sabor_conhecido_aborta(self):
        f = Folha([CAB, ["Bebida Láctea Uht Morango", "LT", "Venda", 3]])
        with self.assertRaises(RelatorioInvalido):
            somar_por_sabor(ler_linhas(f))

    def test_codigo_da_franquia_ignora_o_nome_do_shopping(self):
        self.assertEqual(codigo_da_franquia("MARINGÁ 09 - BONIFÁCIO MALL (PR)"), "MARINGA 09")
        self.assertEqual(codigo_da_franquia("Maringá 09 (PR)"), "MARINGA 09")
        self.assertEqual(codigo_da_franquia("CURITIBA 07 (PR)"), "CURITIBA 07")
        self.assertNotEqual(codigo_da_franquia("Maringá 07 - Shopping Avenida Center (PR)"),
                            codigo_da_franquia("MARINGÁ 03 - SHOPPING AVENIDA CENTER MARINGÁ (PR)"))


if __name__ == "__main__":
    unittest.main()
