# Metas — Chiquinho Sorvetes · Grupo 3HAT

Dashboard interno de metas e ranking de operadores das lojas Chiquinho.

## Estrutura

```
metas/
├── index.html                    ← Dashboard (GitHub Pages)
├── scraper.py                    ← Automação BIP360 → Google Sheets
├── requirements.txt              ← Dependências Python
└── .github/
    └── workflows/
        └── daily.yml             ← Agendamento automático (06h00 BRT)
```

## Como funciona

1. **Todo dia às 06h00** o GitHub Actions executa `scraper.py`
2. O scraper faz login no BIP360, baixa os relatórios de cada loja e atualiza o Google Sheets
3. O dashboard (`index.html`) lê o Google Sheets e exibe os dados automaticamente

## Planilha Google Sheets

ID: `1qU8Ny_OqoF4VrI0IU4JuuRvoNnmBOMSf9JkFs1h4PRY`

| Aba | GID | Conteúdo |
|---|---|---|
| Metas Operacionais | 154705838 | Metas + Realizado por loja/mês |
| Metas TM suporte | 76128240 | Ticket Médio Meta + Realizado |
| Metas Operador suporte | 786119570 | Vendas por operador/loja/mês |

## Secrets necessários no GitHub

| Secret | Descrição |
|---|---|
| `BIP_USER` | Usuário de login do BIP360 |
| `BIP_PASS` | Senha do BIP360 |
| `GOOGLE_CREDENTIALS` | JSON da Service Account Google |

## Lojas

| Código | Loja |
|---|---|
| SJP 1 | Shopping São José — São José dos Pinhais |
| CTBA 3 | Loja Shopping Estação — Curitiba |
| CTBA 5 | Loja Shopping Jockey — Curitiba |
| CTBA 7 | Quiosque Shopping Palladium — Curitiba |
| CTBA 11 | Loja Shopping Palladium — Curitiba |
| CL 2 | City Center Outlet — Campo Largo |
| MGA 3 | Loja Avenida Center — Maringá |
| MGA 5 | Loja Shopping Cidade — Maringá |
| MGA 7 | Quiosque Avenida Center — Maringá |
| MGA 8 | Quiosque Havan — Maringá |

## Escrita dupla no banco do sistema novo (desde 24/09/2026)

Os dois scrapers continuam gravando na planilha exatamente como antes e, depois,
mandam o mesmo dado para o banco do sistema novo (`banco_coleta.py` →
função `ingerir-coleta` do Supabase, repositório `labs3hat/metas-41944e89`).

- Liga com o secret `COLETA_TOKEN` do repositório. Sem ele, nada muda.
- Falha no banco **nunca** derruba a planilha: vira aviso no log.
- Idempotente: reenviar um dia substitui o dia no banco.
- Dia que está na planilha mas falta no banco é recoletado só para o banco,
  no máximo 2 por loja por execução (`COLETA_BACKFILL_MAX_POR_LOJA`).
