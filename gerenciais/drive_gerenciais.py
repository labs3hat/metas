"""Acesso ao Google Drive para as planilhas gerenciais das lojas.

Usa a mesma conta de serviço do scraper (secret GOOGLE_CREDENTIALS) e a API
REST do Drive v3 por `AuthorizedSession` — sem dependência nova além do
google-auth que já está no requirements.

Três operações, todas com timeout e sem efeito colateral escondido:
  - listar_planilhas(pasta): os .xlsm "Gestao Financeira ... 2026" da pasta
  - baixar(arquivo, destino): conteúdo atual
  - subir_nova_versao(arquivo, caminho): antes de gravar, marca a revisão atual
    como "manter para sempre" (fica em Gerenciar versões); depois envia o novo
    conteúdo como nova revisão do MESMO arquivo (id e link não mudam).
"""
from __future__ import annotations

import json
import logging
import os
from dataclasses import dataclass
from pathlib import Path

from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

log = logging.getLogger("gerenciais.drive")

API = "https://www.googleapis.com/drive/v3"
UPLOAD = "https://www.googleapis.com/upload/drive/v3"
TIMEOUT = 120
MIME_XLSM = "application/vnd.ms-excel.sheet.macroenabled.12"


@dataclass(frozen=True)
class ArquivoDrive:
    id: str
    nome: str
    tamanho: int
    modificado_em: str
    revisao_atual: str | None


def sessao() -> AuthorizedSession:
    creds = Credentials.from_service_account_info(
        json.loads(os.environ["GOOGLE_CREDENTIALS"]),
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return AuthorizedSession(creds)


def _erro(resposta, acao: str) -> None:
    if resposta.status_code >= 400:
        raise RuntimeError(f"Drive: {acao} falhou (HTTP {resposta.status_code}): {resposta.text[:300]}")


def listar_planilhas(s: AuthorizedSession, pasta_id: str) -> list[ArquivoDrive]:
    """Arquivos .xlsm 'Gestao Financeira ... 2026' direto na pasta (sem subpastas)."""
    q = (
        f"'{pasta_id}' in parents and trashed = false "
        "and name contains 'Gestao Financeira' and name contains '2026'"
    )
    out: list[ArquivoDrive] = []
    token = None
    while True:
        params = {
            "q": q,
            "fields": "nextPageToken, files(id, name, size, modifiedTime, headRevisionId, mimeType)",
            "pageSize": 100,
            "supportsAllDrives": "true",
            "includeItemsFromAllDrives": "true",
        }
        if token:
            params["pageToken"] = token
        r = s.get(f"{API}/files", params=params, timeout=TIMEOUT)
        _erro(r, "listar pasta")
        dados = r.json()
        for f in dados.get("files", []):
            if not f["name"].endswith(".xlsm"):
                continue
            out.append(ArquivoDrive(
                id=f["id"], nome=f["name"], tamanho=int(f.get("size", 0)),
                modificado_em=f.get("modifiedTime", ""), revisao_atual=f.get("headRevisionId"),
            ))
        token = dados.get("nextPageToken")
        if not token:
            break
    log.info("Pasta %s: %d planilha(s) 'Gestao Financeira ... 2026.xlsm'", pasta_id, len(out))
    return sorted(out, key=lambda a: a.nome)


def baixar(s: AuthorizedSession, arquivo: ArquivoDrive, destino: Path) -> Path:
    r = s.get(f"{API}/files/{arquivo.id}", params={"alt": "media", "supportsAllDrives": "true"},
              timeout=TIMEOUT)
    _erro(r, f"baixar {arquivo.nome}")
    destino.write_bytes(r.content)
    if arquivo.tamanho and len(r.content) != arquivo.tamanho:
        raise RuntimeError(f"{arquivo.nome}: baixados {len(r.content)} bytes, Drive diz {arquivo.tamanho}")
    log.info("  baixado %s (%d KB)", arquivo.nome, len(r.content) // 1024)
    return destino


def preservar_revisao_atual(s: AuthorizedSession, arquivo: ArquivoDrive) -> None:
    """Marca a revisão vigente como keepForever: o original fica em Gerenciar versões
    mesmo depois que o Drive começar a limpar revisões antigas."""
    if not arquivo.revisao_atual:
        log.warning("  %s: sem headRevisionId; sigo sem marcar a revisão anterior", arquivo.nome)
        return
    r = s.patch(
        f"{API}/files/{arquivo.id}/revisions/{arquivo.revisao_atual}",
        json={"keepForever": True}, timeout=TIMEOUT,
    )
    _erro(r, f"preservar revisão de {arquivo.nome}")
    log.info("  revisão anterior de %s preservada (%s)", arquivo.nome, arquivo.revisao_atual)


def subir_nova_versao(s: AuthorizedSession, arquivo: ArquivoDrive, caminho: Path) -> str:
    """Envia `caminho` como nova revisão do mesmo arquivo. Devolve o novo headRevisionId."""
    preservar_revisao_atual(s, arquivo)
    dados = caminho.read_bytes()
    r = s.patch(
        f"{UPLOAD}/files/{arquivo.id}",
        params={"uploadType": "media", "supportsAllDrives": "true",
                "fields": "id, headRevisionId, modifiedTime, size"},
        data=dados, headers={"Content-Type": MIME_XLSM}, timeout=TIMEOUT * 3,
    )
    _erro(r, f"subir nova versão de {arquivo.nome}")
    meta = r.json()
    if int(meta.get("size", len(dados))) != len(dados):
        raise RuntimeError(f"{arquivo.nome}: Drive gravou {meta.get('size')} bytes, enviados {len(dados)}")
    log.info("  nova versão de %s gravada (revisão %s, %d KB)", arquivo.nome, meta.get("headRevisionId"), len(dados) // 1024)
    return str(meta.get("headRevisionId"))
