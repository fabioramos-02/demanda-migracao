"""Preenche coluna 'Resposta' do xlsx de pendências do CMS com tabela markdown de etapas.

Fluxo:
1. Lê xlsx 7 colunas (entidade | source_id | orgao | ativo | registro | motivos | Resposta).
2. Filtra entidade='serviço' AND ativo='sim' AND motivos contém 'sem etapa'.
3. Pra cada id, consulta banco PostgreSQL admin_prd (ou --inspect pra descobrir schema).
4. Formata etapas como tabela markdown e grava na coluna resposta.
5. Backup .bak antes de salvar.

Uso:
    python preencher_etapas.py --inspect                      # lista tabelas gerenciamento_*
    python preencher_etapas.py --input ARQ.xlsx --dry-run     # mostra o que seria preenchido
    python preencher_etapas.py --input ARQ.xlsx --limit 5     # processa 5 primeiros
    python preencher_etapas.py --input ARQ.xlsx               # full
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
from html import unescape
from html.parser import HTMLParser
from pathlib import Path

import psycopg2
from dotenv import load_dotenv
from openpyxl import load_workbook

# ---------- schema (ajustar após --inspect confirmar nome real) ----------
# Hipótese H1: tabela dedicada. Trocar nomes conforme banco.
QUERY_ETAPAS = """
    SELECT ordem, titulo, canal_prestacao, conteudo
    FROM gerenciamento_jornada
    WHERE servico_id = %s
    ORDER BY ordem
"""

FILTRO_RESPOSTA = "sem etapa"  # substring case-insensitive na coluna resposta
CANAL_DEFAULT = "Online"

# ---------- HTML -> texto (copiado de extracao-carta/html_text.py) ----------
_QUEBRA = {"p", "br", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6"}
_ESPACOS = re.compile(r"[ \t]+")
_LINHAS = re.compile(r"\n{3,}")


class _Extrator(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.partes: list[str] = []

    def handle_data(self, data: str) -> None:
        self.partes.append(data)

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in _QUEBRA:
            self.partes.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in _QUEBRA:
            self.partes.append("\n")


def to_text(html: str | None) -> str:
    if not html:
        return ""
    p = _Extrator()
    p.feed(html)
    p.close()
    texto = _ESPACOS.sub(" ", unescape("".join(p.partes)))
    texto = "\n".join(linha.strip() for linha in texto.split("\n"))
    # etapa inteira cabe numa célula; sem quebras pra não quebrar markdown da tabela
    return _LINHAS.sub(" ", texto).replace("\n", " ").strip()


# ---------- DB ----------
def conectar():
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        dbname=os.getenv("DB_NAME"),
        client_encoding="UTF8",
    )


def inspecionar(conn) -> None:
    """Lista tabelas gerenciamento_* e colunas prováveis de etapas."""
    with conn.cursor() as cur:
        cur.execute("""
            SELECT table_name
            FROM information_schema.tables
            WHERE table_schema='public' AND table_name LIKE 'gerenciamento_%'
            ORDER BY table_name
        """)
        tabelas = [r[0] for r in cur.fetchall()]
        print(f"[tabelas gerenciamento_*] {len(tabelas)}")
        for t in tabelas:
            print(f"  - {t}")

        candidatos = [t for t in tabelas if any(k in t for k in ("etapa", "jornada", "passo"))]
        print(f"\n[candidatos a tabela de etapas] {candidatos or 'NENHUM'}")
        for t in candidatos:
            cur.execute(
                "SELECT column_name, data_type FROM information_schema.columns "
                "WHERE table_schema='public' AND table_name=%s ORDER BY ordinal_position",
                (t,),
            )
            print(f"\n[{t}]")
            for col, tipo in cur.fetchall():
                print(f"  {col:30} {tipo}")

        # fallback: coluna jsonb 'jornada' em gerenciamento_servicos
        cur.execute("""
            SELECT column_name, data_type FROM information_schema.columns
            WHERE table_schema='public' AND table_name='gerenciamento_servicos'
              AND column_name IN ('jornada','etapas','passos')
        """)
        extras = cur.fetchall()
        if extras:
            print(f"\n[colunas JSON em gerenciamento_servicos] {extras}")


def buscar_etapas(conn, servico_id: int) -> list[dict]:
    with conn.cursor() as cur:
        cur.execute(QUERY_ETAPAS, (servico_id,))
        rows = cur.fetchall()
    return [
        {
            "ordem": r[0],
            "titulo": r[1],
            "canal": " / ".join(r[2]) if r[2] else CANAL_DEFAULT,
            "conteudo": to_text(r[3]),
        }
        for r in rows
    ]


# ---------- xlsx ----------
def mapear_colunas(ws) -> dict[str, int]:
    """Mapeia cabeçalhos da linha 1 -> índice 1-based. Case-insensitive."""
    header = {}
    for col_idx, cell in enumerate(ws[1], start=1):
        if cell.value:
            header[str(cell.value).strip().lower()] = col_idx
    return header


def formatar_tabela(etapas: list[dict]) -> str:
    linhas = [
        f"| {e['ordem']} | {e['titulo']} | {e['canal']} | {e['conteudo']} |"
        for e in etapas
    ]
    return "\n".join(linhas) + "\n"


# ---------- main ----------
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", help="xlsx de entrada (7 colunas CMS)")
    ap.add_argument("--inspect", action="store_true", help="lista tabelas do banco e sai")
    ap.add_argument("--dry-run", action="store_true", help="não grava, só reporta")
    ap.add_argument("--limit", type=int, help="processa só os N primeiros filtrados")
    args = ap.parse_args()

    load_dotenv()

    if args.inspect:
        with conectar() as conn:
            inspecionar(conn)
        return 0

    if not args.input:
        ap.error("--input é obrigatório salvo em --inspect")

    xlsx = Path(args.input)
    if not xlsx.exists():
        print(f"[erro] arquivo não encontrado: {xlsx}", file=sys.stderr)
        return 1

    wb = load_workbook(xlsx)
    ws = wb.active
    cols = mapear_colunas(ws)

    obrigatorias = ("entidade", "source_id", "ativo", "motivos", "resposta")
    faltam = [c for c in obrigatorias if c not in cols]
    if faltam:
        print(f"[erro] cabeçalhos faltando: {faltam}", file=sys.stderr)
        print(f"[cabeçalhos achados] {list(cols)}", file=sys.stderr)
        return 1

    alvos = []
    for row_idx in range(2, ws.max_row + 1):
        entidade = (ws.cell(row_idx, cols["entidade"]).value or "").strip().lower()
        ativo = (ws.cell(row_idx, cols["ativo"]).value or "").strip().lower()
        motivos = str(ws.cell(row_idx, cols["motivos"]).value or "").lower()
        resposta = ws.cell(row_idx, cols["resposta"]).value
        if entidade == "serviço" and ativo == "sim" and FILTRO_RESPOSTA in motivos and not resposta:
            alvos.append(row_idx)

    if args.limit:
        alvos = alvos[: args.limit]

    print(f"[alvos] {len(alvos)} linhas pra preencher")
    def _sid(r: int) -> int:
        return int(float(str(ws.cell(r, cols["source_id"]).value)))

    if args.dry_run:
        for r in alvos[:20]:
            print(f"  linha {r} id={_sid(r)} orgao={ws.cell(r, cols.get('orgao', 0)).value}")
        if len(alvos) > 20:
            print(f"  ... (+{len(alvos)-20})")
        return 0

    ok = sem_etapas = erros = 0
    with conectar() as conn:
        for r in alvos:
            sid = _sid(r)
            try:
                etapas = buscar_etapas(conn, sid)
            except Exception as e:
                erros += 1
                print(f"[erro] linha {r} id={sid}: {e}", file=sys.stderr)
                continue
            if not etapas:
                sem_etapas += 1
                print(f"[vazio] linha {r} id={sid}: banco sem etapas")
                continue
            ws.cell(r, cols["resposta"]).value = formatar_tabela(etapas)
            ok += 1
        conn.rollback()  # read-only, descarta transação

    backup = xlsx.with_suffix(xlsx.suffix + ".bak")
    shutil.copy2(xlsx, backup)
    wb.save(xlsx)
    print(f"\n[fim] preenchidos={ok} sem_etapas_no_banco={sem_etapas} erros={erros}")
    print(f"[backup] {backup}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
