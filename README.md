# demanda-migracao

Preenche a coluna `resposta` do xlsx exportado do CMS do Portal MS com tabelas markdown de etapas, consultando o banco PostgreSQL `admin_prd`.

CMS exige **≥1 etapa por canal** pra publicar. Serviços ativos sem etapa ficam marcados como `"sem etapa (o CMS exige ao menos uma por canal)"` — esse script resolve em lote.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
cp .env.example .env   # editar com credenciais reais (ver extracao-carta/.env)
```

## Uso

**1. Confirmar schema do banco** (nome real da tabela de etapas):

```powershell
python preencher_etapas.py --inspect
```

Saída lista tabelas `gerenciamento_*` e destaca candidatas (`etapa`, `jornada`, `passo`). Ajustar constante `QUERY_ETAPAS` no topo do script conforme resultado.

**2. Dry-run** (não grava, só reporta linhas-alvo):

```powershell
python preencher_etapas.py --input "arquivo-cms.xlsx" --dry-run
```

**3. Subset pequeno pra validar formato**:

```powershell
python preencher_etapas.py --input "arquivo-cms.xlsx" --limit 5
```

Abrir no Excel, conferir que a coluna `resposta` ganhou tabela markdown no formato:

```
| 1 | Título | Online | Conteúdo da etapa. |
| 2 | ... | ... | ... |
```

**4. Full**:

```powershell
python preencher_etapas.py --input "arquivo-cms.xlsx"
```

Gera backup `.bak` automático antes de salvar.

## Formato esperado do xlsx de entrada

Cabeçalho na linha 1 com colunas (case-insensitive): `tipo | id | órgão | ativo | título | status | resposta`.

Filtro aplicado: `tipo == "serviço"` AND `ativo == "sim"` AND `resposta` contém `"sem etapa"`.

## Riscos conhecidos

- **Schema `gerenciamento_jornada` não confirmado** — rodar `--inspect` primeiro. Se tabela tiver outro nome ou etapas forem JSON em `gerenciamento_servicos.jornada`, editar `QUERY_ETAPAS`.
- **Canal pode não existir no banco** — fallback é `"Online"` (constante `CANAL_DEFAULT`).
- Serviços sem etapas no banco ficam logados como `[vazio]` e **não** sobrescrevem a célula — viram backlog manual.
