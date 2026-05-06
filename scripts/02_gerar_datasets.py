"""
gerar_datasets_analise.py
==========================
Gera 5 datasets analíticos a partir da tabela consolidada atualizada.

Saída em resultados/datasets/:
  a) base_nivel_investimento.csv  — uma linha por fonte de investimento por obra
  b) base_nivel_produtora.csv     — uma linha por produtora independente (com clusters)
  c) base_nivel_chamada.csv       — uma linha por chamada FSA
  d) base_nivel_obra.csv          — uma linha por obra, investimentos como colunas
  e) base_nivel_direcao.csv       — uma linha por diretor (novo)

Enriquecimentos aplicados diretamente:
  - Renúncia fiscal: match 3 camadas título+CNPJ via projetos-com-renuncia-fiscal
  - Crítica: merge por CPB via critica_obras.csv
  - Citações acadêmicas: merge por CPB→diretor via citacoes_diretores.csv
  - Prestígio: merge por CPB→diretor via prestigio_diretores.csv
  - Festivais (diretor): merge por nome via perfil_festivais_diretores.csv

Uso:
    python gerar_datasets_analise.py
"""

import os
import re
import pathlib
import unicodedata
import pandas as pd
import numpy as np

ROOT     = pathlib.Path(__file__).parent.parent
DIR_OUT  = ROOT / "resultados" / "datasets"
DIR_OUT.mkdir(parents=True, exist_ok=True)

MASTER             = ROOT / "resultados" / "tabela_consolidada_obras.xlsx"
FSA_CSV            = ROOT / "raw" / "projetos-fsa.csv"
PROD_CSV           = ROOT / "raw" / "produtores-de-obras-nao-publicitarias-brasileiras.csv"
AGENTES            = ROOT / "raw" / "agentes-economicos-regulares.csv"
INDEP              = ROOT / "raw" / "produtoras-independentes.csv"
IPCA_CSV           = ROOT / "tabelas_apoio" / "deflator_ipca_base2024.csv"
CLUSTERS           = ROOT / "tabelas_apoio" / "Clusters_produtoras.xlsx"
SALIC_CSV          = ROOT / "raw" / "projetos-com-renuncia-fiscal.csv"
CRITICA_CSV        = ROOT / "dados" / "critica_obras.csv"
DIRETORES_CSV      = ROOT / "raw" / "diretores-de-obras-nao-publicitarias-brasileiras.csv"
CITACOES_CSV       = ROOT / "dados" / "citacoes_diretores.csv"
PRESTIGIO_CSV      = ROOT / "dados" / "prestigio_diretores.csv"
PERFIL_FESTIV_CSV  = ROOT / "dados" / "perfil_festivais_diretores.csv"

ENC = "utf-8-sig"

# ── Helpers ────────────────────────────────────────────────────────────────────

def _read(path, **kw):
    return pd.read_csv(path, sep=None, engine="python", encoding=ENC, dtype=str, **kw).fillna("")

def _num(s):
    try:
        return float(str(s).replace(",", ".").strip())
    except Exception:
        return 0.0

def _int(s):
    try:
        v = float(str(s).replace(",", ".").strip())
        return int(v) if not np.isnan(v) else 0
    except Exception:
        return 0

def _cnpj_clean(s):
    return re.sub(r"\D", "", str(s))

def _norm(s):
    """Normaliza título: UPPER, sem acentos, só A-Z 0-9 espaço."""
    if not s:
        return ""
    s = str(s).upper()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _prefix4(s):
    return " ".join(s.split()[:4])

def _brl(s):
    """R$ 1.234,56 → 1234.56"""
    try:
        return float(str(s).replace("R$", "").replace(".", "").replace(",", ".").strip())
    except Exception:
        return 0.0


# ── 1. Carregar master table ───────────────────────────────────────────────────
print("Lendo tabela consolidada...")
df = pd.read_excel(MASTER, sheet_name="Obras", dtype=str).fillna("")
print(f"  {len(df)} obras × {len(df.columns)} colunas")

# PREPARACAO DE PERIODOS
df["Ano"] = pd.to_numeric(df["Ano"], errors="coerce")

# Dataset RESTRITO: 2014-2023 (para análise de chamadas FSA)
df_restrito = df[(df["Ano"] >= 2014) & (df["Ano"] <= 2023)].copy()
print(f"  Periodo restrito 2014-2023: {len(df)} obras antes, {len(df_restrito)} obras depois")

# Dataset AMPLO: sem filtro de período (para análise de produtoras - espectro amplo)
df_amplo = df.copy()
print(f"  Periodo amplo (todo historico): {len(df_amplo)} obras")

# Por enquanto trabalhar com df_restrito para gerar datasets de chamadas
df = df_restrito

# Colunas numéricas essenciais
_NUM_COLS = [
    "Valor FSA (R$)", "Renúncia Art.3/3-A/39 (R$)", "Renúncia Outros Mec. (R$)",
    "Bilheteria Nominal (R$)", "Bilheteria Deflac. (R$)", "Estimativa Outras Janelas (R$)",
    "Valor FSA Deflac. (R$2024)", "Renúncia Total Deflac. (R$2024)",
    "Investimento Total Deflac. (R$2024)",
    "Outras Janelas Deflac. (R$2024)", "Outras Janelas Nominal (R$)",
    "ROI Dom. FSA (deflac)", "ROI Dom. Total (deflac)",
    "ROI Internacional (0-100)", "Pontuação Festivais",
    "Adm. EU — Lumière", "VOD Intl — N Plataformas", "VOD Intl — N Países",
    "Total Países Alcançados", "CRITICA_INDICE_1_5", "IMDB_RATING",
]
for c in _NUM_COLS:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c].str.replace(",", "."), errors="coerce").fillna(0)

def _ensure_col(col, fallback_col=None, default=0):
    if col not in df.columns:
        df[col] = df[fallback_col].fillna(default) if fallback_col and fallback_col in df.columns else default

_ensure_col("Valor FSA Deflac. (R$2024)",          "Valor FSA (R$)")
_ensure_col("Renúncia Total Deflac. (R$2024)",     "Renúncia Art.3/3-A/39 (R$)")
_ensure_col("Investimento Total Deflac. (R$2024)")
_ensure_col("Outras Janelas Deflac. (R$2024)",     "Estimativa Outras Janelas (R$)")
_ensure_col("Outras Janelas Nominal (R$)")
_ensure_col("ROI Dom. FSA (deflac)")
_ensure_col("ROI Dom. Total (deflac)")
_ensure_col("Total Países Alcançados")
_ensure_col("Países Festivais", default="")
_ensure_col("Países Lumière", default="")
_ensure_col("Países VOD Europa", default="")
_ensure_col("Países Lista", default="")

df["_receita_deflac"] = df["Bilheteria Deflac. (R$)"]


# ── 2. CNPJ por obra ──────────────────────────────────────────────────────────
print("Carregando CNPJ das obras...")
cnpj_map = {}
import glob as _glob
for _f in sorted(_glob.glob(str(ROOT / "raw" / "obras-nao-pub-brasileiras-csv" / "*.csv"))):
    try:
        _df = pd.read_csv(_f, sep=None, engine="python", encoding=ENC, dtype=str,
                          usecols=lambda c: c in ["CPB", "CNPJ_REQUERENTE"]).fillna("")
        for _, _r in _df.iterrows():
            _cpb = _r["CPB"].strip()
            if _cpb and _cpb not in cnpj_map:
                cnpj_map[_cpb] = _r.get("CNPJ_REQUERENTE", "").strip()
    except Exception:
        pass
print(f"  CPBs com CNPJ: {len(cnpj_map)}")
df["CNPJ_produtora"] = df["CPB"].map(cnpj_map).fillna("")


# ── 3. Agentes econômicos (razão social, UF) ──────────────────────────────────
print("Carregando agentes econômicos...")
agentes_map = {}
try:
    _ag = _read(AGENTES, usecols=lambda c: c in [
        "CNPJ", "RAZAO_SOCIAL", "UF", "CLASSIFICACAO_AGENTE_ECONOMICO", "BRASILEIRO_INDEPENDENTE"
    ])
    for _, _r in _ag.iterrows():
        _c = _cnpj_clean(_r.get("CNPJ", ""))
        if _c:
            agentes_map[_c] = {
                "razao_social":     _r.get("RAZAO_SOCIAL", ""),
                "UF":               _r.get("UF", ""),
                "classificacao":    _r.get("CLASSIFICACAO_AGENTE_ECONOMICO", ""),
                "br_independente":  _r.get("BRASILEIRO_INDEPENDENTE", ""),
            }
    print(f"  Agentes carregados: {len(agentes_map)}")
except Exception as e:
    print(f"  [AVISO] agentes: {e}")


# ── 4. CNPJs de produtoras independentes ──────────────────────────────────────
print("Carregando produtoras independentes...")
cnpjs_indep = set()
try:
    _ind = _read(INDEP)
    for c in _ind.columns:
        if "cnpj" in c.lower():
            cnpjs_indep = set(_cnpj_clean(v) for v in _ind[c] if v)
            break
    print(f"  Produtoras independentes: {len(cnpjs_indep)}")
except Exception as e:
    print(f"  [AVISO] produtoras-independentes: {e}")


# ── 5. Deflator IPCA ──────────────────────────────────────────────────────────
IPCA_FATORES = {}
try:
    _ip = _read(IPCA_CSV)
    for _, _r in _ip.iterrows():
        try:
            _ano = int(float(str(_r.get("ano", "")).replace(",", ".")))
            _fat = float(str(_r.get("fator_real_2024", "1")).replace(",", "."))
            IPCA_FATORES[_ano] = _fat
        except Exception:
            pass
except Exception:
    pass


# ── 6. Renúncia fiscal — match por título + CNPJ ──────────────────────────────
print("\nCarregando renúncia fiscal (SALIC)...")

_COLS_ART3  = ["CAPTADO_ART3", "CAPTADO_ART3A", "CAPTADO_ART39"]
_COLS_OUTRO = ["CAPTADO_ART1", "CAPTADO_ART1A", "CAPTADO_ART18", "CAPTADO_ART25",
               "CAPTADO_FUNCINES", "CAPTADO_EDITAL_ANCINE", "CAPTADO_PAR", "CAPTADO_PAQ",
               "CAPTADO_OUTROS_EDITAIS", "CAPTADO_LEI_ESTADUAL", "CAPTADO_LEI_MUNICIPAL",
               "CAPTADO_OUTRAS_FONTES", "CAPTADO_CONTRAPARTIDA", "CAPTADO_CONVERSAO"]

ren_by_title  = {}   # titulo_norm → list[{art3, outros, cnpj, titulo}]
ren_by_prefix = {}   # prefix4     → list[...]
_salic_loaded = False

try:
    df_salic = pd.read_csv(SALIC_CSV, sep=";", encoding="utf-8-sig", dtype=str).fillna("")
    print(f"  SALIC: {len(df_salic)} linhas × {len(df_salic.columns)} colunas")

    for _, row in df_salic.iterrows():
        titulo_raw = row.get("TITULO_PROJETO", "")
        if not titulo_raw:
            continue
        tnorm  = _norm(titulo_raw)
        cnpj   = _cnpj_clean(row.get("CNPJ_PROPONENTE", ""))
        art3   = sum(_brl(row.get(c, "")) for c in _COLS_ART3  if c in df_salic.columns)
        outros = sum(_brl(row.get(c, "")) for c in _COLS_OUTRO if c in df_salic.columns)
        entry  = {"art3": art3, "outros": outros, "cnpj": cnpj, "titulo": titulo_raw}

        ren_by_title.setdefault(tnorm, []).append(entry)
        ren_by_prefix.setdefault(_prefix4(tnorm), []).append(entry)

    _salic_loaded = True
    _com_ren = sum(1 for v in ren_by_title.values()
                   if any(e["art3"] + e["outros"] > 0 for e in v))
    print(f"  Títulos únicos: {len(ren_by_title)} | com valor > 0: {_com_ren}")

except Exception as e:
    print(f"  [AVISO] SALIC não carregado: {e}")


def _best_ren(candidates, cnpj_prod=""):
    """Desempata por CNPJ (preferência) ou maior valor."""
    if len(candidates) == 1:
        return candidates[0]
    if cnpj_prod:
        for r in candidates:
            if r["cnpj"] and r["cnpj"] == cnpj_prod:
                return r
    return max(candidates, key=lambda r: r["art3"] + r["outros"])


def _get_salic(titulo_norm, cnpj_prod=""):
    """3 camadas: exato → prefixo → containment 80%."""
    if not titulo_norm:
        return 0.0, 0.0

    # Tier 1: exato
    if titulo_norm in ren_by_title:
        m = _best_ren(ren_by_title[titulo_norm], cnpj_prod)
        return m["art3"], m["outros"]

    # Tier 2: prefixo 4 palavras
    pfx = _prefix4(titulo_norm)
    if pfx in ren_by_prefix:
        m = _best_ren(ren_by_prefix[pfx], cnpj_prod)
        return m["art3"], m["outros"]

    # Tier 3: containment 80%
    toks_obra = set(titulo_norm.split())
    candidates = []
    for tnorm_salic, entries in ren_by_title.items():
        toks_salic = set(tnorm_salic.split())
        inter = len(toks_obra & toks_salic)
        mn    = min(len(toks_obra), len(toks_salic))
        if mn > 0 and inter / mn >= 0.80:
            candidates.extend(entries)
    if candidates:
        m = _best_ren(candidates, cnpj_prod)
        return m["art3"], m["outros"]

    return 0.0, 0.0


if _salic_loaded:
    print("  Aplicando match renúncia fiscal nas obras...")
    _matches = 0
    for idx, row in df.iterrows():
        tnorm = _norm(str(row.get("Projeto", "")))
        cnpj  = _cnpj_clean(str(row.get("CNPJ_produtora", "")))
        art3, outros = _get_salic(tnorm, cnpj)
        if art3 + outros > 0:
            ano = _int(row.get("Ano", 0))
            fat = IPCA_FATORES.get(ano, 1.0)
            df.at[idx, "Renúncia Art.3/3-A/39 (R$)"]  = art3
            df.at[idx, "Renúncia Outros Mec. (R$)"]    = outros
            df.at[idx, "Renúncia Total Deflac. (R$2024)"] = round((art3 + outros) * fat, 2)
            _matches += 1
    print(f"  Obras com renúncia fiscal: {_matches}")

    # Recalcular total de investimento deflac
    _fsa_d = pd.to_numeric(df["Valor FSA Deflac. (R$2024)"], errors="coerce").fillna(0)
    _ren_d = pd.to_numeric(df["Renúncia Total Deflac. (R$2024)"], errors="coerce").fillna(0)
    df["Investimento Total Deflac. (R$2024)"] = _fsa_d + _ren_d
    df["_receita_deflac"] = (
        pd.to_numeric(df["Bilheteria Deflac. (R$)"], errors="coerce").fillna(0)
    )
    # Recalcular ROI Dom. Total com investimento total atualizado (FSA + Renúncia)
    _inv_tot_d = df["Investimento Total Deflac. (R$2024)"]
    _rec_d     = df["_receita_deflac"]
    df["ROI Dom. Total (deflac)"] = (_rec_d / _inv_tot_d).where(_inv_tot_d > 0).round(4)


# ── 7. Crítica — merge por CPB ─────────────────────────────────────────────────
print("\nCarregando crítica...")
_crit_map = {}   # CPB → {indice, fontes, n_fontes, confianca}

try:
    df_crit = pd.read_csv(CRITICA_CSV, sep=None, engine="python", encoding=ENC, dtype=str).fillna("")
    for _, row in df_crit.iterrows():
        cpb = row.get("CPB", "").strip()
        if not cpb:
            continue
        idx_val = row.get("CRITICA_INDICE_1_5", "")
        if idx_val and idx_val != "nan":
            _crit_map[cpb] = {
                "CRITICA_INDICE_1_5": idx_val,
                "CRITICA_FONTES":     row.get("CRITICA_FONTES", ""),
                "CRITICA_N_FONTES":   row.get("CRITICA_N_FONTES", "0"),
                "CRITICA_CONFIANCA":  row.get("CRITICA_CONFIANCA", ""),
            }
    print(f"  Obras com crítica: {len(_crit_map)}")
except Exception as e:
    print(f"  [AVISO] crítica: {e}")

# Aplicar na master
for col in ["CRITICA_INDICE_1_5", "CRITICA_FONTES", "CRITICA_N_FONTES", "CRITICA_CONFIANCA"]:
    if col not in df.columns:
        df[col] = ""

for idx, row in df.iterrows():
    cpb = str(row.get("CPB", "")).strip()
    if cpb in _crit_map:
        for col, val in _crit_map[cpb].items():
            df.at[idx, col] = val


# ── 8. Citações / Prestígio — merge CPB → diretor ────────────────────────────
print("\nCarregando diretores e citações...")

# 8a. CPB → lista de diretores (só brasileiros)
_cpb_diretores = {}   # CPB → [DIRETOR_NORM, ...]
try:
    df_dir = pd.read_csv(DIRETORES_CSV, sep=None, engine="python", encoding=ENC, dtype=str).fillna("")
    for _, row in df_dir.iterrows():
        cpb  = row.get("CPB", "").strip()
        nome = row.get("DIRETOR", "").strip()
        pais = row.get("PAIS_DIRETOR", "").strip().upper()
        if cpb and nome and pais == "BRASIL":
            _cpb_diretores.setdefault(cpb, []).append(_norm(nome))
    print(f"  CPBs com diretor BR: {len(_cpb_diretores)}")
except Exception as e:
    print(f"  [AVISO] diretores: {e}")

# 8b. Citações por nome normalizado
_cit_map = {}   # diretor_norm → dict
try:
    df_cit = pd.read_csv(CITACOES_CSV, sep=None, engine="python", encoding=ENC, dtype=str).fillna("")
    for _, row in df_cit.iterrows():
        nome = _norm(row.get("DIRETOR", ""))
        if nome:
            _cit_map[nome] = {
                "cita_n_papers":   row.get("CITA_N_PAPERS", ""),
                "cita_soma_cit":   row.get("CITA_SOMA_CIT", ""),
                "cita_max_cit":    row.get("CITA_MAX_CIT", ""),
                "cita_venues":     row.get("CITA_VENUES", ""),
                "prestigio_festiv_cit": row.get("PRESTIGIO_FESTIV", ""),
            }
    print(f"  Diretores com citações: {len(_cit_map)}")
except Exception as e:
    print(f"  [AVISO] citações: {e}")

# 8c. Prestígio / Wikidata (conjunto menor, mais curado)
_prest_map = {}   # diretor_norm → dict
try:
    df_prest = pd.read_csv(PRESTIGIO_CSV, sep=None, engine="python", encoding=ENC, dtype=str).fillna("")
    for _, row in df_prest.iterrows():
        nome = _norm(row.get("DIRETOR", ""))
        if nome:
            _prest_map[nome] = {
                "wikidata_id":       row.get("WIKIDATA_ID", ""),
                "prestigio_festiv":  row.get("PRESTIGIO_FESTIV", ""),
                "prestigio_filmes":  row.get("PRESTIGIO_FILMES", ""),
            }
    print(f"  Diretores com prestígio: {len(_prest_map)}")
except Exception as e:
    print(f"  [AVISO] prestígio: {e}")


def _busca_cit(cpb):
    """Retorna o melhor registro de citações para o CPB (maior CITA_N_PAPERS)."""
    diretores = _cpb_diretores.get(cpb, [])
    best = None
    best_n = -1
    for d in diretores:
        rec = _cit_map.get(d)
        if rec:
            n = _int(rec["cita_n_papers"])
            if n > best_n:
                best = rec
                best_n = n
    return best

def _busca_prest(cpb):
    """Retorna prestígio/wikidata para o CPB."""
    for d in _cpb_diretores.get(cpb, []):
        rec = _prest_map.get(d)
        if rec:
            return rec
    return None


# Aplicar na master df
for col in ["CITA_N_PAPERS", "CITA_SOMA_CIT", "CITA_MAX_CIT", "CITA_VENUES",
            "PRESTIGIO_FESTIV", "WIKIDATA_ID_DIR"]:
    if col not in df.columns:
        df[col] = ""

_cit_matches = 0
for idx, row in df.iterrows():
    cpb = str(row.get("CPB", "")).strip()
    rec_cit   = _busca_cit(cpb)
    rec_prest = _busca_prest(cpb)
    if rec_cit:
        df.at[idx, "CITA_N_PAPERS"] = rec_cit["cita_n_papers"]
        df.at[idx, "CITA_SOMA_CIT"] = rec_cit["cita_soma_cit"]
        df.at[idx, "CITA_MAX_CIT"]  = rec_cit["cita_max_cit"]
        df.at[idx, "CITA_VENUES"]   = rec_cit["cita_venues"]
        if not df.at[idx, "PRESTIGIO_FESTIV"]:
            df.at[idx, "PRESTIGIO_FESTIV"] = rec_cit.get("prestigio_festiv_cit", "")
        _cit_matches += 1
    if rec_prest:
        if rec_prest.get("prestigio_festiv"):
            df.at[idx, "PRESTIGIO_FESTIV"] = rec_prest["prestigio_festiv"]
        df.at[idx, "WIKIDATA_ID_DIR"] = rec_prest.get("wikidata_id", "")

print(f"  Obras com citações aplicadas: {_cit_matches}")

# Garantir colunas opcionais de diversidade
_OPT_DIV = ["GENERO_DIRETOR", "DIVERSIDADE_REGIONAL", "PROTAGONISTA_DIVERSO",
             "UF_REQUERENTE", "TEM_DIRETOR_PRESTIGIO", "IMDB_GENRES", "IMDB_RUNTIME",
             "PRESTIGIO_FESTIVAIS", "N_DIRETORES"]
for _c in _OPT_DIV:
    if _c not in df.columns:
        df[_c] = ""


# ═══════════════════════════════════════════════════════════════════════════════
# DATASET A — BASE NÍVEL INVESTIMENTO
# ═══════════════════════════════════════════════════════════════════════════════
print("\nGerando base_nivel_investimento.csv...")

_BASE_COLS = [
    "CPB", "Projeto", "Ano", "Chamada", "Categoria",
    "Bilheteria Nominal (R$)", "Bilheteria Deflac. (R$)",
    "Outras Janelas Deflac. (R$2024)", "Outras Janelas Nominal (R$)",
    "_receita_deflac",
    "ROI Dom. FSA (deflac)", "ROI Dom. Total (deflac)",
    "ROI Internacional (0-100)", "Total Países Alcançados",
    "Pontuação Festivais", "Adm. EU — Lumière",
    "VOD Intl — N Plataformas", "VOD Intl — N Países",
    "CRITICA_INDICE_1_5", "CRITICA_N_FONTES",
    "CITA_N_PAPERS", "CITA_SOMA_CIT", "PRESTIGIO_FESTIV",
    "CNPJ_produtora",
] + _OPT_DIV

_inv_rows = []

for _, obra in df.iterrows():
    _shared = {c: obra.get(c, "") for c in _BASE_COLS}
    _fsa   = _num(obra.get("Valor FSA (R$)", 0))
    _fsa_d = _num(obra.get("Valor FSA Deflac. (R$2024)", 0))
    _r3    = _num(obra.get("Renúncia Art.3/3-A/39 (R$)", 0))
    _rout  = _num(obra.get("Renúncia Outros Mec. (R$)", 0))
    _r_tot_d = _num(obra.get("Renúncia Total Deflac. (R$2024)", 0))
    _ano   = _int(obra.get("Ano", 0))
    _fat   = IPCA_FATORES.get(_ano, 1.0) if IPCA_FATORES else 1.0
    _rec_d = _num(obra.get("_receita_deflac", 0))

    def _roi_sobre(inv_d):
        return round(_rec_d / inv_d, 4) if inv_d > 0 else None

    if _fsa > 0:
        _inv_rows.append({**_shared,
            "tipo_investimento":         "FSA",
            "valor_nominal_r":           round(_fsa, 2),
            "valor_deflac_r2024":        round(_fsa_d, 2),
            "ano_referencia_deflacao":   _ano,
            "roi_sobre_este_inv_deflac": _roi_sobre(_fsa_d),
        })
    if _r3 > 0:
        _r3_d = round(_r3 * _fat, 2)
        _inv_rows.append({**_shared,
            "tipo_investimento":         "RENUNCIA_ART3",
            "valor_nominal_r":           round(_r3, 2),
            "valor_deflac_r2024":        _r3_d,
            "ano_referencia_deflacao":   _ano,
            "roi_sobre_este_inv_deflac": _roi_sobre(_r3_d),
        })
    if _rout > 0:
        _ro_d = round(_rout * _fat, 2)
        _inv_rows.append({**_shared,
            "tipo_investimento":         "RENUNCIA_OUTROS",
            "valor_nominal_r":           round(_rout, 2),
            "valor_deflac_r2024":        _ro_d,
            "ano_referencia_deflacao":   _ano,
            "roi_sobre_este_inv_deflac": _roi_sobre(_ro_d),
        })

df_inv = pd.DataFrame(_inv_rows).rename(columns={
    "Projeto": "titulo", "Bilheteria Nominal (R$)": "bilheteria_nominal_r",
    "Bilheteria Deflac. (R$)": "bilheteria_deflac_r2024",
    "Outras Janelas Deflac. (R$2024)": "outras_janelas_deflac_r2024",
    "Outras Janelas Nominal (R$)": "outras_janelas_nominal_r",
    "_receita_deflac": "receita_total_deflac_r2024",
    "ROI Dom. FSA (deflac)": "roi_dom_fsa_deflac",
    "ROI Dom. Total (deflac)": "roi_dom_total_deflac",
    "ROI Internacional (0-100)": "roi_internacional_0_100",
    "Total Países Alcançados": "total_paises_alcancados",
    "Pontuação Festivais": "pontuacao_festivais",
    "Adm. EU — Lumière": "adm_eu_lumiere",
    "VOD Intl — N Plataformas": "vod_n_plataformas",
    "VOD Intl — N Países": "vod_n_paises",
    "CRITICA_INDICE_1_5": "critica_indice_1_5",
    "CRITICA_N_FONTES": "critica_n_fontes",
    "CITA_N_PAPERS": "cita_n_papers",
    "CITA_SOMA_CIT": "cita_soma_cit",
    "PRESTIGIO_FESTIV": "prestigio_festiv",
    "GENERO_DIRETOR": "genero_diretor",
    "DIVERSIDADE_REGIONAL": "diversidade_regional",
    "PROTAGONISTA_DIVERSO": "protagonista_diverso",
    "UF_REQUERENTE": "UF_requerente",
    "TEM_DIRETOR_PRESTIGIO": "tem_diretor_prestigio",
    "Ano": "ano", "Chamada": "chamada", "Categoria": "categoria",
})

_path_a = DIR_OUT / "base_nivel_investimento.csv"
df_inv.to_csv(_path_a, sep=";", index=False, encoding=ENC)
print(f"  {len(df_inv)} linhas -> {_path_a.name}")


# ═══════════════════════════════════════════════════════════════════════════════
# DATASET B — BASE NÍVEL PRODUTORA (só independentes)
# ═══════════════════════════════════════════════════════════════════════════════
print("\nGerando base_nivel_produtora.csv...")

BILH_THRESHOLD = 2_500_000   # alinhado com painel_produtoras (script 04)
INV_THRESHOLD  = 5_000_000
REC_NEW_TH     = 10_000_000  # limiar superior de receita (Retorno Doméstico)
ROI_DOM_TH     = 0.6         # ROI doméstico mínimo para Retorno Doméstico via ratio
REC_PISO       = 2_500_000   # piso de receita para critério de ROI Dom (alinhado com Duplo Retorno)

def _classificar_cluster(rec_deflac, roi_intl_max, roi_intl_med, inv_total_deflac, roi_dom_total=0.0):
    # Classificação de produtoras por padrão de retorno
    # Duplo Retorno: cauda superior da distribuição (roi_intl_max >= 13, acima do cluster denso 10–13)
    duplo_intl = roi_intl_max >= 13

    if rec_deflac >= BILH_THRESHOLD and duplo_intl:
        return "Duplo Retorno"

    # Retorno Doméstico: renda >= R$10M  OU  (ROI Dom > 0,6 E renda >= R$2,5M)
    is_dom = (rec_deflac >= REC_NEW_TH) or (roi_dom_total > ROI_DOM_TH and rec_deflac >= REC_PISO)
    if is_dom and not duplo_intl:
        return "Retorno Doméstico"

    if duplo_intl:
        return "Retorno Internacional"

    # Fomento Baixo Retorno: Investimento > R$ 5M E sem retorno internacional qualificado
    if inv_total_deflac > INV_THRESHOLD:
        return "Fomento Baixo Retorno"

    return "Pequeno Porte"

# Filtrar TV (conteúdo, não cinema) antes de processar produtoras
# ESCOPO: Incluir PRODECINE + Renúncia Fiscal + SAV/MINC, EXCLUIR TV
df_prod_scope = df[~df["Categoria"].str.contains("_tv_excluir", na=False)].copy()
print(f"\n[ESCOPO PRODUTORA] Filtrando TV: {len(df)} obras antes, {len(df_prod_scope)} depois")
print(f"  Categorias incluídas: PRODECINE + Renúncia Fiscal + SAV/MINC")

df_prod_rows = []
for cnpj, grp in df_prod_scope.groupby("CNPJ_produtora"):
    if not cnpj:
        continue
    cnpj_clean = _cnpj_clean(cnpj)
    if cnpjs_indep and cnpj_clean not in cnpjs_indep:
        continue

    ag = agentes_map.get(cnpj_clean, {})

    n_obras      = len(grp)
    n_fsa        = (grp["Valor FSA (R$)"].astype(float) > 0).sum()
    n_ren        = ((grp["Renúncia Art.3/3-A/39 (R$)"].astype(float) +
                     grp["Renúncia Outros Mec. (R$)"].astype(float)) > 0).sum()
    inv_fsa_d    = grp["Valor FSA Deflac. (R$2024)"].astype(float).sum()
    inv_tot_d    = grp["Investimento Total Deflac. (R$2024)"].astype(float).sum()
    bil_d        = grp["Bilheteria Deflac. (R$)"].astype(float).sum()
    jan_d        = grp["Outras Janelas Deflac. (R$2024)"].astype(float).sum()
    rec_d        = bil_d + jan_d  # Receita = Bilheteria + Outras Janelas
    roi_fsa_d    = round(rec_d / inv_fsa_d,  4) if inv_fsa_d  > 0 else None
    roi_tot_d    = round(rec_d / inv_tot_d,  4) if inv_tot_d  > 0 else None
    roi_intl_med = grp["ROI Internacional (0-100)"].astype(float).mean()
    roi_intl_max = grp["ROI Internacional (0-100)"].astype(float).max()
    n_festival   = (grp["Pontuação Festivais"].astype(float) > 0).sum()
    n_roi_intl_qualificado = (grp["ROI Internacional (0-100)"].astype(float) >= 13).sum()
    n_lumiere    = (grp["Adm. EU — Lumière"].astype(float) > 0).sum()
    n_vod        = (grp["VOD Intl — N Plataformas"].astype(float) > 0).sum()
    fest_max     = grp["Pontuação Festivais"].astype(float).max()
    adm_eu_tot   = grp["Adm. EU — Lumière"].astype(float).sum()
    critica_med  = grp["CRITICA_INDICE_1_5"].astype(float).replace(0, np.nan).mean()
    paises_max   = grp["Total Países Alcançados"].astype(float).max()
    paises_obras = (grp["Total Países Alcançados"].astype(float) > 0).sum()
    # citações máximas da produtora
    cita_max     = grp["CITA_N_PAPERS"].replace("", np.nan).astype(float).max()
    # gênero
    pct_fem = (grp["GENERO_DIRETOR"].str.upper().str.startswith("F")).mean() * 100

    cluster = _classificar_cluster(
        rec_d,
        roi_intl_max if pd.notna(roi_intl_max) else 0,
        roi_intl_med if pd.notna(roi_intl_med) else 0,
        inv_tot_d,
        roi_tot_d if roi_tot_d is not None else 0.0,
    )

    df_prod_rows.append({
        "CNPJ_produtora":              cnpj_clean,
        "razao_social":                ag.get("razao_social", ""),
        "UF":                          ag.get("UF", ""),
        "classificacao_agente":        ag.get("classificacao", ""),
        "n_obras":                     n_obras,
        "n_obras_fsa":                 int(n_fsa),
        "n_obras_renuncia":            int(n_ren),
        "investimento_fsa_deflac":     round(inv_fsa_d, 2),
        "investimento_total_deflac":   round(inv_tot_d, 2),
        "bilheteria_deflac":           round(bil_d, 2),
        "outras_janelas_deflac":       round(jan_d, 2),
        "receita_total_deflac":        round(rec_d, 2),
        "roi_dom_fsa_deflac":          roi_fsa_d,
        "roi_dom_total_deflac":        roi_tot_d,
        "roi_intl_medio":              round(roi_intl_med, 2) if pd.notna(roi_intl_med) else None,
        "roi_intl_max":                round(roi_intl_max, 2) if pd.notna(roi_intl_max) else None,
        "total_paises_max":            int(paises_max) if pd.notna(paises_max) else 0,
        "n_obras_com_presenca_intl":   int(n_roi_intl_qualificado),
        "n_obras_com_festival":        int(n_festival),
        "n_obras_com_lumiere":         int(n_lumiere),
        "n_obras_com_vod":             int(n_vod),
        "pontuacao_festivais_max":     int(fest_max) if pd.notna(fest_max) else 0,
        "adm_eu_total":                int(adm_eu_tot),
        "critica_media":               round(critica_med, 3) if pd.notna(critica_med) else None,
        "cita_n_papers_max_diretor":   int(cita_max) if pd.notna(cita_max) else None,
        "pct_obras_genero_feminino":   round(pct_fem, 1),
        "cluster":                     cluster,
    })

df_prod = pd.DataFrame(df_prod_rows)
_path_b = DIR_OUT / "base_nivel_produtora.csv"
df_prod.to_csv(_path_b, sep=";", index=False, encoding=ENC)
print(f"  {len(df_prod)} produtoras (2014-2023) -> {_path_b.name}")
if len(df_prod) > 0:
    print("  Distribuição de clusters:")
    for cl, cnt in df_prod["cluster"].value_counts().items():
        print(f"    {cl}: {cnt}")

# ─── GERAR DATASET AMPLO (todo período, todo espectro de renúncias) ────
print("\n[DATASET AMPLO] Gerando base_nivel_produtora_amplo.csv...")
print(f"  Usando df_amplo com {len(df_amplo)} obras (todo histórico)")

# Preparar df_amplo com mesmo processamento que df
for col in _NUM_COLS:
    if col in df_amplo.columns:
        df_amplo[col] = pd.to_numeric(df_amplo[col], errors="coerce")
    else:
        df_amplo[col] = 0.0

# Adicionar colunas calculadas (como feito em df)
if "CPB" not in df_amplo.columns:
    df_amplo["CPB"] = ""
if "Categoria" not in df_amplo.columns:
    df_amplo["Categoria"] = ""

# Aplicar mesmos processamentos (categorias, renúncia, etc) - já vem da tabela consolidada
# Apenas filtrar bilheteria > 0
df_amplo_cinema = df_amplo[df_amplo["Bilheteria Deflac. (R$)"].astype(float) > 0].copy()
print(f"  Após filtro bilheteria > 0: {len(df_amplo_cinema)} obras")

# Aplicar mesmo filtro de TV
df_prod_scope_amplo = df_amplo_cinema[~df_amplo_cinema["Categoria"].str.contains("_tv_excluir", na=False)].copy()
print(f"  Após filtro TV: {len(df_prod_scope_amplo)} obras")

# Garantir que CNPJ_produtora existe
if "CNPJ_produtora" not in df_prod_scope_amplo.columns:
    df_prod_scope_amplo["CNPJ_produtora"] = ""
    print("  [AVISO] Adicionando coluna CNPJ_produtora vazia")

# Agregação de produtoras (mesmo código, diferentes dados)
df_prod_rows_amplo = []
for cnpj, grp in df_prod_scope_amplo.groupby("CNPJ_produtora"):
    if not cnpj:
        continue
    cnpj_clean = _cnpj_clean(cnpj)
    if cnpjs_indep and cnpj_clean not in cnpjs_indep:
        continue

    ag = agentes_map.get(cnpj_clean, {})

    n_obras      = len(grp)
    n_fsa        = (grp["Valor FSA (R$)"].astype(float) > 0).sum()
    n_ren        = ((grp["Renúncia Art.3/3-A/39 (R$)"].astype(float) +
                     grp["Renúncia Outros Mec. (R$)"].astype(float)) > 0).sum()
    inv_fsa_d    = grp["Valor FSA Deflac. (R$2024)"].astype(float).sum()
    inv_tot_d    = grp["Investimento Total Deflac. (R$2024)"].astype(float).sum()
    bil_d        = grp["Bilheteria Deflac. (R$)"].astype(float).sum()
    jan_d        = grp["Outras Janelas Deflac. (R$2024)"].astype(float).sum()
    rec_d        = bil_d + jan_d
    roi_fsa_d    = round(rec_d / inv_fsa_d,  4) if inv_fsa_d  > 0 else None
    roi_tot_d    = round(rec_d / inv_tot_d,  4) if inv_tot_d  > 0 else None
    roi_intl_med = grp["ROI Internacional (0-100)"].astype(float).mean()
    roi_intl_max = grp["ROI Internacional (0-100)"].astype(float).max()
    n_festival   = (grp["Pontuação Festivais"].astype(float) > 0).sum()
    n_roi_intl_qualificado = (grp["ROI Internacional (0-100)"].astype(float) >= 13).sum()
    n_lumiere    = (grp["Adm. EU — Lumière"].astype(float) > 0).sum()
    n_vod        = (grp["VOD Intl — N Plataformas"].astype(float) > 0).sum()
    fest_max     = grp["Pontuação Festivais"].astype(float).max()
    adm_eu_tot   = grp["Adm. EU — Lumière"].astype(float).sum()
    critica_med  = grp["CRITICA_INDICE_1_5"].astype(float).replace(0, np.nan).mean()
    paises_max   = grp["Total Países Alcançados"].astype(float).max()
    paises_obras = (grp["Total Países Alcançados"].astype(float) > 0).sum()
    cita_max     = grp["CITA_N_PAPERS"].replace("", np.nan).astype(float).max()
    pct_fem = (grp["GENERO_DIRETOR"].str.upper().str.startswith("F")).mean() * 100

    cluster = _classificar_cluster(
        rec_d,
        roi_intl_max if pd.notna(roi_intl_max) else 0,
        roi_intl_med if pd.notna(roi_intl_med) else 0,
        inv_tot_d,
        roi_tot_d if roi_tot_d is not None else 0.0,
    )

    df_prod_rows_amplo.append({
        "CNPJ_produtora":              cnpj_clean,
        "razao_social":                ag.get("razao_social", ""),
        "UF":                          ag.get("UF", ""),
        "classificacao_agente":        ag.get("classificacao", ""),
        "n_obras":                     n_obras,
        "n_obras_fsa":                 int(n_fsa),
        "n_obras_renuncia":            int(n_ren),
        "investimento_fsa_deflac":     round(inv_fsa_d, 2),
        "investimento_total_deflac":   round(inv_tot_d, 2),
        "bilheteria_deflac":           round(bil_d, 2),
        "outras_janelas_deflac":       round(jan_d, 2),
        "receita_total_deflac":        round(rec_d, 2),
        "roi_dom_fsa_deflac":          roi_fsa_d,
        "roi_dom_total_deflac":        roi_tot_d,
        "roi_intl_medio":              round(roi_intl_med, 2) if pd.notna(roi_intl_med) else None,
        "roi_intl_max":                round(roi_intl_max, 2) if pd.notna(roi_intl_max) else None,
        "total_paises_max":            int(paises_max) if pd.notna(paises_max) else 0,
        "n_obras_com_presenca_intl":   int(n_roi_intl_qualificado),
        "n_obras_com_festival":        int(n_festival),
        "n_obras_com_lumiere":         int(n_lumiere),
        "n_obras_com_vod":             int(n_vod),
        "pontuacao_festivais_max":     int(fest_max) if pd.notna(fest_max) else 0,
        "adm_eu_total":                int(adm_eu_tot),
        "critica_media":               round(critica_med, 3) if pd.notna(critica_med) else None,
        "cita_n_papers_max_diretor":   int(cita_max) if pd.notna(cita_max) else None,
        "pct_obras_genero_feminino":   round(pct_fem, 1),
        "cluster":                     cluster,
    })

df_prod_amplo = pd.DataFrame(df_prod_rows_amplo)
_path_b_amplo = DIR_OUT / "base_nivel_produtora_amplo.csv"
df_prod_amplo.to_csv(_path_b_amplo, sep=";", index=False, encoding=ENC)
print(f"  {len(df_prod_amplo)} produtoras (todo período) -> {_path_b_amplo.name}")
if len(df_prod_amplo) > 0:
    print("  Distribuição de clusters (amplo):")
    for cl, cnt in df_prod_amplo["cluster"].value_counts().items():
        print(f"    {cl}: {cnt}")


# ═══════════════════════════════════════════════════════════════════════════════
# DATASET C — BASE NÍVEL CHAMADA
# ═══════════════════════════════════════════════════════════════════════════════
print("\nGerando base_nivel_chamada.csv...")

def _pct_str_col(series, val):
    if series.empty:
        return 0.0
    return round((series.str.upper().str.startswith(val.upper())).mean() * 100, 1)

def _distinct_paises(series):
    paises = set()
    for v in series:
        if v:
            for p in str(v).split("|"):
                p = p.strip()
                if p:
                    paises.add(p)
    return len(paises)

# Filtrar CHAMADA: PRODECINE APENAS (FSA direto)
# ESCOPO: Apenas obras PRODECINE (sem SAV/MINC, sem Renúncia, sem TV)
df_chamada_scope = df[df["Categoria"].str.contains("FSA", na=False, case=False)].copy()
df_chamada_scope = df_chamada_scope[~df_chamada_scope["Categoria"].str.contains("_tv_excluir|SAV", na=False)].copy()
print(f"\n[ESCOPO CHAMADA] PRODECINE apenas: {len(df_chamada_scope)} obras")
print(f"  Excluído: TV, SAV/MINC, Renúncia Fiscal")

_cat_rows = []
for _ch, _grp in df_chamada_scope.groupby("Chamada"):
    _cat_rows.append((_ch or "(sem chamada)", _grp))

for chamada, grp in _cat_rows:
    if "_tv_excluir" in str(grp["Categoria"].iloc[0] if len(grp) > 0 else ""):
        continue

    cat        = grp["Categoria"].iloc[0] if len(grp) > 0 else ""
    n          = len(grp)
    inv_fsa_d  = grp["Valor FSA Deflac. (R$2024)"].astype(float).sum()
    inv_tot_d  = grp["Investimento Total Deflac. (R$2024)"].astype(float).sum()
    bil_d      = grp["Bilheteria Deflac. (R$)"].astype(float).sum()
    jan_d      = grp["Outras Janelas Deflac. (R$2024)"].astype(float).sum()
    rec_d      = bil_d + jan_d  # Receita total
    n_bil      = (grp["Bilheteria Deflac. (R$)"].astype(float) > 0).sum()
    n_fest     = (grp["Pontuação Festivais"].astype(float) > 0).sum()
    n_lum      = (grp["Adm. EU — Lumière"].astype(float) > 0).sum()
    n_vod      = (grp["VOD Intl — N Plataformas"].astype(float) > 0).sum()
    adm_eu_tot = grp["Adm. EU — Lumière"].astype(float).sum()
    vod_max    = grp["VOD Intl — N Plataformas"].astype(float).max()
    roi_fsa_m  = round(rec_d / inv_fsa_d,  4) if inv_fsa_d  > 0 else None
    roi_tot_m  = round(rec_d / inv_tot_d,  4) if inv_tot_d  > 0 else None
    _rois_fsa  = grp["ROI Dom. FSA (deflac)"].replace("", np.nan).astype(float)
    _rois_tot  = grp["ROI Dom. Total (deflac)"].replace("", np.nan).astype(float)
    roi_fsa_med_obra  = round(_rois_fsa.mean(), 4) if _rois_fsa.notna().any() else None
    roi_tot_med_obra  = round(_rois_tot.mean(), 4) if _rois_tot.notna().any() else None
    roi_intl_med = grp["ROI Internacional (0-100)"].astype(float).mean()
    roi_intl_max = grp["ROI Internacional (0-100)"].astype(float).max()
    fest_med   = grp["Pontuação Festivais"].astype(float).mean()
    fest_max   = grp["Pontuação Festivais"].astype(float).max()
    critica_m  = grp["CRITICA_INDICE_1_5"].astype(float).replace(0, np.nan).mean()
    cita_med   = grp["CITA_N_PAPERS"].replace("", np.nan).astype(float).mean()
    n_prod     = grp["CNPJ_produtora"].nunique()
    n_paises   = _distinct_paises(grp.get("Países Lista", pd.Series(dtype=str)))
    _g_fem     = grp.get("GENERO_DIRETOR", pd.Series(dtype=str))
    _g_reg     = grp.get("DIVERSIDADE_REGIONAL", pd.Series(dtype=str))
    pct_fem    = _pct_str_col(_g_fem, "F")
    pct_reg    = round((_g_reg.astype(str).str.strip().isin(["1", "True", "true"])).mean() * 100, 1)

    _chamada_row = {
        "chamada":                   chamada,
        "categoria":                 cat,
        "n_obras":                   n,
        "n_obras_com_bilheteria":    int(n_bil),
        "n_obras_com_festival":      int(n_fest),
        "n_obras_com_lumiere":       int(n_lum),
        "n_obras_com_vod":           int(n_vod),
        "investimento_fsa_deflac":   round(inv_fsa_d, 0),
        "investimento_total_deflac": round(inv_tot_d, 0),
        "bilheteria_total_deflac":   round(bil_d, 0),
        "outras_janelas_deflac":     round(jan_d, 0),
        "receita_total_deflac":      round(rec_d, 0),
        "roi_dom_fsa_agregado":      roi_fsa_m,
        "roi_dom_total_agregado":    roi_tot_m,
        "roi_dom_fsa_medio_obra":    roi_fsa_med_obra,
        "roi_dom_total_medio_obra":  roi_tot_med_obra,
        "roi_intl_medio":            round(roi_intl_med, 2) if pd.notna(roi_intl_med) else None,
        "roi_intl_max":              round(roi_intl_max, 2) if pd.notna(roi_intl_max) else None,
        "pct_com_festival":          round(n_fest / n * 100, 1) if n > 0 else 0,
        "pct_com_lumiere":           round(n_lum  / n * 100, 1) if n > 0 else 0,
        "pct_com_vod":               round(n_vod  / n * 100, 1) if n > 0 else 0,
        "total_paises_distintos":    n_paises,
        "adm_eu_total":              int(adm_eu_tot),
        "vod_plataformas_max":       int(vod_max) if pd.notna(vod_max) else 0,
        "pontuacao_festivais_media": round(fest_med, 1) if pd.notna(fest_med) else None,
        "pontuacao_festivais_max":   int(fest_max) if pd.notna(fest_max) else 0,
        "critica_media":             round(critica_m, 3) if pd.notna(critica_m) else None,
        "cita_n_papers_media":       round(cita_med, 2) if pd.notna(cita_med) else None,
        "pct_genero_feminino":       pct_fem,
        "pct_diversidade_regional":  pct_reg,
        "n_produtoras_distintas":    int(n_prod),
    }
    _cat_rows[_cat_rows.index((chamada, grp))] = _chamada_row  # replace tuple with dict

# Filtra só os dicts (removendo as tuples não processadas)
df_ch = pd.DataFrame([r for r in _cat_rows if isinstance(r, dict)]) \
          .sort_values("receita_total_deflac", ascending=False)
_path_c = DIR_OUT / "base_nivel_chamada.csv"
df_ch.to_csv(_path_c, sep=";", index=False, encoding=ENC)
print(f"  {len(df_ch)} chamadas -> {_path_c.name}")


# ═══════════════════════════════════════════════════════════════════════════════
# DATASET D — BASE NÍVEL OBRA (investimentos como colunas)
# ═══════════════════════════════════════════════════════════════════════════════
print("\nGerando base_nivel_obra.csv...")

_obra_desired = [
    "CPB", "Projeto", "Ano", "Chamada", "Categoria",
    "Valor FSA (R$)", "Renúncia Art.3/3-A/39 (R$)", "Renúncia Outros Mec. (R$)",
    "Valor FSA Deflac. (R$2024)", "Renúncia Total Deflac. (R$2024)",
    "Investimento Total Deflac. (R$2024)",
    "Bilheteria Nominal (R$)", "Bilheteria Deflac. (R$)",
    "Outras Janelas Nominal (R$)", "Outras Janelas Deflac. (R$2024)",
    "_receita_deflac",
    "ROI Dom. FSA (deflac)", "ROI Dom. Total (deflac)",
    "ROI Internacional (0-100)",
    "Pontuação Festivais", "Adm. EU — Lumière",
    "VOD Intl — N Plataformas", "VOD Intl — N Países",
    "Total Países Alcançados", "Países Festivais", "Países Lumière",
    "Países VOD Europa", "Países Lista",
    "CRITICA_INDICE_1_5", "CRITICA_FONTES", "CRITICA_N_FONTES", "CRITICA_CONFIANCA",
    "CITA_N_PAPERS", "CITA_SOMA_CIT", "CITA_MAX_CIT", "CITA_VENUES",
    "PRESTIGIO_FESTIV", "WIKIDATA_ID_DIR",
    "CNPJ_produtora",
    "GENERO_DIRETOR", "DIVERSIDADE_REGIONAL", "PROTAGONISTA_DIVERSO",
    "UF_REQUERENTE", "TEM_DIRETOR_PRESTIGIO",
    "IMDB_GENRES", "IMDB_RUNTIME",
    "PRESTIGIO_FESTIVAIS", "N_DIRETORES",
]
_obra_cols = [c for c in _obra_desired if c in df.columns]
df_obra = df[_obra_cols].copy()

df_obra["investimento_total_nominal"] = (
    df_obra["Valor FSA (R$)"].astype(float) +
    df_obra["Renúncia Art.3/3-A/39 (R$)"].astype(float) +
    df_obra["Renúncia Outros Mec. (R$)"].astype(float)
)

_inv_nom = df_obra["investimento_total_nominal"]
_bil_nom = df_obra["Bilheteria Nominal (R$)"].astype(float)
_jan_nom = df_obra["Outras Janelas Nominal (R$)"].astype(float)
_fsa_nom = df_obra["Valor FSA (R$)"].astype(float)
_rec_nom = _bil_nom + _jan_nom
df_obra["roi_dom_fsa_nominal"]   = (_rec_nom / _fsa_nom).where(_fsa_nom > 0).round(4)
df_obra["roi_dom_total_nominal"] = (_rec_nom / _inv_nom).where(_inv_nom > 0).round(4)

df_obra.rename(columns={
    "Projeto": "titulo", "Ano": "ano", "Chamada": "chamada", "Categoria": "categoria",
    "Valor FSA (R$)": "investimento_fsa_nominal",
    "Renúncia Art.3/3-A/39 (R$)": "investimento_renuncia_art3_nominal",
    "Renúncia Outros Mec. (R$)": "investimento_renuncia_outros_nominal",
    "Valor FSA Deflac. (R$2024)": "investimento_fsa_deflac",
    "Renúncia Total Deflac. (R$2024)": "investimento_renuncia_total_deflac",
    "Investimento Total Deflac. (R$2024)": "investimento_total_deflac",
    "Bilheteria Nominal (R$)": "bilheteria_nominal",
    "Bilheteria Deflac. (R$)": "bilheteria_deflac",
    "Outras Janelas Nominal (R$)": "outras_janelas_nominal",
    "Outras Janelas Deflac. (R$2024)": "outras_janelas_deflac",
    "_receita_deflac": "receita_total_deflac",
    "ROI Dom. FSA (deflac)": "roi_dom_fsa_deflac",
    "ROI Dom. Total (deflac)": "roi_dom_total_deflac",
    "ROI Internacional (0-100)": "roi_internacional_0_100",
    "Pontuação Festivais": "pontuacao_festivais",
    "Adm. EU — Lumière": "adm_eu_lumiere",
    "VOD Intl — N Plataformas": "vod_n_plataformas",
    "VOD Intl — N Países": "vod_n_paises",
    "Total Países Alcançados": "total_paises_alcancados",
    "Países Festivais": "paises_festivais",
    "Países Lumière": "paises_lumiere",
    "Países VOD Europa": "paises_vod_europa",
    "Países Lista": "paises_lista",
    "CRITICA_INDICE_1_5": "critica_indice_1_5",
    "CRITICA_FONTES": "critica_fontes",
    "CRITICA_N_FONTES": "critica_n_fontes",
    "CRITICA_CONFIANCA": "critica_confianca",
    "CITA_N_PAPERS": "cita_n_papers",
    "CITA_SOMA_CIT": "cita_soma_cit",
    "CITA_MAX_CIT": "cita_max_cit",
    "CITA_VENUES": "cita_venues",
    "PRESTIGIO_FESTIV": "prestigio_festiv",
    "WIKIDATA_ID_DIR": "wikidata_id_dir",
    "GENERO_DIRETOR": "genero_diretor",
    "DIVERSIDADE_REGIONAL": "diversidade_regional",
    "PROTAGONISTA_DIVERSO": "protagonista_diverso",
    "UF_REQUERENTE": "UF_requerente",
    "TEM_DIRETOR_PRESTIGIO": "tem_diretor_prestigio",
    "IMDB_GENRES": "imdb_genres",
    "IMDB_RUNTIME": "imdb_runtime",
    "PRESTIGIO_FESTIVAIS": "prestigio_festivais",
    "CITA_N_PAPERS": "cita_n_papers",
    "N_DIRETORES": "n_diretores",
}, inplace=True)

_path_d = DIR_OUT / "base_nivel_obra.csv"
df_obra.to_csv(_path_d, sep=";", index=False, encoding=ENC)
print(f"  {len(df_obra)} obras -> {_path_d.name}")


# ═══════════════════════════════════════════════════════════════════════════════
# DATASET E — BASE NÍVEL DIREÇÃO
# ═══════════════════════════════════════════════════════════════════════════════
print("\nGerando base_nivel_direcao.csv...")

# Perfil de festivais por diretor
_perfil_festiv_map = {}   # diretor_norm → dict
try:
    df_pf = pd.read_csv(PERFIL_FESTIV_CSV, sep=None, engine="python", encoding=ENC, dtype=str).fillna("")
    _pf_cols = [c for c in df_pf.columns if c != "DIRETOR"]
    for _, row in df_pf.iterrows():
        nome = _norm(row.get("DIRETOR", ""))
        if nome:
            _perfil_festiv_map[nome] = {c: row.get(c, "") for c in _pf_cols}
    print(f"  Perfis de festivais: {len(_perfil_festiv_map)}")
except Exception as e:
    print(f"  [AVISO] perfil_festivais: {e}")

# Construir df de diretores a partir do CSV de diretores ANCINE
try:
    df_dir_raw = pd.read_csv(DIRETORES_CSV, sep=None, engine="python", encoding=ENC, dtype=str).fillna("")
    # filtra só brasileiros e une com df (master) por CPB
    df_dir_br = df_dir_raw[df_dir_raw["PAIS_DIRETOR"].str.upper() == "BRASIL"].copy()
    df_dir_br["DIRETOR_NORM"] = df_dir_br["DIRETOR"].apply(_norm)

    # join com master
    df_master_slim = df[["CPB", "Projeto", "Ano", "Chamada", "Categoria",
                          "Valor FSA (R$)", "Renúncia Art.3/3-A/39 (R$)", "Renúncia Outros Mec. (R$)",
                          "Valor FSA Deflac. (R$2024)", "Renúncia Total Deflac. (R$2024)",
                          "Investimento Total Deflac. (R$2024)",
                          "Bilheteria Deflac. (R$)", "Outras Janelas Deflac. (R$2024)",
                          "_receita_deflac",
                          "ROI Dom. FSA (deflac)", "ROI Internacional (0-100)",
                          "Pontuação Festivais", "Adm. EU — Lumière",
                          "CRITICA_INDICE_1_5",
                          "GENERO_DIRETOR", "DIVERSIDADE_REGIONAL",
                          "CITA_N_PAPERS", "CITA_SOMA_CIT",
                          ]].copy()

    df_dir_joined = df_dir_br.merge(df_master_slim, on="CPB", how="left")

    # Agregação por diretor
    _dir_rows = []
    for dname, grp in df_dir_joined.groupby("DIRETOR_NORM"):
        if not dname:
            continue
        nome_orig = grp["DIRETOR"].iloc[0]
        n         = grp["CPB"].nunique()
        anos      = grp["Ano"].replace("", np.nan).dropna()
        ano_1     = anos.astype(str).min() if len(anos) > 0 else ""
        ano_ult   = anos.astype(str).max() if len(anos) > 0 else ""
        n_fsa     = (grp["Valor FSA (R$)"].astype(float) > 0).sum()
        n_ren     = ((grp["Renúncia Art.3/3-A/39 (R$)"].astype(float) +
                      grp["Renúncia Outros Mec. (R$)"].astype(float)) > 0).sum()
        inv_fsa_d = grp["Valor FSA Deflac. (R$2024)"].astype(float).sum()
        inv_tot_d = grp["Investimento Total Deflac. (R$2024)"].astype(float).sum()
        bil_d     = grp["Bilheteria Deflac. (R$)"].astype(float).sum()
        jan_d     = grp["Outras Janelas Deflac. (R$2024)"].astype(float).sum()
        rec_d     = bil_d + jan_d  # Receita total
        roi_fsa   = round(rec_d / inv_fsa_d, 4) if inv_fsa_d > 0 else None
        roi_tot   = round(rec_d / inv_tot_d, 4) if inv_tot_d > 0 else None
        roi_intl  = grp["ROI Internacional (0-100)"].astype(float).mean()
        fest_max  = grp["Pontuação Festivais"].astype(float).max()
        fest_med  = grp["Pontuação Festivais"].astype(float).mean()
        adm_eu    = grp["Adm. EU — Lumière"].astype(float).sum()
        critica_m = grp["CRITICA_INDICE_1_5"].astype(float).replace(0, np.nan).mean()
        genero    = grp["GENERO_DIRETOR"].replace("", np.nan).dropna()
        genero_v  = genero.iloc[0] if len(genero) > 0 else ""
        div_reg   = grp["DIVERSIDADE_REGIONAL"].replace("", np.nan).dropna()
        div_v     = div_reg.iloc[0] if len(div_reg) > 0 else ""
        # categorias FSA
        cats      = grp["Categoria"].value_counts()
        cat_princ = cats.index[0] if len(cats) > 0 else ""

        # Citações (pega o max registrado nas obras do diretor)
        cit_n     = grp["CITA_N_PAPERS"].replace("", np.nan).astype(float).max()
        cit_soma  = grp["CITA_SOMA_CIT"].replace("", np.nan).astype(float).max()

        # Sobrescreve com dado direto da tabela de citações (mais preciso)
        rec_cit = _cit_map.get(dname)
        if rec_cit:
            cit_n    = _int(rec_cit["cita_n_papers"])
            cit_soma = _int(rec_cit["cita_soma_cit"])

        # Prestígio wikidata
        rec_prest = _prest_map.get(dname, {})
        wikidata  = rec_prest.get("wikidata_id", "")
        prest_fv  = rec_prest.get("prestigio_festiv", "")
        prest_fm  = rec_prest.get("prestigio_filmes", "")
        if not prest_fv and rec_cit:
            prest_fv = rec_cit.get("prestigio_festiv_cit", "")

        # Perfil de festivais
        perf = _perfil_festiv_map.get(dname, {})

        row_dir = {
            "diretor":                 nome_orig,
            "diretor_norm":            dname,
            "n_obras_fomento":         n,
            "ano_primeira_obra":       ano_1,
            "ano_ultima_obra":         ano_ult,
            "n_obras_fsa":             int(n_fsa),
            "n_obras_renuncia":        int(n_ren),
            "investimento_fsa_deflac": round(inv_fsa_d, 2),
            "investimento_total_deflac": round(inv_tot_d, 2),
            "bilheteria_deflac":       round(bil_d, 2),
            "outras_janelas_deflac":   round(jan_d, 2),
            "receita_total_deflac":    round(rec_d, 2),
            "roi_dom_fsa_deflac":      roi_fsa,
            "roi_dom_total_deflac":    roi_tot,
            "roi_intl_medio":          round(roi_intl, 2) if pd.notna(roi_intl) else None,
            "pontuacao_festivais_max": round(fest_max, 1) if pd.notna(fest_max) else None,
            "pontuacao_festivais_media": round(fest_med, 2) if pd.notna(fest_med) else None,
            "adm_eu_total":            int(adm_eu),
            "critica_media":           round(critica_m, 3) if pd.notna(critica_m) else None,
            "categoria_principal":     cat_princ,
            # Diversidade
            "genero_diretor":          genero_v,
            "diversidade_regional":    div_v,
            # Citações acadêmicas
            "cita_n_papers":           int(cit_n) if pd.notna(cit_n) else None,
            "cita_soma_citacoes":      int(cit_soma) if pd.notna(cit_soma) else None,
            # Prestígio
            "wikidata_id":             wikidata,
            "prestigio_festiv":        prest_fv,
            "prestigio_filmes":        prest_fm,
        }
        # Colunas do perfil de festivais
        for col, val in perf.items():
            row_dir[f"fest_{col.lower()}"] = val

        _dir_rows.append(row_dir)

    df_dir_out = pd.DataFrame(_dir_rows).sort_values("n_obras_fomento", ascending=False)
    _path_e = DIR_OUT / "base_nivel_direcao.csv"
    df_dir_out.to_csv(_path_e, sep=";", index=False, encoding=ENC)
    print(f"  {len(df_dir_out)} diretores -> {_path_e.name}")

except Exception as e:
    print(f"  [ERRO] base_nivel_direcao: {e}")
    import traceback; traceback.print_exc()


# ── Resumo final ───────────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("DATASETS GERADOS EM: resultados/datasets/")
print("=" * 60)
print(f"  a) {_path_a.name:<45} {len(df_inv):>5} linhas")
print(f"  b) {_path_b.name:<45} {len(df_prod):>5} linhas")
print(f"  c) {_path_c.name:<45} {len(df_ch):>5} linhas")
print(f"  d) {_path_d.name:<45} {len(df_obra):>5} linhas")
try:
    print(f"  e) {_path_e.name:<45} {len(df_dir_out):>5} linhas")
except Exception:
    pass
print("=" * 60)
