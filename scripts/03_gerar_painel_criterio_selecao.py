"""
gerar_painel_criterio_selecao.py
---------------------------------
Gera painel_criterio_selecao.html na raiz do projeto.

Pergunta analítica: Editais seletivos que usaram histórico de resultados em festivais
internacionais como critério da Fase 1 (antes de ir para leitura de roteiro) produziram
obras com alcance internacional diferente dos editais que usaram critérios comerciais
(bilheteria, market) na Fase 1?

Fonte de dados: tabela_consolidada_obras.xlsx (sheet "Obras") + Festivais_por_obra.xlsx
Saída: painel_criterio_selecao.html (raiz do projeto)
"""

import json, os, re, unicodedata, math, statistics, pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ── Carrega dados principais ────────────────────────────────────────────────
obras_path = os.path.join(ROOT, "resultados", "tabela_consolidada_obras.xlsx")
df_obras = pd.read_excel(obras_path, sheet_name="Obras")
df_obras = df_obras.drop_duplicates(subset=["Projeto", "Ano"], keep="first").reset_index(drop=True)

# Garante colunas numéricas, preenchendo NaN com 0
_num_cols = [
    "Valor FSA (R$)",
    "Renúncia Art.3/3-A/39 (R$)",
    "Renúncia Outros Mec. (R$)",
    "Bilheteria Nominal (R$)",
    "Bilheteria Deflac. (R$)",
    "Estimativa Outras Janelas (R$)",
    "ROI Internacional (0-100)",
]
for col in _num_cols:
    if col in df_obras.columns:
        df_obras[col] = pd.to_numeric(df_obras[col], errors="coerce").fillna(0)
    else:
        df_obras[col] = 0.0

# ── Carrega scores de festivais ─────────────────────────────────────────────
def _normalize(s):
    """Remove acentos, pontuação e caixa para comparação."""
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s.lower())
    return re.sub(r"\s+", " ", s).strip()

festival_scores = {}   # normalized_title -> score (0-100)

try:
    _consol_csv = os.path.join(ROOT, "resultados", "festivais_consolidado.csv")
    df_fest = pd.read_csv(_consol_csv, encoding="utf-8-sig")
    df_fest["pontuacao_total"] = pd.to_numeric(df_fest.get("pontuacao_total", 0), errors="coerce").fillna(0)
    max_score = df_fest["pontuacao_total"].max()
    for _, row in df_fest.iterrows():
        tn = str(row.get("titulo_norm", "")).strip()
        if not tn:
            continue
        raw = float(row["pontuacao_total"])
        festival_scores[tn] = float((raw / max_score * 100)) if max_score > 0 else 0.0
except Exception as _e:
    # Se o arquivo não existir ou tiver problema, segue sem scores
    festival_scores = {}

# ── Mapeamento de Categoria (rótulos refinados) → chave JS ─────────────────
KEY_MAP = {
    # FSA — labels canônicos (de_para_chamadas_categorias.xlsx) → JS keys
    "FSA Pontuação Festivais e Roteiro":              "pont_fest_prod",
    "FSA Pontuação Bilheteria e Roteiro — Produtora": "pont_com_prod",
    "FSA Pontuação Bilheteria e Roteiro — Distribuidora": "pont_com_dist",
    "FSA Automático Bilheteria":                      "automatico",   # merge: Fluxo Contínuo + SUAT Comercial
    "FSA Automático Festivais":                       "automatico_fest",
    "FSA Coprodução Internacional":                   "coprod",
    "FSA SAV/MINC / Arranjos Regionais":              "sav_minc",
    "FSA Apenas roteiro":                             "seletivo_roteiro",
    "FSA Comercialização / Distribuição":             "comercializacao",
    "FSA Complementação":                             "complementacao",
    "Renúncia — Art.3/3-A/39":                        "renuncia_art3",
    "Renúncia — Outros Mecanismos":                   "renuncia_outros",
    # Excluídos do painel
    "_tv_excluir":     "_tv_excluir",
    "sem_categoria":   "_tv_excluir",
    "outros_seletivos":"outros_seletivos",
}

# ── Categorias com chaves JS ────────────────────────────────────────────────
CATEGORIAS = {
    "pont_fest_prod":   {"label": "FSA Pontuação Festivais e Roteiro",              "cor": "#6c7bf7"},
    "pont_com_prod":    {"label": "FSA Pontuação Bilheteria e Roteiro — Produtora", "cor": "#f5c842"},
    "pont_com_dist":    {"label": "FSA Pontuação Bilheteria e Roteiro — Distribuidora", "cor": "#f09020"},
    "automatico":       {"label": "FSA Automático Bilheteria",                      "cor": "#5b8cff"},
    "automatico_fest":  {"label": "FSA Automático Festivais",                       "cor": "#c080ff"},
    "coprod":           {"label": "FSA Coprodução Internacional",                   "cor": "#ff80b0"},
    "seletivo_roteiro": {"label": "FSA Apenas roteiro",                             "cor": "#80ffb0"},
    "comercializacao":  {"label": "FSA Comercialização / Distribuição",             "cor": "#a040e8"},
    "complementacao":   {"label": "FSA Complementação",                             "cor": "#ff60c0"},
}

MIN_OBRAS = 10  # categorias com < MIN_OBRAS obras são excluídas do painel

# Garante colunas novas (compatibilidade com xlsx antigos)
for _col in ["Todas Chamadas FSA", "Todos Valores FSA", "Tem Renuncia Art3", "Tem Renuncia Outros"]:
    if _col not in df_obras.columns:
        df_obras[_col] = ""
df_obras["Tem Renuncia Art3"]   = df_obras["Tem Renuncia Art3"].astype(str).str.upper().isin(["TRUE","1","SIM"])
df_obras["Tem Renuncia Outros"] = df_obras["Tem Renuncia Outros"].astype(str).str.upper().isin(["TRUE","1","SIM"])

# ── Cópia de get_categoria (self-contained, não depende do outro script) ─────
_CAT_TV = "_tv_excluir"
_CONCURSO_LIMIAR = 2_500_000
_CAT_PONT_FEST_PROD = "FSA Pontuação Festivais e Roteiro"
_CAT_PONT_COM_PROD  = "FSA Pontuação Bilheteria e Roteiro — Produtora"
_CAT_PONT_COM_DIST  = "FSA Pontuação Bilheteria e Roteiro — Distribuidora"
_CAT_CONC_A         = _CAT_PONT_COM_PROD   # Módulo A → Comercial
_CAT_CONC_B         = _CAT_PONT_FEST_PROD  # Módulo B → Festivais
_CAT_AUTO           = "FSA Automático Bilheteria"
_CAT_AUTO_DIST      = "FSA Automático Bilheteria"   # merged
_CAT_AUTO_COM       = "FSA Automático Bilheteria"   # merged
_CAT_AUTO_FEST      = "FSA Automático Festivais"
_CAT_COPROD         = "FSA Coprodução Internacional"
_CAT_SAV            = "FSA SAV/MINC / Arranjos Regionais"
_CAT_ROTEIRO        = "FSA Apenas roteiro"
_CAT_REN_ART3       = "Renúncia — Art.3/3-A/39"
_CAT_REN_OUT        = "Renúncia — Outros Mecanismos"
_CAT_COMERCIALIZ    = "FSA Comercialização / Distribuição"
_CAT_COMPL          = "FSA Complementação"

_CATEGORIA_MAP = {
    "PRODECINE 03": _CAT_COMERCIALIZ,
    "PRODECINE 04": _CAT_COMPL,
    "PRODECINE 05": _CAT_PONT_FEST_PROD,
    "PRODECINE 01": _CAT_PONT_COM_PROD,
    "PRODECINE 06": _CAT_COPROD,
    "PRODUÇÃO CINEMA": _CAT_PONT_COM_PROD,
    "PRODECINE 02": _CAT_PONT_COM_DIST,
    "CINEMA VIA DISTRIBUIDORA": _CAT_PONT_COM_DIST,
    "PRODUÇÃO CINEMA VIA DISTRIBUIDORA": _CAT_PONT_COM_DIST,
    "PRODECINE 07": _CAT_COPROD, "PRODECINE 08": _CAT_COPROD, "PRODECINE 09": _CAT_COPROD,
    "PRODECINE 10": _CAT_COPROD, "PRODECINE 11": _CAT_COPROD, "PRODECINE 12": _CAT_COPROD,
    "COPRODUÇÃO INTERNACIONAL": _CAT_COPROD, "COPRODUÇÃO INTERNACIONAL CINEMA": _CAT_COPROD,
    "CONCURSO PRODUÇÃO PARA CINEMA - COPRODUÇÃO CHILE-BRASIL": _CAT_COPROD,
    "CONCURSO PRODUÇÃO PARA CINEMA - COPRODUÇÃO PORTUGAL-BRASIL": _CAT_COPROD,
    "CONCURSO PRODUÇÃO PARA CINEMA - COPRODUÇÃO URUGUAI-BRASIL": _CAT_COPROD,
    "FLUXO CONTÍNUO PRODUÇÃO PARA CINEMA": _CAT_AUTO,
    "FLUXO CONTÍNUO PRODUÇÃO PARA CINEMA - VIA DISTRIBUIDORA": _CAT_AUTO_DIST,
    "SUPORTE AUTOMÁTICO - DESEMPENHO COMERCIAL CINEMA": _CAT_AUTO_COM,
    "COMPLEMENTAÇÃO": _CAT_COMPL,
    "COMERCIALIZAÇÃO EM CINEMA": _CAT_COMERCIALIZ,
    "COMERCIALIZAÇÃO - OPÇÃO DE INVESTIMENTO EM COMERCIALIZAÇÃO": _CAT_COMERCIALIZ,
    "SUPORTE AUTOMÁTICO - DESEMPENHO ARTÍSTICO": _CAT_AUTO_FEST,
    "SAV/MINC 01": _CAT_SAV, "SAV/MINC 02": _CAT_SAV, "SAV/MINC 03": _CAT_SAV,
    "SAV/MINC 04": _CAT_SAV, "SAV/MINC 05": _CAT_SAV, "SAV/MINC 06": _CAT_SAV,
    "SAV/MINC 07": _CAT_SAV, "SAV/MINC 08": _CAT_SAV, "SAV/MINC 09": _CAT_SAV,
    "SAV/MINC 10": _CAT_SAV, "SAV/MINC 11": _CAT_SAV, "SAV/MINC 13": _CAT_SAV,
    "ARRANJOS REGIONAIS": _CAT_SAV,
    "CINEMA NOVOS REALIZADORES": _CAT_ROTEIRO,
    "CONCURSO PRODUÇÃO PARA CINEMA - MODULO A": _CAT_PONT_COM_PROD,
    "CONCURSO PRODUÇÃO PARA CINEMA - Modulo A": _CAT_PONT_COM_PROD,
    "CONCURSO PRODUÇÃO PARA CINEMA - MODULO B": _CAT_PONT_FEST_PROD,
    "CONCURSO PRODUÇÃO PARA CINEMA - Modulo B": _CAT_PONT_FEST_PROD,
    "SUPORTE AUTOMÁTICO - DESEMPENHO COMERCIAL TV E VOD": _CAT_TV,
    "FLUXO CONTÍNUO PRODUÇÃO PARA TELEVISÃO": _CAT_TV,
    "PRODUÇÃO TV-VOD": _CAT_TV, "PRODUÇÃO TV/VOD - NOVOS REALIZADORES": _CAT_TV,
    "PRODUÇÃO TV/VOD – VIA PROGRAMADORA": _CAT_TV,
    "PRODAV 01": _CAT_TV, "PRODAV 02 - PROJETO DERIVADO": _CAT_TV, "PRODAV 03": _CAT_TV,
    "PRODAV 05": _CAT_TV, "PRODAV 06": _CAT_TV, "PRODAV 07": _CAT_TV, "PRODAV 08": _CAT_TV,
    "PRODAV 09": _CAT_TV, "PRODAV 10": _CAT_TV, "PRODAV 11": _CAT_TV, "PRODAV 12": _CAT_TV,
    "PRODAV 13": _CAT_TV, "PRODAV 14": _CAT_TV, "PRODAV - TVS PÚBLICAS": _CAT_TV,
}

def _get_cat(chamada, valor=0):
    if not chamada: return ""
    ch = chamada.strip()
    if ch == "CONCURSO PRODUÇÃO PARA CINEMA":
        return _CAT_PONT_FEST_PROD
    cat = _CATEGORIA_MAP.get(ch)
    if cat: return cat
    up = ch.upper()
    if any(p in up for p in ["PRODAV","TV-VOD","TV/VOD","TELEVISAO","TELEVISÃO","PRODUÇÃO TV"]): return _CAT_TV
    if "COPROD" in up: return _CAT_COPROD
    if "SAV/MINC" in up or "ARRANJOS" in up: return _CAT_SAV
    if "VIA DISTRIBUIDORA" in up and "FLUXO CONTIN" in up: return _CAT_AUTO_DIST
    if "FLUXO CONTIN" in up: return _CAT_AUTO
    if "COMPLEMENT" in up: return _CAT_COMPL
    if "COMERCIALIZ" in up: return _CAT_COMERCIALIZ
    if "DESEMPENHO ART" in up: return _CAT_AUTO_FEST
    if "SUPORTE AUTO" in up or "DESEMPENHO COM" in up: return _CAT_AUTO_COM
    return "outros_seletivos"

# ── Registros com sobreposição: um por (obra, categoria válida) ──────────────
from collections import defaultdict

_TV_KEY = "_tv_excluir"
obras_registros = []

for _, row in df_obras.iterrows():
    titulo     = str(row.get("Projeto", "") or "")
    ano        = row.get("Ano", 0) or 0
    ano_int    = int(ano) if ano else 0
    ren_art3   = float(row.get("Renúncia Art.3/3-A/39 (R$)", 0) or 0)
    ren_outros = float(row.get("Renúncia Outros Mec. (R$)", 0) or 0)
    bilh_nom   = float(row.get("Bilheteria Nominal (R$)", 0) or 0)
    bilh_def   = float(row.get("Bilheteria Deflac. (R$)", 0) or 0)
    est_jan    = float(row.get("Estimativa Outras Janelas (R$)", 0) or 0)
    est_jan_def = float(row.get("Outras Janelas Deflac. (R$2024)", 0) or 0)
    fsa_def_total = float(row.get("Valor FSA Deflac. (R$2024)", 0) or 0)
    ren_def_total = float(row.get("Renúncia Total Deflac. (R$2024)", 0) or 0)
    roi_intl_v = float(row.get("ROI Internacional (0-100)", 0) or 0)
    critica_n  = int(pd.to_numeric(row.get("CRITICA_N_FONTES", 0), errors="coerce") or 0)
    revenue    = bilh_nom
    revenue_def = bilh_def   # receita deflacionada (R$2024) — só bilheteria real

    # Festivais por nome (colunas do xlsx)
    _FEST_COLS = [
        ("Festival — Oscar",       "Oscar"),
        ("Festival — Cannes",      "Cannes"),
        ("Festival — Berlim",      "Berlim"),
        ("Festival — Veneza",      "Veneza"),
        ("Festival — BAFTA",       "BAFTA"),
        ("Festival — Globo Ouro",  "Globo de Ouro"),
        ("Festival — Outros Intl", "Outros Intl"),
        ("Festival — Brasília",    "Brasília"),
        ("Festival — Gramado",     "Gramado"),
        ("Festival — Fest.Rio",    "Fest.Rio"),
    ]
    festivais_list = [short for col, short in _FEST_COLS
                      if pd.to_numeric(row.get(col, 0), errors="coerce") or 0 > 0]
    n_paises   = int(pd.to_numeric(row.get("VOD Intl — N Países", 0), errors="coerce") or 0)
    adm_lumiere = int(pd.to_numeric(row.get("Adm. EU — Lumière", 0), errors="coerce") or 0)

    norm_titulo    = _normalize(titulo)
    festival_score = festival_scores.get(norm_titulo, 0.0)

    # Chamadas FSA (sobreposição)
    chs_raw  = str(row.get("Todas Chamadas FSA", "") or "")
    vals_raw = str(row.get("Todos Valores FSA",  "") or "")
    chs  = [c.strip() for c in chs_raw.split("|") if c.strip()]
    vals = []
    for v in vals_raw.split("|"):
        try:    vals.append(float(v.strip()))
        except: vals.append(0.0)
    while len(vals) < len(chs):
        vals.append(0.0)

    # Ratio de deflação FSA — aplica-se proporcionalmente às chamadas individuais
    _fsa_nom_total = sum(vals)
    _fsa_deflat_ratio = (fsa_def_total / _fsa_nom_total) if _fsa_nom_total > 0 else 1.0

    # Coleta entradas da obra antes de distribuir receita
    obra_entries = []
    has_cinema_fsa = False
    for i, cham in enumerate(chs):
        inv_c     = vals[i]
        inv_c_def = inv_c * _fsa_deflat_ratio   # FSA deflacionado (R$2024)
        cat_label = _get_cat(cham, inv_c)
        cat_key   = KEY_MAP.get(cat_label, "outros_seletivos")
        if cat_key in (_TV_KEY, "sem_categoria", "outros_seletivos") or cat_key not in CATEGORIAS:
            continue
        has_cinema_fsa = True
        obra_entries.append({"chamada": cham, "cat_key": cat_key, "investment": inv_c, "investment_def": inv_c_def})

    if not has_cinema_fsa:
        if not obra_entries:
            continue

    if not obra_entries:
        continue

    # Receita total da obra (usada para ROI por linha individual)
    # Receita proporcional (usada no agregado de categoria, evita dupla contagem)
    total_inv     = sum(e["investment"]     for e in obra_entries)
    total_inv_def = sum(e.get("investment_def", e["investment"]) for e in obra_entries)
    n_entries = len(obra_entries)
    for e in obra_entries:
        inv_def = e.get("investment_def", e["investment"])
        if total_inv > 0:
            prop_rev     = revenue     * (e["investment"] / total_inv)
            prop_rev_def = revenue_def * (inv_def / max(total_inv_def, 1))
        else:
            prop_rev     = revenue     / n_entries
            prop_rev_def = revenue_def / n_entries
        # ROI por linha: FSA desta chamada e total (FSA + renúncia da obra)
        _is_ren = e["cat_key"] in ("renuncia_art3", "renuncia_outros")
        _fsa_this     = 0 if _is_ren else e["investment"]
        _fsa_this_def = 0 if _is_ren else inv_def
        _tot_denom     = _fsa_this     + ren_art3 + ren_outros
        _tot_denom_def = _fsa_this_def + ren_def_total
        obras_registros.append({
            "titulo": titulo, "ano": ano_int, "chamada": e["chamada"],
            "cat_key": e["cat_key"], "investment": e["investment"],
            "revenue": prop_rev,        # proporcional nominal (para agregado)
            "revenue_def": prop_rev_def, # proporcional deflac. (para agregado R$2024)
            "full_rev": revenue,        # receita nominal total da obra
            "full_rev_def": revenue_def, # receita deflac. total da obra (R$2024)
            "festival_score": festival_score, "roi_intl": roi_intl_v,
            "bilh": bilh_nom, "bilh_def": bilh_def, "janelas": est_jan,
            "n_paises": n_paises, "adm_lumiere": adm_lumiere,
            "festivais": festivais_list, "critica_n": critica_n,
            "fsa_this": _fsa_this,           # FSA nominal desta chamada (0 se renúncia)
            "fsa_this_def": _fsa_this_def,   # FSA deflac. desta chamada (0 se renúncia)
            "tot_denom": _tot_denom,         # FSA + renúncia nominal da obra
            "tot_denom_def": _tot_denom_def, # FSA + renúncia deflac. da obra (R$2024)
            # ROI nominal (compatibilidade)
            "roi_fsa": revenue / _fsa_this  if _fsa_this  >= 1_000 else 0,
            "roi_tot": revenue / _tot_denom if _tot_denom >= 1_000 else 0,
            # ROI deflacionado (R$2024) — comentário: receita deflac / FSA deflac
            "roi_fsa_def": revenue_def / _fsa_this_def  if _fsa_this_def  >= 1_000 else 0,
            "roi_tot_def": revenue_def / _tot_denom_def if _tot_denom_def >= 1_000 else 0,
        })

# ── Agregação por (chamada, cat_key) ─────────────────────────────────────────
_grp = defaultdict(list)
for o in obras_registros:
    _grp[(o["chamada"], o["cat_key"])].append(o)

def _agg(nome, lst, cat_key):
    if not lst: return None
    inv = sum(o["investment"] for o in lst)
    rev = sum(o["revenue"]    for o in lst)
    ri  = [o["roi_intl"] for o in lst]
    fs  = [o["festival_score"] for o in lst]
    rd_ind = [(o["full_rev"]/o["investment"], o["investment"]) for o in lst if o["investment"] >= 1_000]
    _fsa_d = sum(o.get("fsa_this", 0) for o in lst)
    _fsa_n = sum(o["full_rev"] for o in lst if o.get("fsa_this", 0) > 0)
    _tot_d = sum(o.get("tot_denom", 0) for o in lst)
    _tot_n = sum(o["full_rev"] for o in lst if o.get("tot_denom", 0) > 0)
    # Deflacionado
    _fsa_d_def = sum(o.get("fsa_this_def", 0) for o in lst)
    _fsa_n_def = sum(o.get("full_rev_def", 0) for o in lst if o.get("fsa_this_def", 0) > 0)
    _tot_d_def = sum(o.get("tot_denom_def", 0) for o in lst)
    _tot_n_def = sum(o.get("full_rev_def", 0) for o in lst if o.get("tot_denom_def", 0) > 0)
    bilh_lst = [o.get("bilh",0) for o in lst]
    pais_lst = [o.get("n_paises",0) for o in lst if o.get("n_paises",0)>0]
    works = sorted(
        [{"title": o["titulo"], "year": o["ano"],
          "domestic": (o["full_rev"]/o["investment"]) if o["investment"]>0 else 0,
          "international": o["festival_score"], "roi_intl": o["roi_intl"],
          "investment": o["investment"],
          "bilh": o.get("bilh", 0), "janelas": o.get("janelas", 0),
          "n_paises": o.get("n_paises", 0), "adm_lumiere": o.get("adm_lumiere", 0),
          "festivais": o.get("festivais", [])}
         for o in lst],
        key=lambda x: x["domestic"], reverse=True)[:30]
    return {"name": nome, "fase1": cat_key, "workCount": len(lst),
            "investment": inv, "revenue": rev,
            "aggregateDomestic": (rev/inv) if inv>0 else 0,
            "avgDomestic": sum(r*w for r,w in rd_ind)/sum(w for _,w in rd_ind) if rd_ind else 0,
            "roi_fsa_agg": _fsa_n / _fsa_d if _fsa_d > 0 else 0,
            "roi_tot_agg": _tot_n / _tot_d if _tot_d > 0 else 0,
            # ROI deflacionado (R$2024) — receita deflac. / FSA deflac.
            "roi_fsa_def": _fsa_n_def / _fsa_d_def if _fsa_d_def > 0 else 0,
            "roi_tot_def": _tot_n_def / _tot_d_def if _tot_d_def > 0 else 0,
            "intlAverage": sum(ri)/len(ri) if ri else 0,
            "intlPeak": max(ri, default=0),
            "festAverage": sum(fs)/len(fs) if fs else 0,
            "bilhAvg": sum(bilh_lst)/len(bilh_lst) if bilh_lst else 0,
            "paisesAvg": sum(pais_lst)/len(pais_lst) if pais_lst else 0,
            "works": works}

chamadas = [r for r in (_agg(n, l, k) for (n,k),l in _grp.items()) if r]

# ── Deduplicação por (titulo, cat_key) ───────────────────────────────────────
# Uma obra pode ter múltiplas chamadas mapeadas para a mesma categoria.
# Métricas de obra-nível (bilh, janelas, festivais, VOD, cobertura de dados)
# devem contá-la uma única vez por categoria.
_seen_obra_cat = set()
_obras_dedup   = []
for _o in obras_registros:
    _k = (_o["titulo"], _o["cat_key"])
    if _k not in _seen_obra_cat:
        _seen_obra_cat.add(_k)
        _obras_dedup.append(_o)

# ROI dom: usa apenas obras com retorno mensurável.
# Obras com receita deflacionada total = 0 E sem presença internacional
# (roi_intl=0, adm_lumiere=0, n_paises=0) são excluídas do numerador E do
# denominador do ROI. Usa full_rev_def (bilh_def + jan_def) para evitar
# inconsistências entre colunas nominais e deflacionadas.
# Permanecem contadas em n_obras e n_sem_dados para transparência da cobertura.
_cat_roi_dom_vals     = defaultdict(list)  # lista de (roi, inv)
_cat_roi_fsa_def_vals = defaultdict(list)
for _o in obras_registros:
    _tem_ret = (_o.get("full_rev_def", 0) > 0 or _o.get("roi_intl", 0) > 0
                or _o.get("adm_lumiere", 0) > 0 or _o.get("n_paises", 0) > 0)
    if not _tem_ret:
        continue
    if _o["investment"] >= 1_000:
        _cat_roi_dom_vals[_o["cat_key"]].append((_o["full_rev"] / _o["investment"], _o["investment"]))
    if _o.get("fsa_this_def", 0) >= 1_000:
        _cat_roi_fsa_def_vals[_o["cat_key"]].append(_o.get("roi_fsa_def", 0))

# ROI agregado por categoria: sum(full_rev) / sum(tot_denom)
# Apenas obras com retorno mensurável (mesma regra do loop acima).
_cat_roi_fsa_num = defaultdict(float)
_cat_roi_fsa_den = defaultdict(float)
_cat_roi_tot_num = defaultdict(float)
_cat_roi_tot_den = defaultdict(float)
# ROI deflacionado por categoria
_cat_roi_fsa_def_num = defaultdict(float)
_cat_roi_fsa_def_den = defaultdict(float)
_cat_roi_tot_def_num = defaultdict(float)
_cat_roi_tot_def_den = defaultdict(float)
for _o in obras_registros:
    _tem_ret = (_o.get("full_rev_def", 0) > 0 or _o.get("roi_intl", 0) > 0
                or _o.get("adm_lumiere", 0) > 0 or _o.get("n_paises", 0) > 0)
    if not _tem_ret:
        continue
    _fth = _o.get("fsa_this", 0)
    _tdn = _o.get("tot_denom", 0)
    _fth_def = _o.get("fsa_this_def", 0)
    _tdn_def = _o.get("tot_denom_def", 0)
    if _fth > 0:
        _cat_roi_fsa_num[_o["cat_key"]] += _o["full_rev"]
        _cat_roi_fsa_den[_o["cat_key"]] += _fth
    if _tdn > 0:
        _cat_roi_tot_num[_o["cat_key"]] += _o["full_rev"]
        _cat_roi_tot_den[_o["cat_key"]] += _tdn
    if _fth_def > 0:
        _cat_roi_fsa_def_num[_o["cat_key"]] += _o.get("full_rev_def", 0)
        _cat_roi_fsa_def_den[_o["cat_key"]] += _fth_def
    if _tdn_def > 0:
        _cat_roi_tot_def_num[_o["cat_key"]] += _o.get("full_rev_def", 0)
        _cat_roi_tot_def_den[_o["cat_key"]] += _tdn_def

# Métricas de obra-nível: lista deduplicada
_cat_bilh_vals   = defaultdict(list)
_cat_jan_vals    = defaultdict(list)
_cat_paises_vals = defaultdict(list)
_cat_intl_vals   = defaultdict(list)
_cat_n_distinct  = defaultdict(int)
for _o in _obras_dedup:
    _cat_n_distinct[_o["cat_key"]] += 1
    _cat_bilh_vals[_o["cat_key"]].append(_o.get("bilh", 0))
    _cat_jan_vals[_o["cat_key"]].append(_o.get("janelas", 0))
    _np = _o.get("n_paises", 0)
    if _np > 0:
        _cat_paises_vals[_o["cat_key"]].append(_np)
    _cat_intl_vals[_o["cat_key"]].append(_o["roi_intl"])

_cat_com_dados = defaultdict(int)
_cat_sem_dados = defaultdict(int)
for _od in _obras_dedup:
    _tem = (_od.get("bilh",0)>0 or _od.get("janelas",0)>0
            or len(_od.get("festivais",[]))>0
            or _od.get("adm_lumiere",0)>0 or _od.get("n_paises",0)>0)
    if _tem: _cat_com_dados[_od["cat_key"]] += 1
    else:    _cat_sem_dados[_od["cat_key"]] += 1

_cat_fest_count = defaultdict(int)
_cat_total_fest  = defaultdict(int)
for _o in _obras_dedup:
    nf = len(_o.get("festivais", []))
    if nf > 0:
        _cat_fest_count[_o["cat_key"]] += 1
    _cat_total_fest[_o["cat_key"]] += nf

def aggregate_cat(cat_key):
    items = [c for c in chamadas if c["fase1"] == cat_key]
    n_entries  = sum(c["workCount"] for c in items)
    n_distinct = _cat_n_distinct.get(cat_key, n_entries)
    if n_distinct < MIN_OBRAS: return None
    inv = sum(c["investment"] for c in items)
    rev = sum(c["revenue"]    for c in items)
    iv  = _cat_intl_vals.get(cat_key, [])
    rd_vals = _cat_roi_dom_vals.get(cat_key, [])
    _rf_n = _cat_roi_fsa_num.get(cat_key, 0)
    _rf_d = _cat_roi_fsa_den.get(cat_key, 0)
    _rt_n = _cat_roi_tot_num.get(cat_key, 0)
    _rt_d = _cat_roi_tot_den.get(cat_key, 0)
    bv = _cat_bilh_vals.get(cat_key, [])
    jv = _cat_jan_vals.get(cat_key, [])
    pv = _cat_paises_vals.get(cat_key, [])
    # Classificação: seletivo (tem leitura de roteiro) vs automático
    _SELETIVOS = {"pont_fest_prod","pont_com_prod","pont_com_dist","sav_minc","coprod","seletivo_roteiro","complementacao","comercializacao"}
    _is_seletivo = cat_key in _SELETIVOS
    _roi_fsa_def_n = _cat_roi_fsa_def_num.get(cat_key, 0)
    _roi_fsa_def_d = _cat_roi_fsa_def_den.get(cat_key, 0)
    _roi_tot_def_n = _cat_roi_tot_def_num.get(cat_key, 0)
    _roi_tot_def_d = _cat_roi_tot_def_den.get(cat_key, 0)
    _roi_fsa_def_med = statistics.median(_cat_roi_fsa_def_vals.get(cat_key, [0])) if _cat_roi_fsa_def_vals.get(cat_key) else 0
    return {"key": cat_key, "label": CATEGORIAS[cat_key]["label"], "cor": CATEGORIAS[cat_key]["cor"],
            "is_seletivo": _is_seletivo,   # True = seletivo c/ leitura de roteiro
            "n_chamadas": len(items), "n_obras": n_distinct, "investment": inv, "revenue": rev,
            "roi_dom_agg": (rev/inv) if inv>0 else 0,
            "roi_dom_avg": sum(r*w for r,w in rd_vals)/sum(w for _,w in rd_vals) if rd_vals else 0,
            "roi_fsa_agg": _rf_n / _rf_d if _rf_d > 0 else 0,
            "roi_tot_agg": _rt_n / _rt_d if _rt_d > 0 else 0,
            # ROI deflacionado (R$2024) — métricas principais para v2
            "roi_fsa_def_agg": _roi_fsa_def_n / _roi_fsa_def_d if _roi_fsa_def_d > 0 else 0,
            "roi_tot_def_agg": _roi_tot_def_n / _roi_tot_def_d if _roi_tot_def_d > 0 else 0,
            "roi_fsa_def_med": _roi_fsa_def_med,
            "intl_avg": sum(iv)/len(iv) if iv else 0,
            "intl_peak": max((c["intlPeak"] for c in items), default=0), "evid_avg": 0,
            "bilh_total": sum(bv), "bilh_avg": sum(bv)/len(bv) if bv else 0,
            "jan_total": sum(jv), "jan_avg": sum(jv)/len(jv) if jv else 0,
            "paises_avg": sum(pv)/len(pv) if pv else 0, "paises_total": sum(pv), "n_com_paises": len(pv), "n_com_festivais": _cat_fest_count.get(cat_key,0), "total_festivais": _cat_total_fest.get(cat_key,0), "n_com_dados": _cat_com_dados.get(cat_key,0), "n_sem_dados": _cat_sem_dados.get(cat_key,0)}

cats = [aggregate_cat(k) for k in CATEGORIAS]
cats = [c for c in cats if c]
valid_keys = {c["key"] for c in cats}
chamadas   = [c for c in chamadas if c["fase1"] in valid_keys]

# ── Obras excluídas por categoria (para modal dos cards) ─────────────────────
_included_by_cat = defaultdict(set)
for _oo in obras_registros:
    if _oo["cat_key"] in valid_keys:
        _included_by_cat[_oo["cat_key"]].add(_oo["titulo"])

_excluded_per_cat = defaultdict(list)
for _, _row2 in df_obras.iterrows():
    _t2 = str(_row2.get("Projeto", "") or "").strip()
    if not _t2: continue
    _a2 = int(_row2.get("Ano", 0) or 0)
    _b2 = round(float(_row2.get("Bilheteria Nominal (R$)", 0) or 0))
    _chs2 = [c.strip() for c in str(_row2.get("Todas Chamadas FSA", "") or "").split("|") if c.strip()]
    _vs2 = []
    for _v in str(_row2.get("Todos Valores FSA", "") or "").split("|"):
        try: _vs2.append(float(_v.strip()))
        except: _vs2.append(0.0)
    while len(_vs2) < len(_chs2): _vs2.append(0.0)
    for _i, _cham2 in enumerate(_chs2):
        _ck2 = KEY_MAP.get(_get_cat(_cham2, _vs2[_i]), "")
        if _ck2 in valid_keys and _t2 not in _included_by_cat[_ck2]:
            _excluded_per_cat[_ck2].append({"t": _t2, "a": _a2, "bilh": _b2})
    for _ck2, _val2 in [
        ("renuncia_art3",   float(_row2.get("Renúncia Art.3/3-A/39 (R$)", 0) or 0)),
        ("renuncia_outros", float(_row2.get("Renúncia Outros Mec. (R$)", 0) or 0)),
    ]:
        if _val2 > 0 and _ck2 in valid_keys and _t2 not in _included_by_cat[_ck2]:
            _excluded_per_cat[_ck2].append({"t": _t2, "a": _a2, "bilh": _b2})

excl_js = {k: sorted(v, key=lambda x: -x["bilh"])[:50]
           for k, v in _excluded_per_cat.items()}

# ── Timeline: investimento por (ano, cat_key) ─────────────────────────────
_tl_acc = defaultdict(lambda: defaultdict(float))
for _o in obras_registros:
    if _o["cat_key"] in valid_keys and _o["ano"] > 0 and _o["investment"] > 0:
        _tl_acc[_o["cat_key"]][_o["ano"]] += _o["investment"]
_tl_anos = sorted({a for cat_d in _tl_acc.values() for a in cat_d})
timeline_js = {
    "anos": _tl_anos,
    "series": [
        {"key": k, "label": CATEGORIAS[k]["label"], "cor": CATEGORIAS[k]["cor"],
         "values": [_tl_acc[k].get(a, 0) for a in _tl_anos]}
        for k in CATEGORIAS if k in valid_keys
    ]
}

# ── Prepara JS ───────────────────────────────────────────────────────────────
def safe_json(obj):
    return json.dumps(obj, ensure_ascii=False).replace("</", "<\\/")

chamadas_js = []
for c in chamadas:
    chamadas_js.append({
        "ch": c["name"], "f1": c["fase1"],
        "f1l": CATEGORIAS[c["fase1"]]["label"], "cor": CATEGORIAS[c["fase1"]]["cor"],
        "n": c["workCount"], "inv": round(c["investment"]), "rec": round(c["revenue"]),
        "rda": round(c["aggregateDomestic"],4),
        "rda_avg": round(c["avgDomestic"],4),
        "roi_fsa": round(c.get("roi_fsa_agg",0),4),
        "roi_tot": round(c.get("roi_tot_agg",0),4),
        "roi_fsa_def": round(c.get("roi_fsa_def",0),4),
        "roi_tot_def": round(c.get("roi_tot_def",0),4),
        "rim": round(c["intlAverage"],2), "rimax": round(c["intlPeak"],2),
        "bilh_avg": round(c["bilhAvg"]), "paises_avg": round(c["paisesAvg"],1),
        "evid": 0,
        "show_scatter": c["workCount"] >= 3,
        "obras": [{"t": w["title"], "a": w.get("year",0),
                   "rd": round(w["domestic"],4), "ri": round(w["international"],2),
                   "ri2": round(w.get("roi_intl",0),2), "inv": round(w["investment"]),
                   "bilh": round(w.get("bilh",0)), "jan": round(w.get("janelas",0)),
                   "np": w.get("n_paises",0), "adm": w.get("adm_lumiere",0),
                   "fest": w.get("festivais",[])}
                  for w in (c.get("works") or [])[:30]],
    })

cats_js = [{"key": c["key"], "label": c["label"], "cor": c["cor"],
            "is_seletivo": c.get("is_seletivo", False),
            "n_chamadas": c["n_chamadas"], "n_obras": c["n_obras"],
            "inv": round(c["investment"]), "rev": round(c["revenue"]),
            "rda": round(c["roi_dom_agg"],4),
            "rda_avg": round(c["roi_dom_avg"],4),
            "roi_fsa": round(c["roi_fsa_agg"],4),
            "roi_tot": round(c["roi_tot_agg"],4),
            # Deflacionado (R$2024) — métricas primárias para análise
            "roi_fsa_def": round(c.get("roi_fsa_def_agg",0),4),
            "roi_tot_def": round(c.get("roi_tot_def_agg",0),4),
            "roi_fsa_def_med": round(c.get("roi_fsa_def_med",0),4),
            "intl_avg": round(c["intl_avg"],2), "intl_peak": round(c["intl_peak"],2),
            "evid": round(c["evid_avg"],1),
            "bilh_total": round(c["bilh_total"]), "bilh_avg": round(c["bilh_avg"]),
            "jan_total": round(c["jan_total"]), "jan_avg": round(c["jan_avg"]),
            "paises_avg": round(c["paises_avg"],1), "paises_total": c["paises_total"], "n_com_paises": c["n_com_paises"], "n_com_festivais": c["n_com_festivais"], "total_festivais": c["total_festivais"], "bilh_prop": round(c["revenue"]), "n_com_dados": c["n_com_dados"], "n_sem_dados": c["n_sem_dados"]}
           for c in cats]

obras_det_js = []
for o in obras_registros:
    if o["cat_key"] not in valid_keys: continue
    inv = o["investment"]; full_rev = o.get("full_rev", o["revenue"])
    obras_det_js.append({
        "titulo": o["titulo"], "ano": o["ano"],
        "cat_key": o["cat_key"], "cat_label": CATEGORIAS[o["cat_key"]]["label"],
        "cor": CATEGORIAS[o["cat_key"]]["cor"], "chamada": o["chamada"],
        "inv": round(inv), "rev": round(full_rev),
        "roi_dom": round(full_rev/inv,4) if inv>0 else 0,
        "roi_fsa": round(o.get("roi_fsa",0),4),
        "roi_tot": round(o.get("roi_tot",0),4),
        # Deflacionado
        "roi_fsa_def": round(o.get("roi_fsa_def",0),4),
        "roi_tot_def": round(o.get("roi_tot_def",0),4),
        "roi_intl": round(o["roi_intl"],2), "fest": round(o["festival_score"],1),
        "bilh": round(o.get("bilh",0)), "bilh_def": round(o.get("bilh_def",0)),
        "jan": round(o.get("janelas",0)),
        "n_paises": o.get("n_paises",0), "adm": o.get("adm_lumiere",0),
        "festivais": o.get("festivais",[]),
        "n_festivais": len(o.get("festivais",[])),
        "critica_n": o.get("critica_n", 0),
    })

summary = {
    "totalUniverse": len(df_obras),
    "calculable": int((df_obras["Valor FSA (R$)"] + df_obras["Renúncia Art.3/3-A/39 (R$)"] + df_obras["Renúncia Outros Mec. (R$)"]).gt(0).sum()),
    "totalInvestment": float(df_obras["Valor FSA (R$)"].sum()),
    "totalRevenue": float((df_obras["Bilheteria Nominal (R$)"] + df_obras["Estimativa Outras Janelas (R$)"]).sum()),
}

# ── Best category per metric (FSA-only, excl. standalone renúncia) ──────────
_RENUNCIA_KEYS = {"renuncia_art3", "renuncia_outros"}
_fsa_cats = [c for c in cats if c["key"] not in _RENUNCIA_KEYS]

_METRIC_DEFS = [
    ("roi_tot_def_agg",  "ROI Total Deflacionado (R$2024)",      "higher", "Retorno por real de capital público total (FSA + renúncia fiscal). Métrica primária."),
    ("roi_fsa_def_agg",  "ROI FSA Deflacionado (R$2024)",       "higher", "Retorno por real investido apenas pelo FSA direto (excluindo renúncia fiscal)."),
    ("roi_fsa_def_med",  "ROI FSA Deflacionado — Mediana",       "higher", "Mediana dos ROIs por obra — menos sensível a outliers que a média agregada."),
    ("roi_dom_agg",      "ROI Doméstico Agregado (nominal)",     "higher", "Retorno doméstico proporcional ao financiamento, sem deflação."),
    ("intl_avg",         "ROI Internacional — Média",            "higher", "Score médio de presença internacional (0–100): festivais, VOD, EU."),
    ("intl_peak",        "ROI Internacional — Pico",             "higher", "Score máximo de presença internacional na categoria."),
    ("evid_avg",         "Pontuação Festivais Internacionais",   "higher", "Score médio de festivais internacionais (Oscar, Cannes, Berlim, Veneza…)."),
    ("paises_avg",       "Países Alcançados (média por obra)",   "higher", "Média de países com distribuição VOD por obra da categoria."),
]

best_by_metric = []
for _mk, _ml, _dir, _desc in _METRIC_DEFS:
    _ranked = sorted(_fsa_cats, key=lambda c: c.get(_mk, 0), reverse=True)
    _top3 = []
    for _c in _ranked[:3]:
        _v = _c.get(_mk, 0)
        if _v <= 0:
            break
        _top3.append({"key": _c["key"], "label": _c["label"], "cor": _c["cor"],
                      "val": round(_v, 4), "n": _c.get("n_obras", 0)})
    best_by_metric.append({"metric": _mk, "label": _ml, "desc": _desc, "top3": _top3})

sintese_js = best_by_metric

rank_options = "\n".join(
    f'      <option value="{c["key"]}">{c["label"]}</option>'
    for c in cats
)

def fmt_money(v):
    return f"R$ {v/1e9:.1f}B" if v >= 1e9 else f"R$ {v/1e6:.0f}M"

def fmt_int(v):
    return f"{v:,.0f}".replace(",", ".")


# ── HTML ───────────────────────────────────────────────────────────────────
ANO_MAX = 2023  # para referência no texto do painel

html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cinema Brasileiro — Metodologia de Decisão de Investimento</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Mono:wght@300;400;500&family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{{
  --bg:#0b0d14;--surface:#14171f;--surface2:#1a1e2c;--surface3:#212638;
  --border:#282d42;--border-light:#343a54;
  --accent:#6c7bf7;--accent-dim:rgba(108,123,247,.12);
  --gold:#fbbf24;--coral:#f87171;--purple:#a78bfa;--cyan:#38bdf8;
  --green:#34d399;--muted-blue:#5fd1ff;
  --text:#e2e8f0;--text2:#c1c9d9;--muted:#7b849a;--dim:#282d42;
  --font-head:'DM Serif Display',serif;--font-mono:'DM Mono',monospace;--font-ui:'Inter',system-ui,sans-serif;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:var(--font-mono);font-size:13px;min-height:100vh}}
.header{{padding:20px 32px;border-bottom:1px solid var(--border);background:linear-gradient(180deg,#0c0e18 0%,var(--bg) 100%);display:flex;align-items:center;gap:20px}}
.header-icon{{width:46px;height:46px;border-radius:12px;background:var(--accent);display:flex;align-items:center;justify-content:center;flex-shrink:0}}
.header-icon svg{{width:22px;height:22px}}
.header-text h1{{font-family:var(--font-head);font-size:20px;font-weight:400;letter-spacing:-.3px;line-height:1.1}}
.header-text p{{font-size:11px;color:var(--muted);margin-top:3px}}
.header-meta{{margin-left:auto;text-align:right;font-size:10px;color:var(--muted);line-height:1.7}}
.tabs{{display:flex;padding:0 32px;border-bottom:1px solid var(--border);background:var(--surface)}}
.tab{{padding:13px 22px;font-size:11px;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;transition:all .2s;letter-spacing:.06em;text-transform:uppercase;font-family:var(--font-mono)}}
.tab.active{{color:var(--accent);border-bottom-color:var(--accent)}}
.tab:hover:not(.active){{color:var(--text)}}
.panel{{display:none;padding:28px 32px;animation:fadeIn .2s ease}}
.panel.active{{display:block}}
@keyframes fadeIn{{from{{opacity:0;transform:translateY(4px)}}to{{opacity:1;transform:translateY(0)}}}}
.kpi-bar{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:24px}}
.kpi{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px;position:relative;overflow:hidden}}
.kpi::after{{content:'';position:absolute;inset:0;background:radial-gradient(circle at 80% 20%,rgba(108,123,247,.04),transparent 70%);pointer-events:none}}
.kpi-label{{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;margin-bottom:8px}}
.kpi-val{{font-family:var(--font-head);font-size:24px;font-style:italic;color:var(--accent);line-height:1}}
.kpi-sub{{font-size:10px;color:var(--muted);margin-top:5px}}
.cat-pills{{display:flex;gap:10px;margin-bottom:20px;flex-wrap:wrap}}
.cat-pill{{padding:8px 18px;border-radius:20px;border:1px solid var(--border);background:var(--surface);cursor:pointer;font-family:var(--font-mono);font-size:11px;color:var(--muted);transition:all .2s}}
.cat-pill.active{{border-color:var(--accent);color:var(--accent);background:var(--accent-dim)}}
.cat-pill:hover:not(.active){{color:var(--text);border-color:var(--muted)}}
.chart-wrap{{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:22px;margin-bottom:20px}}
.chart-header{{display:flex;align-items:center;gap:16px;margin-bottom:16px;flex-wrap:wrap}}
.chart-title{{font-family:var(--font-head);font-size:15px;font-weight:400;font-style:italic}}
.ctrl-group{{display:flex;gap:6px;margin-left:auto}}
.ctrl-btn{{padding:5px 12px;border-radius:6px;border:1px solid var(--border);background:transparent;color:var(--muted);cursor:pointer;font-family:var(--font-mono);font-size:10px;transition:all .2s}}
.ctrl-btn.active,.ctrl-btn:hover{{background:var(--accent);color:#000;border-color:var(--accent);font-weight:600}}
canvas{{display:block;width:100%;cursor:crosshair}}
.filter-bar{{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:16px}}
.filter-bar select,.filter-bar input{{background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:7px 12px;border-radius:6px;font-family:var(--font-mono);font-size:11px}}
.filter-bar input{{width:200px}}
.table-wrap{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse}}
thead th{{font-size:9px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);padding:10px 12px;text-align:left;border-bottom:1px solid var(--border);cursor:pointer}}
thead th:hover{{color:var(--accent)}}
tbody tr{{border-bottom:1px solid var(--border);transition:background .15s;cursor:pointer}}
tbody tr:hover{{background:var(--surface2)}}
tbody td{{padding:9px 12px;font-size:11px;vertical-align:middle}}
.tag{{display:inline-block;padding:2px 8px;border-radius:4px;font-size:9px;font-weight:600;letter-spacing:.04em}}
.bar-inline{{display:flex;align-items:center;gap:8px}}
.bar-bg{{flex:1;height:3px;background:rgba(30,32,53,.8);border-radius:2px;max-width:80px}}
.bar-fill{{height:3px;border-radius:2px}}
.comp-grid{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}}
.comp-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:20px}}
.comp-card h3{{font-family:var(--font-head);font-size:14px;font-style:italic;margin-bottom:14px}}
.comp-row{{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid var(--dim);font-size:11px}}
.comp-row:last-child{{border-bottom:none}}
.comp-row .label{{color:var(--muted)}}
.comp-row .val{{font-weight:500}}
.mod-bars{{display:flex;flex-direction:column;gap:10px;margin-bottom:20px}}
.mod-bar-row{{display:flex;align-items:center;gap:14px}}
.mod-bar-label{{min-width:220px;font-size:11px;color:var(--text)}}
.mod-bar-track{{flex:1;height:24px;background:var(--surface2);border-radius:4px;position:relative;overflow:hidden}}
.mod-bar-fill{{height:100%;border-radius:4px;display:flex;align-items:center;padding-left:8px;font-size:10px;font-weight:500;color:#000}}
.mod-bar-val{{min-width:70px;font-size:11px;color:var(--muted);text-align:right}}
.modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.8);z-index:200;align-items:center;justify-content:center}}
.modal-overlay.open{{display:flex}}
.modal{{background:var(--surface);border:1px solid var(--border);border-radius:14px;width:min(760px,94vw);max-height:86vh;overflow:hidden;display:flex;flex-direction:column}}
.modal-head{{padding:22px 26px;border-bottom:1px solid var(--border);display:flex;align-items:flex-start;gap:16px;background:linear-gradient(135deg,var(--surface) 0%,var(--surface2) 100%)}}
.modal-head h2{{font-family:var(--font-head);font-size:18px;font-style:italic;margin-bottom:6px}}
.modal-meta{{display:flex;gap:14px;flex-wrap:wrap}}
.modal-kpi{{font-size:11px;color:var(--muted)}}
.modal-kpi span{{color:var(--accent);font-weight:500}}
.modal-close{{width:30px;height:30px;border-radius:50%;border:1px solid var(--border);background:transparent;color:var(--muted);cursor:pointer;font-size:14px;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-left:auto}}
.modal-close:hover{{background:var(--surface2)}}
.modal-body{{padding:20px 24px;overflow-y:auto;flex:1}}
.meto-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.meto-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px}}
.meto-card h3{{font-family:var(--font-head);font-size:13px;font-style:italic;margin-bottom:12px;color:var(--accent)}}
.meto-row{{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border);font-size:11px}}
.meto-row:last-child{{border-bottom:none}}
.meto-row span:last-child{{color:var(--accent);font-weight:500}}
.meto-note{{grid-column:1/-1;background:var(--surface2);border:1px solid var(--dim);border-radius:8px;padding:16px;font-size:11px;line-height:1.8;color:var(--muted)}}
.meto-note strong{{color:var(--text)}}
#tooltip{{position:fixed;display:none;background:var(--surface2);border:1px solid var(--accent);border-radius:8px;padding:12px 14px;font-size:11px;line-height:1.7;pointer-events:none;z-index:999;max-width:310px;box-shadow:0 8px 32px rgba(0,0,0,.6)}}
#tooltip strong{{font-family:var(--font-head);font-size:14px;font-style:italic;color:var(--accent);display:block;margin-bottom:4px}}
@media(max-width:768px){{.comp-grid,.meto-grid{{grid-template-columns:1fr}}}}
</style>
</head>
<body>

<div class="header">
  <div class="header-icon">
    <svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2">
      <circle cx="12" cy="12" r="3"/><circle cx="12" cy="12" r="8" stroke-dasharray="4 2"/>
      <line x1="12" y1="2" x2="12" y2="5"/><line x1="12" y1="19" x2="12" y2="22"/>
      <line x1="2" y1="12" x2="5" y2="12"/><line x1="19" y1="12" x2="22" y2="12"/>
    </svg>
  </div>
  <div class="header-text">
    <h1>Metodologia de Decisão de Investimento — FSA Cinema</h1>
    <p>Impacto da metodologia de seleção nos resultados das obras · obras com ano de produção ≤ {ANO_MAX}</p>
  </div>
  <div class="header-meta">
    {fmt_int(summary.get("totalUniverse", 0))} obras · {fmt_int(summary.get("calculable", 0))} com ROI calculável<br>
    ANCINE · FSA/BRDE · SALIC/MinC · Lumière/CNC
  </div>
</div>

<div class="tabs">
  <div class="tab active" onclick="showTab('tab-comp',this)">Comparação</div>
  <div class="tab" onclick="showTab('tab-rank',this)">Chamadas Detalhadas</div>
  <div class="tab" onclick="showTab('tab-obras',this)">Obras Detalhadas</div>
  <div class="tab" onclick="showTab('tab-meto',this)">Metodologia</div>
  <div class="tab" onclick="showTab('tab-sint',this)">Síntese</div>
</div>

<!-- ═══ TAB 1: COMPARAÇÃO ══════════════════════════════════════════════════ -->
<div id="tab-comp" class="panel active">
  <div class="kpi-bar" id="kpi-bar"></div>

  <!-- Quadrante Comercial × Internacional -->
  <div class="chart-wrap" style="margin-bottom:20px">
    <div class="chart-header">
      <div>
        <div class="chart-title">Vocação Comercial vs. Alcance Internacional por Categoria</div>
        <div style="font-size:10px;color:var(--muted);margin-top:3px">
          Eixo X: ROI doméstico deflacionado (R$2024) ÷ FSA deflacionado · Eixo Y: ROI Internacional médio (0–100) · Tamanho: nº de obras
        </div>
      </div>
      <div class="ctrl-group" id="quad-axis-btns">
        <button class="ctrl-btn active" onclick="setQuadAxis('roi_tot_def','intl_avg',this)">ROI Total Deflac. × Intl.</button>
        <button class="ctrl-btn" onclick="setQuadAxis('inv','roi_tot_def',this)">Inv. Total × ROI Dom.</button>
        <button class="ctrl-btn" onclick="setQuadAxis('inv','intl_avg',this)">Inv. Total × ROI Intl.</button>
      </div>
    </div>
    <canvas id="canvas-quadrant" height="440"></canvas>
  </div>

  <!-- Ranking por Categoria — full width -->
  <div class="chart-wrap" style="margin-bottom:20px">
    <div class="chart-header">
      <div class="chart-title">Ranking por Categoria</div>
      <div class="ctrl-group" id="rank-sort-btns" style="flex-wrap:wrap">
        <button class="ctrl-btn" onclick="setRankSort('roi_fsa_def',this)">ROI FSA Deflac.</button>
        <button class="ctrl-btn active" onclick="setRankSort('roi_tot_def',this)">ROI Total Deflac.</button>
        <button class="ctrl-btn" onclick="setRankSort('roi_fsa',this)">ROI FSA Nominal</button>
        <button class="ctrl-btn" onclick="setRankSort('roi_tot',this)">ROI Total Nominal</button>
        <button class="ctrl-btn" onclick="setRankSort('intl_avg',this)">ROI Intl.</button>
        <button class="ctrl-btn" onclick="setRankSort('bilh_total',this)">Bilheteria Total</button>
        <button class="ctrl-btn" onclick="setRankSort('bilh_avg',this)">Bilheteria Média</button>
        <button class="ctrl-btn" onclick="setRankSort('bilh_prop',this)">Bilh. Prop.</button>
        <button class="ctrl-btn" onclick="setRankSort('paises_avg',this)">Países VOD Média</button>
        <button class="ctrl-btn" onclick="setRankSort('paises_total',this)">Países VOD Total</button>
        <button class="ctrl-btn" onclick="setRankSort('total_festivais',this)">Festivais</button>
        <button class="ctrl-btn" onclick="setRankSort('jan_avg',this)">Outras Janelas</button>
        <button class="ctrl-btn" onclick="setRankSort('n_com_dados',this)">Com Dados</button>
        <button class="ctrl-btn" onclick="setRankSort('n_sem_dados',this)">Sem Dados</button>
        <button class="ctrl-btn" onclick="setRankSort('n_obras',this)">N Obras</button>
        <button class="ctrl-btn" onclick="setRankSort('inv',this)">Investimento</button>
      </div>
    </div>
    <div style="font-size:10px;color:var(--muted);margin-bottom:10px" id="rank-sort-label">ROI proporcional: parcela da receita atribuída à categoria ÷ investimento total</div>
    <div id="rank-dom" class="mod-bars"></div>
  </div>

  <!-- Cards de categoria -->
  <div class="chart-wrap" style="margin-bottom:20px">
    <div class="chart-header">
      <div>
        <div class="chart-title">Categorias de Fomento — Critérios e Resultados</div>
        <div style="font-size:10px;color:var(--muted);margin-top:3px">Mecanismo de seleção, retorno médio e alcance por categoria · clique para ver obras</div>
      </div>
    </div>
    <div id="cat-desc-grid" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:14px;margin-top:10px;align-items:start"></div>
  </div>

  <div class="chart-wrap">
    <div class="chart-header">
      <div>
        <div class="chart-title">Valor Investido por Ano e Categoria</div>
        <div style="font-size:10px;color:var(--muted);margin-top:3px">Total FSA/renúncia captado por ano de produção da obra (R$)</div>
      </div>
    </div>
    <canvas id="canvas-timeline" height="300"></canvas>
    <div id="timeline-legend" style="display:flex;flex-wrap:wrap;gap:14px;margin-top:14px;font-size:10px;line-height:1.6"></div>
  </div>
</div>



<!-- ═══ TAB 3: RANKING ════════════════════════════════════════════════════ -->
<div id="tab-rank" class="panel">
  <div class="filter-bar">
    <input type="text" id="search-rank" placeholder="Buscar chamada…" oninput="renderTable()">
    <select id="filter-fase1" onchange="renderTable()">
      <option value="">Todos os critérios</option>
      {rank_options}
    </select>
    <select id="sort-rank" onchange="renderTable()">
      <option value="roi_fsa_def">ROI FSA Deflac. ↓</option>
      <option value="roi_tot_def" selected>ROI Total Deflac. ↓</option>
      <option value="roi_fsa">ROI FSA Nominal ↓</option>
      <option value="roi_tot">ROI Total Nominal ↓</option>
      <option value="rim">ROI Intl Médio ↓</option>
      <option value="rimax">ROI Intl Máx. ↓</option>
      <option value="n">N° Obras ↓</option>
      <option value="inv">Investimento ↓</option>
      <option value="bilh_avg">Bilheteria Média ↓</option>
      <option value="paises_avg">Países VOD ↓</option>
    </select>
    <div class="cat-pills" id="cat-pills-rank"></div>
    <span id="rank-count" style="font-size:10px;color:var(--muted);margin-left:auto"></span>
  </div>
  <div class="table-wrap" id="rank-table-wrap"></div>
</div>

<!-- ═══ TAB 4: OBRAS DETALHADAS ═══════════════════════════════════════════ -->
<div id="tab-obras" class="panel">
  <div class="filter-bar" style="margin-bottom:16px">
    <input type="text" id="search-obras" placeholder="Buscar por título…" oninput="renderObras()">
    <select id="filter-obras-cat" onchange="renderObras()">
      <option value="">Todas as categorias</option>
    </select>
    <select id="sort-obras" onchange="renderObras()">
      <option value="titulo">Título A–Z</option>
      <option value="ano">Ano ↓</option>
      <option value="inv">Investimento ↓</option>
      <option value="roi_dom">ROI Dom. ↓</option>
      <option value="roi_intl" selected>ROI Intl. ↓</option>
      <option value="bilh">Bilheteria ↓</option>
      <option value="n_paises">Países VOD ↓</option>
      <option value="bilh">Bilheteria ↓</option>
      <option value="n_paises">Países VOD ↓</option>
      <option value="adm">Lumière (EU) ↓</option>
      <option value="n_festivais">Qtd Festivais ↓</option>
    </select>
    <span id="obras-count" style="font-size:10px;color:var(--muted);margin-left:auto"></span>
  </div>
  <div class="table-wrap" id="obras-table-wrap"></div>
  <p style="font-size:10px;color:var(--muted);margin-top:14px;line-height:1.7">
    <strong>Nota:</strong> uma mesma obra pode aparecer em múltiplas linhas quando recebeu
    chamadas FSA de categorias distintas e/ou ambos os mecanismos de renúncia fiscal.
    O investimento de cada linha é o valor da chamada específica; não somar para totais globais.
  </p>
</div>

<!-- ═══ TAB 5: METODOLOGIA ════════════════════════════════════════════════ -->
<div id="tab-meto" class="panel">
  <div style="max-width:900px;margin:0 auto">

    <div class="meto-card" style="margin-bottom:16px">
      <h3>Pergunta analítica</h3>
      <p style="font-size:12px;color:var(--muted);line-height:1.8;margin-bottom:0">
        Editais FSA que usaram <strong style="color:var(--text)">histórico de festivais internacionais</strong>
        como critério de seleção produziram obras com maior alcance internacional do que editais que usaram
        <strong style="color:var(--text)">critérios comerciais</strong> (bilheteria, market share)?
        E como os mecanismos <strong style="color:var(--text)">automáticos e de renúncia fiscal</strong>
        se comparam a ambos em retorno doméstico?
      </p>
    </div>

    <div class="meto-grid" style="margin-bottom:16px">
      <div class="meto-card">
        <h3>Mecanismos de Investimento Direto FSA</h3>
        <p style="font-size:11px;color:var(--muted);line-height:1.7;margin-bottom:10px">
          O FSA (Fundo Setorial do Audiovisual) financia diretamente via editais do BRDE.
          Editais <strong style="color:var(--text)">seletivos</strong> usam duas fases:
          Fase 1 pontua o histórico da empresa, Fase 2 avalia o roteiro.
          Editais <strong style="color:var(--text)">automáticos</strong> usam critérios objetivos de desempenho anterior.
        </p>
        <div class="meto-row"><span style="color:#6c7bf7">FSA Pontuação Festivais e Roteiro</span><span>PRODECINE 03–05 · Concurso Mód. B</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Fase 1: pontuação por premiações/seleções em festivais internacionais (Oscar, Cannes, Berlim, Veneza…). Fase 2: leitura de roteiro. Inscrição pela produtora. Editais seletivos com histórico artístico.</span></div>
        <div class="meto-row"><span style="color:#f5c842">FSA Pontuação Bilheteria e Roteiro — Produtora</span><span>PRODECINE 01 · 06 · Produção Cinema</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Fase 1: desempenho comercial — bilheteria e market share das obras anteriores da produtora. Fase 2: leitura de roteiro. Editais seletivos com histórico comercial.</span></div>
        <div class="meto-row"><span style="color:#f09020">FSA Pontuação Bilheteria e Roteiro — Distribuidora</span><span>PRODECINE 02 · Cinema via Dist.</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Mesma lógica comercial, mas a candidatura é feita pela distribuidora, que assume a responsabilidade pela comercialização do projeto.</span></div>
        <div class="meto-row"><span style="color:#5b8cff">FSA Automático Bilheteria</span><span>Fluxo Contínuo · SUAT Comercial</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">100% automático: produtora ou distribuidora recebe benefício com base no resultado de bilheteria da obra anterior — sem edital competitivo. Inclui Fluxo Contínuo Produção Cinema, Fluxo Contínuo via Distribuidora e Suporte Automático Comercial.</span></div>
        <div class="meto-row"><span style="color:#c080ff">FSA Automático Festivais</span><span>Suporte Auto. Artístico (SUAT)</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Automático por desempenho artístico: FSA recompensa obras reconhecidas em festivais nacionais e internacionais (Brasília, Gramado, Fest.Rio, e festivais internacionais de prestígio).</span></div>
        <div class="meto-row"><span style="color:#ff80b0">FSA Coprodução Internacional</span><span>PRODECINE 07–12 · acordos bilaterais</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Exige parceria formal com produtora estrangeira. Inclui acordos bilaterais Brasil–Chile, Brasil–Portugal, Brasil–Uruguai e editais gerais de coprodução.</span></div>
        <div class="meto-row"><span style="color:#40c8ff">FSA SAV/MINC / Arranjos Regionais</span><span>SAV/MINC 01–13 · Arranjos</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Editais temáticos e regionais coordenados pelo Ministério da Cultura. Inclui políticas de descentralização regional e fomento diversificado.</span></div>
        <div class="meto-row"><span style="color:#e060e0">FSA Comercialização / Distribuição</span><span>COMERCIALIZAÇÃO EM CINEMA · COMPLEMENTAÇÃO · Opção Comercialização · Suporte Auto. Desempenho Comercial Cinema</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Investimento direto em lançamento e distribuição de obras já produzidas. Inclui chamadas exclusivamente voltadas à etapa de comercialização — não de produção.</span></div>
      </div>

      <div class="meto-card">
        <h3>Renúncia Fiscal — SALIC</h3>
        <p style="font-size:11px;color:var(--muted);line-height:1.7;margin-bottom:10px">
          A renúncia fiscal não é investimento direto: o Estado abre mão de receita tributária
          permitindo que empresas (patrocinadores) apliquem impostos devidos em projetos audiovisuais.
          Diferente do FSA, <strong style="color:var(--text)">não há seleção ANCINE/BRDE</strong>
          — a aprovação é administrativa (registro ANCINE + captação no mercado).
        </p>
        <div class="meto-row"><span style="color:#ff8040">Renúncia — Art. 3 / 3-A / 39</span><span>Lei do Audiovisual</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Art. 3: renúncia de IR/CSLL de empresas sobre investimentos em obras audiovisuais brasileiras — mecanismo principal do setor. Art. 3-A: complementar para coproduções. Art. 39: mecanismo da Lei 12.761/2012.</span></div>

        <h3 style="margin-top:16px">Métricas de Retorno</h3>
        <div class="meto-row"><span>ROI Total Deflacionado (R$2024) — <strong style="color:var(--text)">métrica primária</strong></span><span>receita deflac. / (FSA deflac. + renúncia deflac.)</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Receita = bilheteria deflacionada + outras janelas deflacionadas (R$2024). <strong>Nota:</strong> as janelas (TV Paga, VOD, TV Aberta, DVD) são estimativas com valores fixos por tier, deflacionadas pela data de emissão do CRT — não são receita observada (dados sintéticos). Denominador = capital público total alocado na obra (FSA direto + renúncia fiscal Art. 3/3-A). Mede o retorno sobre o conjunto do fomento público, tratando FSA e renúncia como co-investimento do Estado.</span></div>
        <div class="meto-row"><span>ROI FSA Deflacionado (R$2024)</span><span>receita deflac. / FSA deflac.</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Considera apenas o investimento FSA direto no denominador, excluindo renúncia fiscal. Útil para isolar a eficiência do FSA como instrumento — mas subestima o capital público total mobilizado em obras com forte participação de renúncia.</span></div>
        <div class="meto-row"><span>ROI Doméstico Nominal (compatibilidade)</span><span>receita nominal total da obra / investimento nominal deste mecanismo</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Versão nominal (não deflacionada), mantida para comparação. Distorcido por inflação — obras mais antigas parecem ter ROI menor. Use o deflacionado para análise de política.</span></div>
        <div class="meto-row"><span>ROI Doméstico Mediano por categoria</span><span>Mediana dos ROIs individuais</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Mediana dos ROIs (receita total da obra ÷ investimento do mecanismo) de cada obra na categoria. Menos sensível a outliers do que a média — uma obra com ROI 50x não distorce o valor típico da categoria.</span></div>
        <div class="meto-row"><span>ROI Doméstico Agregado por categoria</span><span>Proporcional ao financiamento</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Para cada obra com múltiplos mecanismos, a receita é repartida proporcionalmente ao investimento de cada um. O agregado da categoria soma receitas proporcionais / investimento total. Mede quanto a categoria "merece" do retorno dado sua participação no financiamento das obras — evita dupla contagem ao comparar categorias.</span></div>
        <div class="meto-row"><span>Deduplicação por obra×categoria</span><span>uma obra, uma contagem</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Obras com múltiplas chamadas mapeadas para a mesma categoria aparecem uma única vez no cômputo de bilheteria, janelas, festivais e VOD da categoria. O ROI usa todas as entradas (cada chamada tem seu próprio investimento).</span></div>
        <div class="meto-row"><span>ROI Internacional</span><span>Score composto 0–100</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">70% score de festivais internacionais (Oscar, Cannes, Berlim, Veneza, BAFTA, Globo de Ouro) + 20% admissões na União Europeia (base Lumière/CNC) + 10% presença em plataformas VOD internacionais.</span></div>
        <div class="meto-row"><span>Países alcançados (total)</span><span>soma de países VOD por obra</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Soma das entradas de países com distribuição VOD internacional em todas as obras da categoria (base Lumière). Mede o alcance geográfico acumulado — não países únicos da categoria, mas contatos por obra.</span></div>
        <div class="meto-row"><span>Filtragem TV/VOD</span><span>Apenas obras de cinema</span></div>
        <div class="meto-row"><span style="font-size:10px;color:var(--muted)">Chamadas PRODAV, Fluxo Contínuo TV e demais editais de televisão são excluídos. Renúncia só é atribuída a obras com ao menos uma chamada FSA cinema confirmada.</span></div>

        <h3 style="margin-top:16px">Fontes de Dados</h3>
        <div class="meto-row"><span>Universo de obras</span><span>obras_fomento_unificado.csv · ANCINE</span></div>
        <div class="meto-row"><span>Projetos FSA</span><span>projetos-fsa.csv · BRDE/FSA</span></div>
        <div class="meto-row"><span>Renúncia fiscal</span><span>projetos-com-renuncia-fiscal.csv · SALIC/MinC</span></div>
        <div class="meto-row"><span>Bilheteria</span><span>por_filme_ano.csv · ANCINE (2014+)</span></div>
        <div class="meto-row"><span>Admissões EU</span><span>lumiere_search.xlsx · Lumière/CNC</span></div>
        <div class="meto-row"><span>VOD internacional</span><span>lumiere_vod_search.xlsx · Lumière/CNC</span></div>
        <div class="meto-row"><span>Festivais</span><span>festivais_consolidado.csv · Atas BRDE/FSA consolidadas (2014–2024) + pesquisa bibliográfica (complemento)</span></div>
      </div>
    </div>

    <div class="meto-note">
      <strong>Limitações metodológicas:</strong>
      (1) Obras com ano de produção &gt; {ANO_MAX} são excluídas — ciclo de vida incompleto.
      (2) Bilheteria pré-2014 estimada por título (menor precisão).
      (3) Festivais cobre apenas as principais premiações internacionais — participações em festivais menores não são contabilizadas.
      (4) Renúncia fiscal exclui obras sem FSA confirmado para evitar superestimação do ROI (denominador seria apenas o captado SALIC, que é parcial).
      (5) Categorias com menos de 5 obras são excluídas do painel por insuficiência amostral.
      (6) Obras com múltiplas chamadas na mesma categoria são deduplicadas para métricas de obra-nível (bilheteria, janelas, festivais, VOD). O ROI usa todas as entradas para refletir o investimento real de cada chamada.
      (7) <strong>Causalidade vs auto-seleção:</strong> produtores comerciais tendem a buscar chamadas com critério de bilheteria, enquanto cineastas autorais buscam chamadas com critério festival. A correlação entre critério e tipo de retorno pode refletir tanto o efeito do critério quanto a auto-seleção dos candidatos. Os dados não permitem distinguir os dois mecanismos.
    </div>
  </div>
</div>

<!-- ═══ TAB 5: SÍNTESE ════════════════════════════════════════════════════ -->
<div id="tab-sint" class="panel">
  <div style="max-width:900px;margin:0 auto;padding-top:16px">
    <div class="meto-card" style="margin-bottom:16px">
      <h3>Qual metodologia é melhor para cada métrica?</h3>
      <p style="font-size:12px;color:var(--muted);line-height:1.8;margin-bottom:0">
        A tabela abaixo responde a pergunta central: <strong style="color:var(--text)">dada uma meta de política pública
        (retorno doméstico, alcance internacional, presença em festivais…), qual metodologia FSA deve ser priorizada?</strong>
        Análise exclui renúncia fiscal isolada — ela é tratada como co-investimento de mercado, não como instrumento gerenciado.
        Apenas categorias com pelo menos 10 obras incluídas.
      </p>
    </div>
    <div id="sint-table"></div>
  </div>
</div>

<!-- MODAL -->
<div class="modal-overlay" id="modal-overlay" onclick="closeModal(event)">
  <div class="modal">
    <div class="modal-head">
      <div style="flex:1">
        <h2 id="modal-title"></h2>
        <div class="modal-meta" id="modal-meta"></div>
      </div>
      <button class="modal-close" onclick="closeModal()">✕</button>
    </div>
    <div class="modal-body">
      <div id="modal-content"></div>
    </div>
  </div>
</div>

<div id="tooltip"></div>

<script>
const CATS = {safe_json(cats_js)};
const CHM  = {safe_json(chamadas_js)};
const EXCL = {safe_json(excl_js)};
const OBRAS_DET = {safe_json(obras_det_js)};
const SINTESE = {safe_json(sintese_js)};
const TIMELINE = {safe_json(timeline_js)};

// ── Utilitários ────────────────────────────────────────────────────────────
function fmtRatio(v){{return (v||0).toLocaleString('pt-BR',{{minimumFractionDigits:3,maximumFractionDigits:3}})+'x'}}
function fmtMoney(v){{
  if(v>=1e9) return 'R$ '+(v/1e9).toFixed(1)+'B';
  if(v>=1e6) return 'R$ '+(v/1e6).toFixed(1)+'M';
  return 'R$ '+Math.round(v||0).toLocaleString('pt-BR');
}}
function fmtInt(v){{return Math.round(v||0).toLocaleString('pt-BR')}}
function showTab(id,el){{
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  if(el) el.classList.add('active');
  if(id==='tab-comp')    setTimeout(()=>{{drawQuadrant();drawRankings();drawTimeline();}},50);
  if(id==='tab-obras')   {{buildCatOptions();renderObras();}}
  if(id==='tab-sint')    renderSintese();
}}

// ── KPI bar ────────────────────────────────────────────────────────────────
function buildKpis(){{
  const fest =CATS.find(c=>c.key==='pont_fest_prod')||{{}};
  const com  =CATS.find(c=>c.key==='pont_com_prod') ||{{}};
  const comd =CATS.find(c=>c.key==='pont_com_dist') ||{{}};
  const com_ref=Math.max(com.intl_avg||0,comd.intl_avg||0);
  const dN=fest.intl_avg&&com_ref?((fest.intl_avg-com_ref)/Math.max(com_ref,.001)*100):null;
  const delta=dN!==null?dN.toFixed(0):'—';
  document.getElementById('kpi-bar').innerHTML=`
    <div class="kpi"><div class="kpi-label">Festivais — Produtora</div><div class="kpi-val" style="color:#6c7bf7">${{fmtInt(fest.n_obras||0)}}</div><div class="kpi-sub">ROI intl médio ${{(fest.intl_avg||0).toFixed(1)}}</div></div>
    <div class="kpi"><div class="kpi-label">Comercial — Produtora</div><div class="kpi-val" style="color:#f5c842">${{fmtInt(com.n_obras||0)}}</div><div class="kpi-sub">ROI intl médio ${{(com.intl_avg||0).toFixed(1)}}</div></div>
    <div class="kpi"><div class="kpi-label">Comercial — Distribuidora</div><div class="kpi-val" style="color:#e8b030">${{fmtInt(comd.n_obras||0)}}</div><div class="kpi-sub">ROI intl médio ${{(comd.intl_avg||0).toFixed(1)}}</div></div>
    <div class="kpi"><div class="kpi-label">Δ ROI Intl Festivais vs Comercial</div><div class="kpi-val" style="color:${{dN!==null?(dN>0?'var(--accent)':'var(--coral)'):'var(--muted)'}}">${{delta!=='—'?(Number(delta)>0?'+':'')+delta+'%':delta}}</div><div class="kpi-sub">Festivais-Prod. vs max(Com-Prod.,Com-Dist.)</div></div>`;
}}

// ── Quadrante Comercial × Internacional ─────────────────────────────────────
let _quadX='roi_tot_def', _quadY='intl_avg';
const _quadXLabels={{'roi_tot_def':'ROI Total Deflac. R$2024 (x)','inv':'Investimento Total (R$)'}};
const _quadYLabels={{'intl_avg':'Score Internacional Médio (0–100)','roi_tot_def':'ROI Total Deflac. R$2024 (x)'}};
function setQuadAxis(xKey,yKey,btn){{
  _quadX=xKey; _quadY=yKey;
  document.querySelectorAll('#quad-axis-btns .ctrl-btn').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  drawQuadrant();
}}
function drawQuadrant(){{
  const cv=document.getElementById('canvas-quadrant'); if(!cv) return;
  const W=cv.offsetWidth||cv.parentElement.clientWidth||900;
  cv.width=W; cv.height=parseInt(cv.getAttribute('height')||440);
  const H=cv.height;
  const ctx=cv.getContext('2d');
  ctx.clearRect(0,0,W,H);

  const PAD={{l:72,r:30,t:20,b:58}};
  const PW=W-PAD.l-PAD.r, PH=H-PAD.t-PAD.b;

  const cats=CATS.filter(c=>c.n_obras>=1);
  if(!cats.length) return;

  const xs=cats.map(c=>c[_quadX]||0);
  const ys=cats.map(c=>c[_quadY]||0);
  const maxX=Math.max(...xs)*1.12||1;
  const maxY=Math.max(...ys)*1.15||1;
  const maxN=Math.max(...cats.map(c=>c.n_obras))||1;

  // medians for quadrant lines
  const sortedX=[...xs].sort((a,b)=>a-b);
  const sortedY=[...ys].sort((a,b)=>a-b);
  const medX=sortedX[Math.floor(sortedX.length/2)]||maxX/2;
  const medY=sortedY[Math.floor(sortedY.length/2)]||maxY/2;

  function toCanvas(xv,yv){{
    return [PAD.l+xv/maxX*PW, PAD.t+PH-(yv/maxY*PH)];
  }}

  const [qx]=toCanvas(medX,0);
  const [,qy]=toCanvas(0,medY);

  // quadrant backgrounds
  const quads=[
    {{x:PAD.l, y:PAD.t, w:qx-PAD.l, h:qy-PAD.t, color:'rgba(160,64,232,.07)', label:'Voca\u00e7\u00e3o|Internacional'}},
    {{x:qx,    y:PAD.t, w:PAD.l+PW-qx, h:qy-PAD.t, color:'rgba(108,123,247,.07)',  label:'Duplo|Impacto'}},
    {{x:PAD.l, y:qy,    w:qx-PAD.l, h:PAD.t+PH-qy, color:'rgba(90,96,128,.05)', label:'Baixo Retorno|Detectado'}},
    {{x:qx,    y:qy,    w:PAD.l+PW-qx, h:PAD.t+PH-qy, color:'rgba(245,200,66,.07)', label:'Voca\u00e7\u00e3o|Comercial'}},
  ];
  quads.forEach(q=>{{
    ctx.fillStyle=q.color;
    ctx.fillRect(q.x,q.y,q.w,q.h);
  }});

  // quadrant border lines
  ctx.strokeStyle='rgba(94,100,140,.35)';
  ctx.setLineDash([5,5]);
  ctx.lineWidth=1;
  ctx.beginPath(); ctx.moveTo(qx,PAD.t); ctx.lineTo(qx,PAD.t+PH); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(PAD.l,qy); ctx.lineTo(PAD.l+PW,qy); ctx.stroke();
  ctx.setLineDash([]);

  // quadrant labels
  const _qfont='10px DM Mono, monospace';
  ctx.font=_qfont; ctx.textAlign='center';
  quads.forEach(q=>{{
    const lx=q.x+q.w/2, ly=q.y+q.h/2;
    ctx.fillStyle='rgba(200,210,240,.18)';
    const parts=q.label.split('|');
    parts.forEach((p,i)=>ctx.fillText(p,lx,ly+(i-(parts.length-1)/2)*14));
  }});

  // grid lines + axes
  ctx.strokeStyle='rgba(30,32,53,.9)'; ctx.lineWidth=1;
  ctx.strokeRect(PAD.l,PAD.t,PW,PH);

  // X axis labels
  ctx.fillStyle='rgba(90,96,128,.9)'; ctx.font='10px DM Mono,monospace'; ctx.textAlign='center';
  const nTicks=5;
  for(let i=0;i<=nTicks;i++){{
    const v=maxX*i/nTicks;
    const [cx_]=toCanvas(v,0);
    ctx.fillStyle='rgba(30,32,53,.6)'; ctx.fillRect(cx_,PAD.t,1,PH);
    ctx.fillStyle='rgba(90,96,128,.9)';
    const lbl=_quadX==='inv'?'R$'+(v/1e6).toFixed(0)+'M':v.toFixed(2)+'x';
    ctx.fillText(lbl,cx_,PAD.t+PH+14);
  }}
  // Y axis labels
  ctx.textAlign='right';
  for(let i=0;i<=4;i++){{
    const v=maxY*i/4;
    const [,cy_]=toCanvas(0,v);
    ctx.fillStyle='rgba(30,32,53,.6)'; ctx.fillRect(PAD.l,cy_,PW,1);
    ctx.fillStyle='rgba(90,96,128,.9)';
    ctx.fillText(v.toFixed(1),PAD.l-6,cy_+4);
  }}

  // axis titles
  ctx.save(); ctx.translate(14,PAD.t+PH/2); ctx.rotate(-Math.PI/2);
  ctx.fillStyle='rgba(90,96,128,.9)'; ctx.font='10px DM Mono,monospace'; ctx.textAlign='center';
  ctx.fillText(_quadYLabels[_quadY]||_quadY,0,0); ctx.restore();
  ctx.textAlign='center'; ctx.fillStyle='rgba(90,96,128,.9)';
  ctx.fillText(_quadXLabels[_quadX]||_quadX,PAD.l+PW/2,H-6);

  // bubbles + labels
  const rendered=[];
  cats.forEach(c=>{{
    const xv=c[_quadX]||0, yv=c[_quadY]||0;
    const [cx2,cy2]=toCanvas(xv,yv);
    const r=Math.max(7,Math.min(32,6+Math.sqrt(c.n_obras/maxN)*28));
    rendered.push({{c,cx2,cy2,r}});

    // bubble
    ctx.beginPath(); ctx.arc(cx2,cy2,r,0,Math.PI*2);
    ctx.fillStyle=c.cor+'33'; ctx.fill();
    ctx.strokeStyle=c.cor; ctx.lineWidth=2; ctx.stroke();

    // short label — strip 'FSA ' prefix for space
        const short=c.label.replace(/^FSA /,'').replace(/—\s?/g,'|').replace(/ \(SUAT\)/,'');
    const parts=short.split('|').map(s=>s.trim()).filter(Boolean);
    ctx.font='bold 9px DM Mono,monospace'; ctx.textAlign='center'; ctx.fillStyle=c.cor;
    const ly_base=cy2+r+12;
    parts.forEach((p,i)=>ctx.fillText(p,cx2,ly_base+i*11));
  }});

  // store for hover
  cv._rendered=rendered;
}}

// Quadrant hover
(function(){{
  const cv=document.getElementById('canvas-quadrant'); if(!cv) return;
  const tip=document.getElementById('tooltip');
  cv.addEventListener('mousemove',e=>{{
    if(!cv._rendered) return;
    const rect=cv.getBoundingClientRect();
    const mx=e.clientX-rect.left, my=e.clientY-rect.top;
    let hit=null;
    for(const p of cv._rendered){{
      const dx=mx-p.cx2, dy=my-p.cy2;
      if(Math.sqrt(dx*dx+dy*dy)<=p.r+4){{hit=p;break;}}
    }}
    if(hit){{
      const c=hit.c;
      const xLabel=_quadXLabels[_quadX]||_quadX;
      const yLabel=_quadYLabels[_quadY]||_quadY;
      const xVal=_quadX==='inv'?fmtMoney(c[_quadX]||0):fmtRatio(c[_quadX]||0);
      const yVal=_quadY==='intl_avg'?(c.intl_avg||0).toFixed(1):fmtRatio(c[_quadY]||0);
      tip.style.display='block';
      tip.style.left=(e.clientX+16)+'px';
      tip.style.top=(e.clientY-10)+'px';
      tip.innerHTML=`<strong style="color:${{c.cor}}">${{c.label}}</strong>
${{xLabel}}: <span style="color:${{c.cor}}">${{xVal}}</span><br>
${{yLabel}}: <span style="color:var(--accent)">${{yVal}}</span><br>
Festivais (obras c/ dados): ${{c.n_com_festivais}} · ${{c.total_festivais}} seleções<br>
Bilheteria média: ${{fmtMoney(c.bilh_avg||0)}}<br>
Nº obras: ${{c.n_obras}}`;
      cv.style.cursor='default';
    }} else {{
      tip.style.display='none';
      cv.style.cursor='crosshair';
    }}
  }});
  cv.addEventListener('mouseleave',()=>{{ document.getElementById('tooltip').style.display='none'; }});
}})();

// ── Rankings em HTML (legíveis) ────────────────────────────────────────────
const RANK_LABELS={{"roi_tot_def":"ROI total deflacionado (R$2024): receita deflac. ÷ (FSA deflac. + renúncia deflac.) — métrica primária (capital público total)","roi_fsa_def":"ROI FSA deflacionado (R$2024): receita deflac. total da obra ÷ FSA deflac. — exclui renúncia fiscal do denominador","roi_fsa":"ROI nominal FSA (receita total da obra ÷ total FSA da obra): mediana por categoria ou chamada","roi_tot":"ROI nominal total (receita total da obra ÷ total investimento público da obra): mediana por categoria ou chamada","rda":"ROI proporcional: parcela da receita atribuída à categoria (proporcional ao investimento) ÷ investimento total","rda_avg":"ROI médio por obra: média de (receita total da obra ÷ investimento deste mecanismo) — mede alavancagem típica por R$ investido","intl_avg":"ROI Internacional: score composto 0–100 (70% festivais intl + 20% Lumière/CNC + 10% VOD países)","n_obras":"Número de obras na categoria com dados calculáveis","inv":"Total investido pelo mecanismo FSA/renúncia (R$)","bilh_total":"Bilheteria nominal total das obras da categoria (R$)","bilh_avg":"Bilheteria nominal média por obra na categoria (R$)","paises_avg":"Média de países com VOD internacional (obras com dados)","total_festivais":"Total de seleções em festivais somando todas as obras da categoria","bilh_prop":"Receita proporcional ao investimento da categoria (bilheteria + outras janelas, R$)","jan_avg":"Receita média de outras janelas (TV, streaming, VOD) por obra com dados (R$)","n_com_dados":"Obras com ao menos uma fonte de dados (bilheteria, janelas, festivais ou VOD)","n_sem_dados":"Obras sem nenhuma fonte de dados encontrada","paises_total":"Soma de entradas VOD por país em todas as obras da categoria (contagem de contatos, não países únicos)"}};
let _rankSortKey='roi_tot_def';
function setRankSort(key,btn){{
  _rankSortKey=key;
  document.querySelectorAll('#rank-sort-btns .ctrl-btn').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  const lbl=document.getElementById('rank-sort-label');
  if(lbl) lbl.textContent=RANK_LABELS[key]||'';
  drawRankings();
}}
function drawRankings(){{
  const cats=CATS.filter(c=>c.n_obras>0);

  function renderRank(containerId, valueKey, fmtFn){{
    const sorted=[...cats].filter(c=>c[valueKey]>0).sort((a,b)=>b[valueKey]-a[valueKey]);
    const maxV=sorted[0]?sorted[0][valueKey]:1;
    document.getElementById(containerId).innerHTML=sorted.map(c=>{{
      const pct=Math.max(4,(c[valueKey]/maxV*100)).toFixed(1);
      return `<div class="mod-bar-row" style="margin-bottom:6px">
        <div class="mod-bar-label" style="color:${{c.cor}};flex:0 0 210px;width:210px;white-space:normal;line-height:1.3;font-size:10px">${{c.label}}</div>
        <div class="mod-bar-track" style="flex:1;margin-left:10px">
          <div class="mod-bar-fill" style="width:${{pct}}%;background:${{c.cor}};color:#000;font-size:10px;min-width:2px">&nbsp;</div>
        </div>
        <div class="mod-bar-val" style="min-width:56px;text-align:right;font-size:11px;color:${{c.cor}}">${{fmtFn(c[valueKey])}}</div>
      </div>`;
    }}).join('');
  }}

  const _fmtForKey=k=>k==='n_obras'||k==='n_com_dados'||k==='n_sem_dados'||k==='total_festivais'||k==='paises_total'?v=>fmtInt(v):k==='inv'||k==='bilh_total'||k==='bilh_avg'||k==='bilh_prop'||k==='jan_avg'?v=>fmtMoney(v):k==='intl_avg'||k==='paises_avg'?v=>v.toFixed(1):v=>fmtRatio(v);
  renderRank('rank-dom', _rankSortKey, _fmtForKey(_rankSortKey));
}}

// ── Timeline — gráfico de linhas ───────────────────────────────────────────
function drawTimeline(){{
  const cv=document.getElementById('canvas-timeline');
  if(!cv||!TIMELINE||!TIMELINE.anos.length) return;
  const ctx=cv.getContext('2d');
  const dpr=window.devicePixelRatio||1;
  cv.width=cv.offsetWidth*dpr; cv.height=300*dpr;
  ctx.scale(dpr,dpr);
  const W=cv.offsetWidth, H=300;
  ctx.clearRect(0,0,W,H);
  const pad={{l:72,r:20,t:20,b:36}};
  const anos=TIMELINE.anos;
  const series=TIMELINE.series.filter(s=>s.values.some(v=>v>0));
  if(!series.length||!anos.length) return;

  const maxVal=Math.max(...series.flatMap(s=>s.values),1);
  const xOf=i=>pad.l+i*(W-pad.l-pad.r)/(Math.max(anos.length-1,1));
  const yOf=v=>pad.t+(1-v/maxVal)*(H-pad.t-pad.b);

  // Y grid
  const nTicks=5;
  ctx.strokeStyle='rgba(30,32,53,.9)';ctx.lineWidth=1;
  for(let t=0;t<=nTicks;t++){{
    const v=maxVal*t/nTicks;
    const yy=yOf(v);
    ctx.beginPath();ctx.moveTo(pad.l,yy);ctx.lineTo(W-pad.r,yy);ctx.stroke();
    ctx.fillStyle='rgba(90,96,128,.65)';ctx.font='9px DM Mono,monospace';ctx.textAlign='right';
    ctx.fillText(fmtMoney(v),pad.l-4,yy+3);
  }}

  // X labels (anos)
  ctx.fillStyle='rgba(90,96,128,.7)';ctx.font='9px DM Mono,monospace';ctx.textAlign='center';
  anos.forEach((a,i)=>ctx.fillText(String(a),xOf(i),H-pad.b+14));

  // Linhas por série
  series.forEach(sr=>{{
    ctx.strokeStyle=sr.cor;ctx.lineWidth=2;ctx.globalAlpha=.9;
    ctx.beginPath();
    anos.forEach((_,i)=>{{
      const x=xOf(i),y=yOf(sr.values[i]);
      i===0?ctx.moveTo(x,y):ctx.lineTo(x,y);
    }});
    ctx.stroke();
    // Pontos
    ctx.fillStyle=sr.cor;ctx.globalAlpha=1;
    anos.forEach((_,i)=>{{
      if(!sr.values[i]) return;
      ctx.beginPath();ctx.arc(xOf(i),yOf(sr.values[i]),3.5,0,Math.PI*2);ctx.fill();
    }});
  }});
  ctx.globalAlpha=1;

  // Legenda
  const lg=document.getElementById('timeline-legend');
  if(lg){{
    lg.innerHTML=series.map(s=>
      `<span style="display:flex;align-items:center;gap:5px">
        <span style="width:18px;height:3px;background:${{s.cor}};border-radius:2px;flex-shrink:0;display:inline-block"></span>
        <span style="color:rgba(200,210,230,.8);font-size:10px">${{s.label}}</span>
      </span>`
    ).join('');
  }}
}}





// ── Tabela ranking ─────────────────────────────────────────────────────────
let sortKey='n', sortAsc=false;
function sortBy(k){{
  if(sortKey===k){{sortAsc=!sortAsc;}}else{{sortKey=k;sortAsc=false;}}
  document.getElementById('sort-rank').value=k in {{n:1,roi_fsa_def:1,roi_tot_def:1,roi_fsa:1,roi_tot:1,rimax:1,rim:1,inv:1}}?k:'n';
  renderTable();
}}
function getFiltered(){{
  const q=document.getElementById('search-rank').value.toLowerCase();
  const f1=document.getElementById('filter-fase1').value;
  const sort=document.getElementById('sort-rank').value;
  let rows=CHM.filter(c=>(!q||c.ch.toLowerCase().includes(q))&&(!f1||c.f1===f1));
  const sk=sortKey||sort;
  rows.sort((a,b)=>{{
    const d=sk==='ch'?(a.ch>b.ch?1:-1):sk==='n'?b.n-a.n
      :sk==='roi_fsa_def'?(b.roi_fsa_def||0)-(a.roi_fsa_def||0)
      :sk==='roi_tot_def'?(b.roi_tot_def||0)-(a.roi_tot_def||0)
      :sk==='roi_fsa'?(b.roi_fsa||0)-(a.roi_fsa||0):sk==='roi_tot'?(b.roi_tot||0)-(a.roi_tot||0)
      :sk==='rda'?b.rda-a.rda:sk==='rimax'?b.rimax-a.rimax:sk==='rim'?b.rim-a.rim
      :sk==='bilh_avg'?(b.bilh_avg||0)-(a.bilh_avg||0):sk==='paises_avg'?(b.paises_avg||0)-(a.paises_avg||0):b.inv-a.inv;
    return sortAsc?-d:d;
  }});
  return rows;
}}
function renderTable(){{
  const rows=getFiltered();
  document.getElementById('rank-count').textContent=rows.length+' chamadas';
  const maxTot=Math.max(...rows.map(r=>r.roi_tot_def||r.roi_tot||0),.001);
  const maxRimax=Math.max(...rows.map(r=>r.rimax),1);
  document.getElementById('rank-table-wrap').innerHTML=`<table>
    <thead><tr>
      <th onclick="sortBy('ch')">Chamada</th>
      <th>Categoria</th>
      <th onclick="sortBy('n')">Obras</th>
      <th onclick="sortBy('inv')">Investimento</th>
      <th onclick="sortBy('roi_tot_def')">ROI Tot. Deflac.</th>
      <th onclick="sortBy('roi_fsa_def')">ROI FSA Deflac.</th>
      <th onclick="sortBy('rimax')">ROI Intl.</th>
    </tr></thead>
    <tbody>
    ${{rows.map(r=>`<tr onclick='openModal(${{JSON.stringify(r)}})'>
      <td>${{r.ch}}</td>
      <td><span class="tag" style="background:${{r.cor}}22;color:${{r.cor}}">${{r.f1l}}</span></td>
      <td>${{r.n}}</td>
      <td>${{fmtMoney(r.inv)}}</td>
      <td><div class="bar-inline"><div class="bar-bg"><div class="bar-fill" style="width:${{((r.roi_tot_def||r.roi_tot||0)/maxTot*80).toFixed(0)}}%;background:${{r.cor}}"></div></div>${{fmtRatio(r.roi_tot_def||r.roi_tot||0)}}</div></td>
      <td style="color:var(--muted)">${{fmtRatio(r.roi_fsa_def||r.roi_fsa||0)}}</td>
      <td><div class="bar-inline"><div class="bar-bg"><div class="bar-fill" style="width:${{(r.rimax/maxRimax*80).toFixed(0)}}%;background:${{r.cor}}"></div></div>${{r.rimax.toFixed(1)}}</div></td>
    </tr>`).join('')}}
    </tbody></table>`;
}}

// ── Modal ──────────────────────────────────────────────────────────────────
function openModal(c){{
  document.getElementById('modal-title').textContent=c.ch;
  document.getElementById('modal-meta').innerHTML=`
    <div class="modal-kpi">Critério <span style="color:${{c.cor}}">${{c.f1l}}</span></div>
    <div class="modal-kpi">Obras <span>${{c.n}}</span></div>
    <div class="modal-kpi">Investimento <span>${{fmtMoney(c.inv)}}</span></div>
    <div class="modal-kpi">ROI FSA Deflac. <span>${{fmtRatio(c.roi_fsa_def||c.roi_fsa||0)}}</span></div>
    <div class="modal-kpi">ROI Total Deflac. <span>${{fmtRatio(c.roi_tot_def||c.roi_tot||0)}}</span></div>
    <div class="modal-kpi">ROI intl máx <span>${{c.rimax.toFixed(1)}}</span></div>
  `;
  const obras=(c.obras||[]).slice(0,30).sort((a,b)=>b.rd-a.rd);
  const maxRd=Math.max(...obras.map(o=>o.rd),.001);
  document.getElementById('modal-content').innerHTML=`
    <table><thead><tr>
      <th>Obra</th><th>Ano</th><th>Bilheteria</th><th>Outras Janelas</th>
      <th>ROI Dom.</th><th>Países</th><th>Adm. EU</th><th>Festivais</th>
    </tr></thead>
    <tbody>${{obras.map(o=>`<tr>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{o.t}}">${{o.t}}</td>
      <td>${{o.a||'—'}}</td>
      <td>${{o.bilh?fmtMoney(o.bilh):'—'}}</td>
      <td>${{o.jan?fmtMoney(o.jan):'—'}}</td>
      <td><div class="bar-inline"><div class="bar-bg"><div class="bar-fill" style="width:${{(o.rd/maxRd*80).toFixed(0)}}%;background:${{c.cor}}"></div></div>${{fmtRatio(o.rd)}}</div></td>
      <td>${{o.np||'—'}}</td>
      <td>${{o.adm?fmtInt(o.adm):'—'}}</td>
      <td style="font-size:10px;color:var(--muted)">${{(o.fest||[]).join(', ')||'—'}}</td>
    </tr>`).join('')}}</tbody></table>
    ${{obras.length===30?'<p style="font-size:10px;color:var(--muted);margin-top:10px">Exibindo até 30 obras.</p>':''}}
  `;
  document.getElementById('modal-overlay').classList.add('open');
}}
function closeModal(e){{
  if(!e||e.target===document.getElementById('modal-overlay'))
    document.getElementById('modal-overlay').classList.remove('open');
}}
document.addEventListener('keydown',e=>{{if(e.key==='Escape')closeModal()}});

// ── Pills ──────────────────────────────────────────────────────────────────
function buildPills(containerId, onChange){{
  const el=document.getElementById(containerId);
  el.innerHTML='<div class="cat-pill active" data-key="">Todos</div>'
    +CATS.map(c=>`<div class="cat-pill" data-key="${{c.key}}" style="--pill-col:${{c.cor}}">${{c.label}} <span style="opacity:.5">${{c.n_chamadas}}</span></div>`).join('');
  el.querySelectorAll('.cat-pill').forEach(p=>{{
    p.style.setProperty('--p-col',p.dataset.key?CATS.find(c=>c.key===p.dataset.key)?.cor||'':'');
    p.addEventListener('click',()=>{{
      el.querySelectorAll('.cat-pill').forEach(x=>x.classList.remove('active'));
      p.classList.add('active');
      onChange(p.dataset.key);
    }});
  }});
}}

// ── Aba Obras Detalhadas ───────────────────────────────────────────────────
let sortObrasKey='roi_intl',sortObrasAsc=true;
function buildCatOptions(){{
  const sel=document.getElementById('filter-obras-cat');
  if(!sel||sel.options.length>1) return;
  CATS.forEach(c=>{{const o=document.createElement('option');o.value=c.key;o.textContent=c.label;sel.appendChild(o);}});
}}
function sortObrasBy(k){{
  if(sortObrasKey===k)sortObrasAsc=!sortObrasAsc;else{{sortObrasKey=k;sortObrasAsc=(k==='titulo');}}
  renderObras();
}}
function renderObras(){{
  const q  =document.getElementById('search-obras').value.toLowerCase();
  const cat=document.getElementById('filter-obras-cat').value;
  const sk =sortObrasKey;
  let rows=OBRAS_DET.filter(o=>(!q||o.titulo.toLowerCase().includes(q))&&(!cat||o.cat_key===cat));
  rows.sort((a,b)=>{{
    let d=sk==='titulo'?a.titulo.localeCompare(b.titulo,'pt-BR'):
          sk==='ano'?a.ano-b.ano:sk==='inv'?b.inv-a.inv:
          sk==='roi_fsa'?b.roi_fsa-a.roi_fsa:sk==='roi_tot'?b.roi_tot-a.roi_tot:sk==='roi_dom'?b.roi_dom-a.roi_dom:sk==='roi_intl'?b.roi_intl-a.roi_intl:
          sk==='fest'?b.fest-a.fest:sk==='bilh'?b.bilh-a.bilh:sk==='n_paises'?(b.n_paises||0)-(a.n_paises||0):sk==='adm'?(b.adm||0)-(a.adm||0):sk==='n_festivais'?(b.n_festivais||0)-(a.n_festivais||0):0;
    return sortObrasAsc?d:-d;
  }});
  document.getElementById('obras-count').textContent=rows.length+' registros';
  const maxInv=Math.max(...rows.map(r=>r.inv),1);
  const maxRd=Math.max(...rows.map(r=>r.roi_tot_def||r.roi_tot||0),.001);
  document.getElementById('obras-table-wrap').innerHTML=`<table><thead><tr>
    <th onclick="sortObrasBy('titulo')">Obra</th><th onclick="sortObrasBy('ano')">Ano</th>
    <th>Categoria</th>
    <th onclick="sortObrasBy('inv')">Investimento</th>
    <th onclick="sortObrasBy('bilh')">Bilheteria</th>
    <th onclick="sortObrasBy('roi_tot_def')">ROI Tot. Deflac.</th>
    <th onclick="sortObrasBy('roi_fsa_def')">ROI FSA Deflac.</th>
    <th onclick="sortObrasBy('roi_intl')">ROI Intl.</th>
    <th onclick="sortObrasBy('n_paises')">Países VOD</th>
    <th onclick="sortObrasBy('adm')">Lumière (EU)</th>
    <th onclick="sortObrasBy('n_festivais')">Festivais</th>
  </tr></thead><tbody>
  ${{rows.map(r=>`<tr>
    <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.titulo}}">${{r.titulo}}</td>
    <td>${{r.ano||'—'}}</td>
    <td><span class="tag" style="background:${{r.cor}}22;color:${{r.cor}};white-space:nowrap;font-size:9px">${{r.cat_label}}</span></td>
    <td><div class="bar-inline"><div class="bar-bg"><div class="bar-fill" style="width:${{(r.inv/maxInv*80).toFixed(0)}}%;background:${{r.cor}}"></div></div>${{fmtMoney(r.inv)}}</div></td>
    <td>${{r.bilh>0?fmtMoney(r.bilh):'—'}}</td>
    <td><div class="bar-inline"><div class="bar-bg"><div class="bar-fill" style="width:${{((r.roi_tot_def||r.roi_tot||0)/maxRd*80).toFixed(0)}}%;background:${{r.cor}}"></div></div>${{fmtRatio(r.roi_tot_def||r.roi_tot||0)}}</div></td>
    <td style="color:var(--muted)">${{fmtRatio(r.roi_fsa_def||r.roi_fsa||0)}}</td>
    <td>${{r.roi_intl.toFixed(1)}}</td>
    <td>${{r.n_paises>0?r.n_paises+' países':'—'}}</td>
    <td>${{r.adm>0?fmtInt(r.adm)+' adm.':'—'}}</td>
    <td style="font-size:10px;color:var(--muted);max-width:140px">${{(r.festivais||[]).join(', ')||'—'}}</td>
  </tr>`).join('')}}
  </tbody></table>`;
}}

// ── Descrições de categoria ────────────────────────────────────────────────
const CAT_DESC = {{
  "pont_fest_prod":  {{fase1:"Histórico de festivais internacionais (Fase 1 seletiva)", fase2:"Leitura de roteiro (Fase 2)", mecanismo:"FSA direto — produtora", nota:"PRODECINE 03–05 + Concurso Módulo B — seletivo c/ leitura de roteiro"}},
  "pont_com_prod":   {{fase1:"Desempenho comercial — bilheteria/market share (Fase 1 seletiva)", fase2:"Leitura de roteiro (Fase 2)", mecanismo:"FSA direto — produtora", nota:"PRODECINE 01, 06, Produção Cinema — seletivo c/ leitura de roteiro"}},
  "pont_com_dist":   {{fase1:"Market share / desempenho comercial da distribuidora (Fase 1 seletiva)", fase2:"Leitura de roteiro (Fase 2)", mecanismo:"FSA direto — distribuidora", nota:"PRODECINE 02, Cinema via Distribuidora — seletivo c/ leitura de roteiro"}},
  "complementacao":  {{fase1:"Complementação de financiamento para obras em produção ou finalização", fase2:"Sem leitura de roteiro — apoio complementar ao orçamento", mecanismo:"FSA direto — complementação", nota:"COMPLEMENTAÇÃO, PRODECINE 04 — apoio complementar à produção"}},
  "automatico":      {{fase1:"Resultado de bilheteria anterior (100% automático — sem seleção)", fase2:"Sem leitura de roteiro", mecanismo:"FSA direto — produtora ou distribuidora", nota:"Fluxo Contínuo Produção Cinema + Fluxo Contínuo via Distribuidora + Suporte Automático Comercial"}},
  "automatico_fest": {{fase1:"Resultado em festivais e reconhecimento artístico (automático)", fase2:"Sem leitura de roteiro", mecanismo:"FSA direto — artístico", nota:"Suporte Automático Artístico (SUAT) — festivais nacionais e internacionais"}},
  "coprod":          {{fase1:"Parceria com produtora estrangeira confirmada", fase2:"Avaliação do projeto binacional", mecanismo:"FSA direto — coprodução", nota:"PRODECINE 07–12, acordos bilaterais Chile, Portugal, Uruguai"}},
  "seletivo_roteiro":{{fase1:"Somente roteiro (sem histórico exigido)", fase2:"Avaliação de roteiro", mecanismo:"FSA direto", nota:"Cinema Novos Realizadores — foco em estreantes"}},
  "comercializacao": {{fase1:"Investimento direto em lançamento e distribuição de obras já produzidas", fase2:"Sem seleção por roteiro — foco em comercialização e distribuição", mecanismo:"FSA direto — comercialização / suporte automático comercial", nota:"COMERCIALIZAÇÃO EM CINEMA · COMPLEMENTAÇÃO · Opção de Investimento em Comercialização · Suporte Automático — Desempenho Comercial Cinema"}},
}};

function buildCatDescriptions(){{
  const el=document.getElementById('cat-desc-grid');
  if(!el) return;
  el.innerHTML=CATS.map(c=>{{
    const d=CAT_DESC[c.key]||{{}};
    const hp=c.n_com_paises>0;
    const selBadge=c.is_seletivo
      ?`<span style="font-size:8px;padding:1px 6px;border-radius:3px;background:rgba(100,200,255,.12);color:#6ecfff;letter-spacing:.05em;margin-left:6px;vertical-align:middle;border:1px solid rgba(100,200,255,.25)">SELETIVO</span>`
      :`<span style="font-size:8px;padding:1px 6px;border-radius:3px;background:rgba(100,255,180,.1);color:#50e8a0;letter-spacing:.05em;margin-left:6px;vertical-align:middle;border:1px solid rgba(100,255,180,.2)">AUTOMÁTICO</span>`;
    return `<div style="background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:14px;border-left:3px solid ${{c.cor}};cursor:pointer;transition:box-shadow .2s" onmouseenter="this.style.boxShadow='0 0 0 1px '+c.cor+'66'" onmouseleave="this.style.boxShadow=''" onclick="openCatWorks('${{c.key}}','${{c.label}}')">
      <div style="font-size:11px;font-weight:600;color:${{c.cor}};margin-bottom:8px;line-height:1.3">${{c.label}}${{selBadge}}</div>
      ${{d.fase1?`<div style="font-size:10px;color:var(--muted);margin-bottom:2px"><span style="color:rgba(200,210,230,.4)">Fase 1 · </span>${{d.fase1}}</div>`:''}}
      ${{d.fase2?`<div style="font-size:10px;color:var(--muted);margin-bottom:2px"><span style="color:rgba(200,210,230,.4)">Fase 2 · </span>${{d.fase2}}</div>`:''}}
      ${{d.mecanismo?`<div style="font-size:10px;color:var(--muted);margin-bottom:6px"><span style="color:rgba(200,210,230,.4)">Mec. · </span>${{d.mecanismo}}</div>`:''}}
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px 10px;margin-top:10px;padding-top:10px;border-top:1px solid var(--dim);align-items:start">
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">ROI Total Deflac. (R$2024)</div><div style="font-size:13px;font-weight:600;color:${{c.cor}}">${{c.roi_tot_def>0?fmtRatio(c.roi_tot_def):(c.roi_tot>0?fmtRatio(c.roi_tot):'—')}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">ROI FSA Deflac. (R$2024)</div><div style="font-size:13px;font-weight:600;color:${{c.cor}}">${{c.roi_fsa_def>0?fmtRatio(c.roi_fsa_def):(c.roi_fsa>0?fmtRatio(c.roi_fsa):'—')}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">ROI Internacional (0–100)</div><div style="font-size:13px;font-weight:600;color:${{c.cor}}">${{c.intl_avg>0?c.intl_avg.toFixed(1):'—'}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Investimento total</div><div style="font-size:11px;color:var(--text)">${{fmtMoney(c.inv)}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Bilheteria total</div><div style="font-size:11px;color:var(--text)">${{c.bilh_total>0?fmtMoney(c.bilh_total):'—'}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Receita prop. ao investimento</div><div style="font-size:11px;color:var(--text)">${{c.bilh_prop>0?fmtMoney(c.bilh_prop):'—'}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Bilheteria média/obra</div><div style="font-size:11px;color:var(--text)">${{c.bilh_avg>0?fmtMoney(c.bilh_avg):'—'}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Outras janelas média</div><div style="font-size:11px;color:var(--text)">${{c.jan_avg>0?fmtMoney(c.jan_avg):'—'}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Países VOD (média/obra)</div><div style="font-size:11px;color:var(--text)">${{hp?c.paises_avg.toFixed(1)+' países':'—'}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Países VOD (total contatos)</div><div style="font-size:11px;color:var(--text)">${{c.paises_total>0?c.paises_total+' em '+c.n_com_paises+' obras':'—'}}</div></div>
        <div style="grid-column:1/-1"><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Total seleções em festivais</div><div style="font-size:11px;color:var(--text)">${{c.total_festivais>0?c.total_festivais+' seleções em '+c.n_com_festivais+' obras':'—'}}</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Com dados</div><div style="font-size:11px;color:var(--accent)">${{c.n_com_dados}} obras ({{Math.round(c.n_com_dados/c.n_obras*100)}}%)</div></div>
        <div><div style="font-size:9px;color:rgba(200,210,230,.4);margin-bottom:2px">Sem dados</div><div style="font-size:11px;color:var(--muted)">${{c.n_sem_dados}} obras</div></div>
      </div>
      ${{d.nota?`<div style="font-size:9px;color:rgba(90,96,128,.7);margin-top:7px;border-top:1px solid var(--dim);padding-top:5px">${{d.nota}}</div>`:''}}
    </div>`;
  }}).join('');
}}

// ── Modal de categoria (obras consideradas / excluídas) ───────────────────
function openCatWorks(key, label){{
  const obras=OBRAS_DET.filter(o=>o.cat_key===key).sort((a,b)=>b.roi_dom-a.roi_dom);
  const excl=(EXCL[key]||[]);
  const cat=CATS.find(c=>c.key===key)||{{}};
  const maxRd=Math.max(...obras.map(o=>o.roi_dom),.001);
  document.getElementById('modal-title').textContent=label;
  document.getElementById('modal-meta').innerHTML=`
    <div class="modal-kpi">Incluídas <span style="color:var(--accent)">${{obras.length}}</span></div>
    <div class="modal-kpi">Excluídas <span style="color:var(--muted)">${{excl.length}}</span></div>
    <div class="modal-kpi">ROI Dom. Médio <span>${{fmtRatio(cat.rda_avg||0)}}</span></div>
    <div class="modal-kpi">ROI Proporcional <span>${{fmtRatio(cat.rda||0)}}</span></div>
    <div class="modal-kpi">Investimento <span>${{fmtMoney(cat.inv||0)}}</span></div>`;
  let html=`<div style="font-size:10px;font-weight:600;color:var(--accent);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Obras consideradas — ${{obras.length}}</div>
    <div style="overflow-x:auto;margin-bottom:20px"><table><thead><tr>
      <th>Obra</th><th>Ano</th><th>Investimento</th><th>Bilheteria</th><th>ROI Dom.</th><th>ROI Intl.</th><th>Festivais</th>
    </tr></thead><tbody>
    ${{obras.map(o=>`<tr>
      <td style="max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${{o.titulo}}">${{o.titulo}}</td>
      <td>${{o.ano||'—'}}</td><td>${{fmtMoney(o.inv)}}</td>
      <td>${{o.bilh>0?fmtMoney(o.bilh):'—'}}</td>
      <td><div class="bar-inline"><div class="bar-bg"><div class="bar-fill" style="width:${{Math.min(80,o.roi_dom/maxRd*80).toFixed(0)}}%;background:${{cat.cor||'var(--accent)'}}"></div></div>${{fmtRatio(o.roi_dom)}}</div></td>
      <td>${{o.roi_intl.toFixed(1)}}</td>
      <td style="font-size:10px;color:var(--muted)">${{(o.festivais||[]).join(', ')||'—'}}</td>
    </tr>`).join('')}}
    </tbody></table></div>`;
  if(excl.length>0){{
    html+=`<div style="font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">Obras não consideradas — ${{excl.length}}</div>
    <p style="font-size:10px;color:var(--muted);margin-bottom:8px;line-height:1.6">Obras com captações nesta categoria excluídas da análise: produção TV/VOD, renúncia sem FSA cinema confirmado, ou dados insuficientes.</p>
    <div style="overflow-x:auto"><table><thead><tr><th>Obra</th><th>Ano</th><th>Bilheteria</th></tr></thead><tbody>
    ${{excl.map(o=>`<tr><td style="max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${{o.t}}</td><td>${{o.a||'—'}}</td><td>${{o.bilh>0?fmtMoney(o.bilh):'—'}}</td></tr>`).join('')}}
    </tbody></table></div>`;
  }}
  document.getElementById('modal-content').innerHTML=html;
  document.getElementById('modal-overlay').classList.add('open');
}}

// ── Init ───────────────────────────────────────────────────────────────────
buildKpis();
requestAnimationFrame(()=>{{{{ drawQuadrant(); drawRankings(); drawTimeline(); }}}});
buildCatDescriptions();

buildPills('cat-pills-rank',key=>{{
  document.getElementById('filter-fase1').value=key;
  renderTable();
}});
renderTable();
window.addEventListener('resize',()=>{{
  drawQuadrant(); drawRankings(); drawTimeline();
}});

// ── Síntese: melhor metodologia por métrica ───────────────────────────────
function _fmtVal(v, metric) {{
  if(metric.includes('roi') || metric.includes('rda')) return (v||0).toFixed(3)+'x';
  if(metric.includes('paises')) return (v||0).toFixed(1);
  if(metric==='evid_avg' || metric.includes('intl')) return (v||0).toFixed(1);
  return (v||0).toFixed(3);
}}
function renderSintese() {{
  const el = document.getElementById('sint-table');
  if(!el || el.dataset.built) return;
  el.dataset.built = '1';
  let html = `<table style="width:100%;border-collapse:collapse;font-size:12px">
  <thead>
    <tr style="border-bottom:1px solid var(--border)">
      <th style="text-align:left;padding:8px 12px;color:var(--muted);font-weight:400;letter-spacing:.06em;text-transform:uppercase;font-size:10px">Métrica</th>
      <th style="text-align:left;padding:8px 12px;color:var(--muted);font-weight:400;letter-spacing:.06em;text-transform:uppercase;font-size:10px">1º Lugar</th>
      <th style="text-align:left;padding:8px 12px;color:var(--muted);font-weight:400;letter-spacing:.06em;text-transform:uppercase;font-size:10px">2º Lugar</th>
      <th style="text-align:left;padding:8px 12px;color:var(--muted);font-weight:400;letter-spacing:.06em;text-transform:uppercase;font-size:10px">3º Lugar</th>
      <th style="text-align:left;padding:8px 12px;color:var(--muted);font-weight:400;letter-spacing:.06em;text-transform:uppercase;font-size:10px;max-width:280px">Nota</th>
    </tr>
  </thead>
  <tbody>`;
  SINTESE.forEach((row, ri) => {{
    const bg = ri%2===0 ? 'background:var(--surface)' : '';
    html += `<tr style="${{bg}};border-bottom:1px solid var(--border)">
      <td style="padding:10px 12px;font-weight:600;font-size:11px">${{row.label}}</td>`;
    for(let i=0;i<3;i++) {{
      const t = row.top3[i];
      if(t) {{
        const medal = i===0?'🥇':i===1?'🥈':'🥉';
        html += `<td style="padding:10px 12px">
          <div style="display:flex;align-items:center;gap:6px">
            <span style="width:8px;height:8px;border-radius:50%;background:${{t.cor}};flex-shrink:0;display:inline-block"></span>
            <span style="color:var(--text)">${{t.label}}</span>
          </div>
          <div style="color:var(--accent);font-size:11px;margin-top:2px;padding-left:14px">${{_fmtVal(t.val, row.metric)}} &nbsp;<span style="color:var(--muted);font-size:10px">(${{t.n}} obras)</span></div>
        </td>`;
      }} else {{
        html += `<td style="padding:10px 12px;color:var(--muted)">—</td>`;
      }}
    }}
    html += `<td style="padding:10px 12px;color:var(--muted);font-size:11px;max-width:280px;line-height:1.5">${{row.desc}}</td></tr>`;
  }});
  html += `</tbody></table>
  <div style="margin-top:12px;font-size:10px;color:var(--muted);line-height:1.6">
    <strong>Nota:</strong> Apenas categorias FSA com ≥ 10 obras. Renúncia fiscal excluída — é tratada como co-investimento de mercado, não como instrumento gerenciado.
    Os valores representam a métrica agregada por categoria (receita proporcional ao FSA / investimento FSA total da categoria).
  </div>`;
  el.innerHTML = html;
}}
</script>
</body>
</html>"""

out_html = os.path.join(ROOT, "resultados", "painel_criterio_selecao.html")
with open(out_html, "w", encoding="utf-8") as f:
    f.write(html)

print(f"[OK] Gerado: {out_html}")
print()
print(f"Registros com sobreposição: {len(obras_registros)}")
print(f"Categorias válidas (>= {MIN_OBRAS} obras): {len(cats)}")
print()
for c in sorted(cats, key=lambda x: -x["n_obras"]):
    print(f"  {c['label'][:55]:55s}  n={c['n_obras']:5d}  ROI_dom={c['roi_dom_agg']:.3f}x")
excl = [k for k in CATEGORIAS if k not in valid_keys]
if excl:
    print(f"\nExcluídas (< {MIN_OBRAS} obras): {', '.join(excl)}")
