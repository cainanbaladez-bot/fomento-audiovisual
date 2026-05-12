# -*- coding: utf-8 -*-
"""
gerar_painel_produtoras.py
Gera painel_produtoras.html — dashboard 3 abas para produtoras independentes.

Join chain:
  tabela_consolidada_obras.xlsx
    → normalize title → obras_fomento_unificado.csv (CPB)
    → produtores-de-obras-nao-publicitarias-brasileiras.csv (CNPJ_PRODUTOR)
    → filter agentes-economicos-regulares.csv (BRASILEIRO_INDEPENDENTE=SIM)
"""
import json
import pathlib
import re
import unicodedata

import numpy as np
import pandas as pd

BASE = pathlib.Path(__file__).parent.parent

# ── helpers ───────────────────────────────────────────────────────────────

def norm_title(s):
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().upper()

def norm_cnpj(s):
    if not isinstance(s, str):
        s = str(s) if pd.notna(s) else ""
    digits = re.sub(r"\D", "", s)
    return digits.lstrip("0").zfill(14)

# ── 1. Load tabela consolidada ─────────────────────────────────────────────
print("Carregando tabela consolidada (PERIODO AMPLO - TODO HISTORICO)...")
_master_xlsx = BASE / "resultados" / "tabela_consolidada_obras.xlsx"
try:
    df = pd.read_excel(_master_xlsx, sheet_name="Obras")
except Exception:
    df = pd.read_excel(_master_xlsx, sheet_name=0)
df.columns = df.columns.str.strip()
print(f"  {len(df)} obras carregadas (sem filtro de período)")

COL_MAP = {
    "Projeto": "titulo",
    "Ano": "ano",
    "Chamada": "chamada",
    "Categoria": "categoria",
    "Valor FSA (R$)": "vfsa",
    "Renúncia Art.3/3-A/39 (R$)": "vren3",
    "Renúncia Outros Mec. (R$)": "vrenO",
    "Bilheteria Nominal (R$)": "bilheteria",
    "Bilheteria Deflac. (R$)": "bilh_def",
    "Estimativa Outras Janelas (R$)": "janelas",
    "Outras Janelas Deflac. (R$2024)": "janelas_def",
    "Investimento Total Deflac. (R$2024)": "inv_def_col",
    "ROI Internacional (0-100)": "roi_intl",
    "VOD Intl \u2014 N Pa\u00edses": "n_paises",
    "Todas Chamadas FSA": "todas_chamadas",
    # Festivais internacionais → países
    "Festival \u2014 Oscar": "f_oscar",
    "Festival \u2014 Cannes": "f_cannes",
    "Festival \u2014 Berlim": "f_berlim",
    "Festival \u2014 Veneza": "f_veneza",
    "Festival \u2014 Sundance": "f_sundance",
    "Festival \u2014 Locarno": "f_locarno",
    "Festival \u2014 TIFF": "f_tiff",
    "Festival \u2014 San Sebasti\u00e1n": "f_sanseb",
    "Festival \u2014 Rotterdam": "f_rotterdam",
    "Festival \u2014 Annecy": "f_annecy",
    "Festival \u2014 NYFF": "f_nyff",
    "Festival \u2014 BFI London": "f_bfi",
    "Festival \u2014 BAFTA": "f_bafta",
    "Festival \u2014 Globo Ouro": "f_globo",
    "Festival \u2014 Havana": "f_havana",
    "CRITICA_INDICE_1_5": "critica",
    "CRITICA_N_FONTES":   "critica_n",
}
df = df.rename(columns={k: v for k, v in COL_MAP.items() if k in df.columns})

FEST_COLS = ["f_oscar","f_cannes","f_berlim","f_veneza","f_sundance","f_locarno",
             "f_tiff","f_sanseb","f_rotterdam","f_annecy","f_nyff","f_bfi",
             "f_bafta","f_globo","f_havana"]
FESTIVAL_COUNTRIES = {
    "f_oscar": "US", "f_cannes": "FR", "f_berlim": "DE", "f_veneza": "IT",
    "f_sundance": "US", "f_locarno": "CH", "f_tiff": "CA", "f_sanseb": "ES",
    "f_rotterdam": "NL", "f_annecy": "FR", "f_nyff": "US", "f_bfi": "GB",
    "f_bafta": "GB", "f_globo": "US", "f_havana": "CU",
}

for col in ["vfsa", "vren3", "vrenO", "bilheteria", "bilh_def", "janelas", "janelas_def", "inv_def_col", "roi_intl", "n_paises", "critica", "critica_n"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    else:
        df[col] = 0.0

for col in FEST_COLS:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    else:
        df[col] = 0.0

if "ano" in df.columns:
    df["ano"] = pd.to_numeric(df["ano"], errors="coerce").fillna(0).astype(int)
else:
    df["ano"] = 0

if "titulo" not in df.columns:
    df["titulo"] = ""
if "categoria" not in df.columns:
    df["categoria"] = ""

df["titulo_norm"] = df["titulo"].apply(norm_title)
df["inv_linha"] = df["vfsa"] + df["vren3"] + df["vrenO"]
df["full_rev"] = df["bilheteria"]

# ── 2. CPB lookup a partir dos arquivos ANCINE brutos ─────────────────────
def _build_cpb_lookup(base):
    """Substitui obras_fomento_unificado.csv — full outer join on-the-fly."""
    fsa = pd.read_csv(
        base / "raw" / "obras-nao-pub-brasileiras-investimento-fsa.csv",
        sep=None, engine="python", encoding="utf-8-sig", dtype=str,
        usecols=["CPB", "TITULO_ORIGINAL"], on_bad_lines="skip",
    ).fillna("").drop_duplicates("CPB")
    ind = pd.read_csv(
        base / "raw" / "obras-nao-pub-brasileiras-fomento-indireto.csv",
        sep=None, engine="python", encoding="utf-8-sig", dtype=str,
        usecols=["CPB", "TITULO_ORIGINAL"], on_bad_lines="skip",
    ).fillna("").drop_duplicates("CPB")
    return pd.concat([fsa, ind]).drop_duplicates("CPB", keep="first")

print("Carregando lumiere_vod_search.xlsx para países de alcance...")
try:
    lum_vod = pd.read_excel(BASE / "raw" / "lumiere_vod_search.xlsx", usecols=["Original title", "Country"])
    lum_vod["titulo_norm"] = lum_vod["Original title"].apply(norm_title)
    lum_countries: dict = (
        lum_vod.groupby("titulo_norm")["Country"]
        .apply(lambda x: set(x.dropna().astype(str)))
        .to_dict()
    )
    print(f"  {len(lum_countries)} títulos com dados Lumière VOD")
except Exception as e:
    print(f"  AVISO: {e}")
    lum_countries = {}

print("Construindo lookup CPB-titulo a partir dos arquivos ANCINE brutos...")
obras = _build_cpb_lookup(BASE)
obras["titulo_norm"] = obras["TITULO_ORIGINAL"].apply(norm_title)
obras = obras.drop_duplicates("titulo_norm")

# CPB já vem do Excel (pré-requisito gerar_tabela_consolidada corrigido)
# Dropa antes do merge para evitar conflito de colunas
if "CPB" in df.columns:
    df = df.drop(columns=["CPB"])
df = df.merge(obras[["titulo_norm", "CPB", "TITULO_ORIGINAL"]], on="titulo_norm", how="left")
print(f"  CPB matched: {df['CPB'].notna().sum()}/{len(df)} linhas")

# ── 3. produtores → CNPJ_PRODUTOR (ALL producers, not just primary) ────────
print("Carregando produtores-de-obras-nao-publicitarias-brasileiras.csv ...")
prod = pd.read_csv(
    BASE / "raw" / "produtores-de-obras-nao-publicitarias-brasileiras.csv",
    sep=";",
    encoding="utf-8",
    usecols=["CPB", "CNPJ_PRODUTOR"],
    on_bad_lines="skip",
)
prod = prod.dropna(subset=["CPB", "CNPJ_PRODUTOR"])
prod["cnpj_norm"] = prod["CNPJ_PRODUTOR"].apply(norm_cnpj)

# ── 4. agentes → filter independentes ─────────────────────────────────────
print("Carregando agentes-economicos-regulares.csv ...")
agentes = pd.read_csv(
    BASE / "raw" / "agentes-economicos-regulares.csv",
    sep=";",
    encoding="utf-8",
    usecols=["RAZAO_SOCIAL", "CNPJ", "BRASILEIRO_INDEPENDENTE"],
    on_bad_lines="skip",
)
agentes["cnpj_norm"] = agentes["CNPJ"].apply(norm_cnpj)
ind = agentes[agentes["BRASILEIRO_INDEPENDENTE"].str.strip().str.upper() == "SIM"].copy()
ind = ind.drop_duplicates("cnpj_norm")
print(f"  Agentes independentes: {len(ind)}")
ind_cnpjs = set(ind["cnpj_norm"])

# Filter all producers to independent ones only
cpb_to_cnpj_ind = prod[prod["cnpj_norm"].isin(ind_cnpjs)][["CPB", "cnpj_norm"]].drop_duplicates()

# Build titulo_norm → CPB bridge (one CPB per obra, via obras_fomento_unificado)
cpb_map = df[["titulo_norm", "CPB"]].dropna(subset=["CPB"]).drop_duplicates("titulo_norm")

# Many-to-many: titulo_norm → cnpj_norm (each obra gets all its independent co-producers)
obra_prod_map = cpb_to_cnpj_ind.merge(cpb_map, on="CPB", how="inner")[["titulo_norm", "cnpj_norm"]].drop_duplicates()
obra_prod_map = obra_prod_map.merge(ind[["cnpj_norm", "RAZAO_SOCIAL"]], on="cnpj_norm", how="left")

# Explode: each df row gets replicated for each independent producer of its obra
df_ind = df.merge(obra_prod_map, on="titulo_norm", how="inner")
print(f"  Linhas com produtora independente identificada: {len(df_ind)}")
print(f"  Produtoras únicas: {df_ind['cnpj_norm'].nunique()}")

# ── 5. Build raw per-produtora data (pre-cluster) ──────────────────────────
print("Agregando por produtora ...")

raw_records = []
for cnpj, grp in df_ind.groupby("cnpj_norm"):
    nm = grp["RAZAO_SOCIAL"].iloc[0]

    # Deduplicated obra-level metrics (one row per obra)
    obra_dedup = (
        grp.groupby("titulo_norm")
        .agg(
            bilheteria=("bilheteria", "max"),
            bilh_def=("bilh_def", "max"),
            janelas=("janelas", "max"),
            janelas_def=("janelas_def", "max"),
            roi_intl=("roi_intl", "max"),
            n_paises=("n_paises", "max"),
            ano=("ano", "first"),
            titulo=("titulo", "first"),
        )
        .reset_index()
    )
    n = len(obra_dedup)

    pub = float(obra_dedup["bilheteria"].sum())
    inv_total = float(grp["inv_def_col"].sum())        # deflacionado (R$2024)
    rec_total = float(obra_dedup["bilh_def"].sum()) + float(obra_dedup["janelas_def"].sum())
    # FSA vs Renúncia deflacionados (proporção nominal aplicada ao total deflac por obra)
    def _fsa_ren_deflac(g):
        inv_tot_nom = (g["vfsa"] + g["vren3"] + g["vrenO"]).sum()
        fsa_nom     = g["vfsa"].sum()
        frac_fsa    = (fsa_nom / inv_tot_nom) if inv_tot_nom > 0 else 0.0
        inv_d       = g["inv_def_col"].sum()
        return inv_d * frac_fsa, inv_d * (1 - frac_fsa)
    inv_fsa_d, inv_ren_d = _fsa_ren_deflac(grp)

    rda = rec_total / inv_total if inv_total >= 1000 else 0.0

    # rdm: median of per-obra ROI deflacionado
    obra_rois = []
    for tit, og in grp.groupby("titulo_norm"):
        oi = og["inv_def_col"].sum()
        full_rev_obra = og["bilh_def"].max()
        if oi >= 1000:
            obra_rois.append(full_rev_obra / oi)
    rdm = float(np.median(obra_rois)) if obra_rois else 0.0

    rim = float(obra_dedup["roi_intl"].max())
    ria = float(obra_dedup["roi_intl"].mean())
    np_total = int(obra_dedup["n_paises"].sum())

    a0 = int(obra_dedup["ano"].min())
    a1 = int(obra_dedup["ano"].max())

    all_paises: set = set()
    obras_list = []
    for tit, og in grp.groupby("titulo_norm"):
        oi_def = og["inv_def_col"].sum()
        bilh_def_ob = float(og["bilh_def"].max())
        jan_def_ob  = float(og["janelas_def"].max())
        full_rev_def = bilh_def_ob + jan_def_ob
        rd_val = full_rev_def / oi_def if oi_def >= 1000 else 0.0
        ri_val = float(og["roi_intl"].max())
        td_col = og["TITULO_ORIGINAL"].dropna() if "TITULO_ORIGINAL" in og.columns else pd.Series(dtype=str)
        # Países únicos: festivais + Lumière VOD
        paises_obra: set = set()
        for fc, cc in FESTIVAL_COUNTRIES.items():
            if fc in og.columns and og[fc].max() > 0:
                paises_obra.add(cc)
        paises_obra |= lum_countries.get(tit, set())
        all_paises |= paises_obra
        td_val = str(td_col.iloc[0]) if len(td_col) else str(og["titulo"].iloc[0]) if "titulo" in og.columns else tit
        cat_col = og["categoria"].dropna() if "categoria" in og.columns else pd.Series(dtype=str)
        cat_val = str(cat_col.iloc[0]) if len(cat_col) else ""
        cr_val  = round(float(og["critica"].max()), 2)  if "critica"   in og.columns else 0.0
        crn_val = int(og["critica_n"].max())             if "critica_n" in og.columns else 0
        obras_list.append({
            "td": td_val,
            "a": int(og["ano"].iloc[0]),
            "p": int(float(og["bilheteria"].max())),
            "bd": int(bilh_def_ob),
            "rec": int(full_rev_def),
            "rd": round(rd_val, 3),
            "ri": round(ri_val, 1),
            "inv": int(float(og["inv_linha"].sum())),
            "inv_def": int(oi_def),
            "cat": cat_val,
            "np": len(paises_obra),
            "cr": cr_val,
            "crn": crn_val,
        })
    obras_list.sort(key=lambda x: x["rd"], reverse=True)

    critica_obras = [o["cr"] for o in obras_list if o["crn"] >= 2 and o["cr"] > 0]
    critica_avg = round(sum(critica_obras) / len(critica_obras), 2) if critica_obras else 0.0

    raw_records.append({
        "nm": nm,
        "cnpj": cnpj,
        "tipo": "produtora",
        "n": n,
        "pub": int(pub),
        "pub_def": int(obra_dedup["bilh_def"].sum()),
        "inv": int(float(grp["inv_linha"].sum())),
        "inv_def": int(inv_total),
        "inv_fsa_d": int(inv_fsa_d),
        "inv_ren_d": int(inv_ren_d),
        "rec": int(rec_total),
        "rec_def": int(rec_total),
        "rda": round(rda, 4),
        "roi_def": round(rda, 4),
        "rdm": round(rdm, 4),
        "rim": round(rim, 1),
        "ria": round(ria, 1),
        "np": np_total,
        "np_uniq": len(all_paises),
        "paises": sorted(all_paises),
        "a0": a0,
        "a1": a1,
        "obras": obras_list,
        "critica_avg": critica_avg,
        "critica_n_obras": len(critica_obras),
    })

print(f"  Produtoras: {len(raw_records)}")

# ── 5b. Add FSA captadores (proponentes) not yet in produtora list ─────────
import csv as _csv

def _fix_enc(s):
    try:
        return s.encode('latin-1').decode('utf-8')
    except Exception:
        return s

cap_map_fsa: dict = {}
_projetos_fsa = BASE / "raw" / "projetos-fsa.csv"
try:
    with open(_projetos_fsa, encoding='latin-1') as _f:
        for _row in _csv.DictReader(_f, delimiter=';'):
            _tit = norm_title(_fix_enc(_row.get('TITULO_PROJETO', '') or ''))
            _prop = _fix_enc(_row.get('RAZAO_SOCIAL_PROPONENTE', '') or '').strip()
            _prod = _fix_enc(_row.get('RAZAO_SOCIAL_PRODUTORA', '') or '').strip()
            if not _tit:
                continue
            _captador = _prop if _prop else _prod
            _tipo = 'produtora' if (not _prop) or _prop.upper() == _prod.upper() else 'captador'
            cap_map_fsa[_tit] = {'captador': _captador, 'tipo': _tipo}
    print(f"  cap_map_fsa: {len(cap_map_fsa)} títulos")
except Exception as _e:
    print(f"  AVISO cap_map_fsa: {_e}")

if cap_map_fsa:
    df["_captador"] = df["titulo_norm"].map(lambda t: cap_map_fsa.get(t, {}).get("captador", ""))
    df["_captador_tipo"] = df["titulo_norm"].map(lambda t: cap_map_fsa.get(t, {}).get("tipo", "produtora"))

    # Only FSA-funded obras, exclude TV
    _cat_col = "categoria" if "categoria" in df.columns else None
    if _cat_col:
        df_cap = df[(df["vfsa"] > 0) & (df["_captador"] != "") &
                    (~df[_cat_col].str.upper().str.contains("TV_EXCLUIR", na=False))].copy()
    else:
        df_cap = df[(df["vfsa"] > 0) & (df["_captador"] != "")].copy()

    existing_names_up = {r["nm"].upper() for r in raw_records}
    cap_raw: list = []

    for cap_nm, grp in df_cap.groupby("_captador"):
        if not cap_nm:
            continue
        # Only process distribuidoras (proponente ≠ produtora)
        if grp["_captador_tipo"].iloc[0] != "captador":
            continue

        obra_dedup2 = (
            grp.groupby("titulo_norm")
            .agg(
                bilheteria=("bilheteria", "max"),
                bilh_def=("bilh_def", "max"),
                janelas=("janelas", "max"),
                janelas_def=("janelas_def", "max"),
                roi_intl=("roi_intl", "max"),
                n_paises=("n_paises", "max"),
                ano=("ano", "first"),
                titulo=("titulo", "first"),
            )
            .reset_index()
        )
        n2 = len(obra_dedup2)

        pub2 = float(obra_dedup2["bilheteria"].sum())
        inv2 = float(grp["inv_def_col"].sum())       # deflacionado (R$2024)
        rec2 = float(obra_dedup2["bilh_def"].sum()) + float(obra_dedup2["janelas_def"].sum())
        rda2 = rec2 / inv2 if inv2 >= 1000 else 0.0

        obra_rois2 = []
        for _tit2, _og2 in grp.groupby("titulo_norm"):
            _oi2 = _og2["inv_def_col"].sum()
            _fr2 = _og2["bilh_def"].max()
            if _oi2 >= 1000:
                obra_rois2.append(_fr2 / _oi2)
        rdm2 = float(np.median(obra_rois2)) if obra_rois2 else 0.0

        rim2 = float(obra_dedup2["roi_intl"].max())
        ria2 = float(obra_dedup2["roi_intl"].mean())

        a0_2 = int(obra_dedup2["ano"].min())
        a1_2 = int(obra_dedup2["ano"].max())

        all_p2: set = set()
        obras2 = []
        for _tit2, _og2 in grp.groupby("titulo_norm"):
            _oi2_def = _og2["inv_def_col"].sum()
            _bd2 = float(_og2["bilh_def"].max())
            _jd2 = float(_og2["janelas_def"].max())
            _fr2 = _bd2 + _jd2
            _rd2 = _fr2 / _oi2_def if _oi2_def >= 1000 else 0.0
            _ri2 = float(_og2["roi_intl"].max())
            _paises2: set = set()
            for _fc, _cc in FESTIVAL_COUNTRIES.items():
                if _fc in _og2.columns and _og2[_fc].max() > 0:
                    _paises2.add(_cc)
                    all_p2.add(_cc)
            _paises2 |= lum_countries.get(_tit2, set())
            all_p2 |= _paises2
            _cat2 = str(_og2["categoria"].dropna().iloc[0]) if "categoria" in _og2.columns and len(_og2["categoria"].dropna()) else ""
            _cr2  = round(float(_og2["critica"].max()), 2)  if "critica"   in _og2.columns else 0.0
            _crn2 = int(_og2["critica_n"].max())             if "critica_n" in _og2.columns else 0
            obras2.append({
                "td": str(_og2["titulo"].iloc[0]) if "titulo" in _og2.columns else _tit2,
                "a": int(_og2["ano"].iloc[0]),
                "p": int(float(_og2["bilheteria"].max())), "bd": int(_bd2),
                "rec": int(_fr2),
                "rd": round(_rd2, 3), "ri": round(_ri2, 1),
                "inv": int(float(_og2["inv_linha"].sum())), "inv_def": int(_oi2_def),
                "cat": _cat2,
                "np": len(_paises2),
                "cr": _cr2,
                "crn": _crn2,
            })
        obras2.sort(key=lambda x: x["rd"], reverse=True)
        critica_obras2 = [o["cr"] for o in obras2 if o["crn"] >= 2 and o["cr"] > 0]
        critica_avg2 = round(sum(critica_obras2) / len(critica_obras2), 2) if critica_obras2 else 0.0

        cap_raw.append({
            "nm": cap_nm,
            "tipo": "captador",
            "n": n2,
            "pub": int(pub2),
            "pub_def": int(obra_dedup2["bilh_def"].sum()),
            "inv": int(float(grp["inv_linha"].sum())),
            "inv_def": int(inv2),
            "rec": int(rec2),
            "rec_def": int(rec2),
            "rda": round(rda2, 4),
            "roi_def": round(rda2, 4),
            "rdm": round(rdm2, 4),
            "rim": round(rim2, 1),
            "ria": round(ria2, 1),
            "np": 0,
            "np_uniq": len(all_p2),
            "paises": sorted(all_p2),
            "a0": a0_2,
            "a1": a1_2,
            "obras": obras2,
            "critica_avg": critica_avg2,
            "critica_n_obras": len(critica_obras2),
        })

    # Captadores replace any existing produtora records with same name (captador data is more complete)
    cap_names_up = {r["nm"].upper() for r in cap_raw}
    raw_records = [r for r in raw_records if r["nm"].upper() not in cap_names_up]
    raw_records.extend(cap_raw)
    print(f"  Captadores/distribuidoras: {len(cap_raw)} (substituíram registros produtora com mesmo nome)")

# ── 6. Cluster assignment ──────────────────────────────────────────────────
# Critérios (tabela de apoio: Clusters_produtoras.xlsx)
BILH_THRESHOLD_HIGH = 2_500_000   # R$ 2,5M — limiar receita (Duplo Retorno + piso ROI Dom)
INV_THRESHOLD_HIGH  = 5_000_000   # R$ 5M   — limiar investimento (sem_retorno)
REC_NEW_THRESHOLD   = 10_000_000  # R$ 10M  — limiar superior (Retorno Doméstico via receita)
ROI_DOM_THRESHOLD   = 0.6         # ROI doméstico mínimo para Retorno Doméstico via ratio
REC_PISO            = 2_500_000   # piso de receita para critério de ROI Dom

print(f"\n-- Thresholds ------------------------------------------")
print(f"  BILH_HIGH : R$ {BILH_THRESHOLD_HIGH:,.0f}")
print(f"  INV_HIGH  : R$ {INV_THRESHOLD_HIGH:,.0f}")
print(f"  REC_NEW   : R$ {REC_NEW_THRESHOLD:,.0f}")
print(f"  ROI_DOM   : {ROI_DOM_THRESHOLD}")
print(f"--------------------------------------------------------\n")

def assign_cluster(r):
    rec  = r.get("rec", 0)   # receita total = bilheteria + janelas
    rim  = r.get("rim", 0)
    inv  = r.get("inv", 0)
    rda  = r.get("rda", 0)   # ROI doméstico total deflacionado
    duplo_intl = rim >= 13   # cauda superior: acima do cluster denso 10–13
    if rec >= BILH_THRESHOLD_HIGH and duplo_intl:  return "duplo"
    # Retorno Doméstico: renda >= R$10M  OU  (ROI Dom > 0,6 E renda >= R$2,5M)
    is_dom = (rec >= REC_NEW_THRESHOLD) or (rda > ROI_DOM_THRESHOLD and rec >= REC_PISO)
    if is_dom and not duplo_intl:                  return "dom"
    if duplo_intl:                                 return "intl"
    if inv > INV_THRESHOLD_HIGH:                   return "sem_retorno"
    return "pequeno"

records = []
for r in raw_records:
    cl = assign_cluster(r)
    rec = dict(r)
    rec["cl"] = cl
    rec.pop("cnpj", None)
    records.append(rec)

records.sort(key=lambda x: (-x["n"], -x["rda"]))

cluster_counts = {}
for r in records:
    cluster_counts[r["cl"]] = cluster_counts.get(r["cl"], 0) + 1

print("-- Cluster counts --------------------------------------")
for cl, cnt in sorted(cluster_counts.items(), key=lambda x: -x[1]):
    print(f"  {cl:20s}: {cnt}")
print(f"  {'TOTAL':20s}: {len(records)}")
print(f"--------------------------------------------------------\n")

all_anos = [r["a0"] for r in records] + [r["a1"] for r in records]
year_min = min(all_anos) if all_anos else 2014
year_max = max(all_anos) if all_anos else 2024

json_data = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
thresholds_json = json.dumps(
    {"bilh_high": BILH_THRESHOLD_HIGH, "inv_high": INV_THRESHOLD_HIGH,
     "rda": 0.25, "rim": 1.0},
    ensure_ascii=False
)

# ── 7. Generate HTML ───────────────────────────────────────────────────────
print("Gerando painel_produtoras.html ...")

HTML = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Produtoras Independentes BR — Painel</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Mono:wght@300;400;500&family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg:#0b0d14;--surface:#14171f;--surface2:#1a1e2c;--surface3:#212638;
  --border:#282d42;--border-light:#343a54;
  --accent:#6c7bf7;--accent-dim:rgba(108,123,247,.12);
  --gold:#fbbf24;--coral:#f87171;--purple:#a78bfa;--cyan:#38bdf8;
  --green:#34d399;--muted-blue:#5fd1ff;
  --text:#e2e8f0;--text2:#c1c9d9;--muted:#7b849a;--dim:#282d42;
  --font-head:'DM Serif Display',serif;--font-mono:'DM Mono',monospace;--font-ui:'Inter',system-ui,sans-serif;
}}
html,body{{height:100%;overflow:hidden}}
body{{background:var(--bg);color:var(--text);font-family:var(--font-mono);font-size:12px;display:flex;flex-direction:column}}

/* Header */
.hdr{{padding:10px 18px 0;flex-shrink:0;display:flex;align-items:baseline;gap:10px;flex-wrap:wrap}}
.hdr-t{{font-family:var(--font-head);font-size:16px;font-weight:800;letter-spacing:-.01em}}
.hdr-s{{color:var(--muted);font-size:10px}}

/* Tabs */
.tab-bar{{display:flex;gap:0;padding:8px 18px 0;flex-shrink:0;border-bottom:1px solid var(--border)}}
.tab-btn{{padding:6px 14px;font-family:var(--font-mono);font-size:10px;letter-spacing:.1em;text-transform:uppercase;
  background:none;color:var(--muted);border:none;border-bottom:2px solid transparent;cursor:pointer;
  transition:color .15s,border-color .15s;margin-bottom:-1px}}
.tab-btn:hover{{color:var(--text)}}
.tab-btn.active{{color:var(--accent);border-bottom-color:var(--accent)}}

/* Tab content */
.tab-panel{{display:none;flex:1;flex-direction:column;min-height:0;overflow:hidden}}
.tab-panel.active{{display:flex}}

/* ── Tab 1: Visão Geral ── */
.t1-scroll{{flex:1;overflow-y:auto;padding:14px 18px 18px}}
.t1-scroll::-webkit-scrollbar{{width:4px}}
.t1-scroll::-webkit-scrollbar-thumb{{background:var(--border);border-radius:2px}}

.kpi-bar{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px}}
.kpi{{background:var(--surface);border:1px solid var(--border);border-radius:3px;padding:10px 14px;min-width:130px;flex:1}}
.kpi-l{{font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:5px}}
.kpi-v{{font-family:var(--font-head);font-size:20px;font-weight:700;color:var(--text)}}
.kpi-u{{font-size:9px;color:var(--muted);margin-left:3px}}

.section-title{{font-family:var(--font-head);font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
  color:var(--muted);margin-bottom:10px;padding-bottom:5px;border-bottom:1px solid var(--border)}}

.cl-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px;margin-bottom:20px}}
.cl-card{{background:var(--surface);border:1px solid var(--border);border-radius:3px;padding:10px 12px;
  border-left:3px solid var(--border);cursor:pointer;transition:background .12s}}
.cl-card:hover{{background:var(--surface2)}}
.cl-card-head{{display:flex;align-items:center;gap:7px;margin-bottom:5px}}
.cl-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.cl-name{{font-size:10px;font-weight:700;color:var(--text)}}
.cl-n{{margin-left:auto;font-family:var(--font-head);font-size:18px;font-weight:700}}
.cl-pct{{font-size:9px;color:var(--muted);margin-left:4px}}
.cl-desc{{font-size:9px;color:var(--muted);line-height:1.5;margin-bottom:7px}}
.cl-stats{{display:grid;grid-template-columns:1fr 1fr;gap:4px 12px;margin-bottom:7px}}
.cl-stat{{display:flex;flex-direction:column;gap:1px}}
.cl-stat span{{font-size:8px;letter-spacing:.08em;color:var(--muted);text-transform:uppercase}}
.cl-stat b{{font-size:11px;color:var(--text)}}
.cl-top{{font-size:8px;color:var(--muted);border-top:1px solid var(--border);padding-top:5px;margin-top:2px;line-height:1.5}}
.cl-top b{{color:var(--text)}}

.quad-wrap{{background:var(--surface);border:1px solid var(--border);border-radius:3px;padding:10px;margin-bottom:20px}}
.quad-wrap canvas{{cursor:crosshair}}
.ov-rank-head{{display:flex;align-items:center;gap:8px;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid var(--border);flex-wrap:wrap}}
.bar-chart{{display:flex;flex-direction:column;gap:3px}}
.bc-row{{display:flex;align-items:center;gap:8px;padding:2px 0}}
.bc-rank{{font-size:9px;color:var(--muted);width:20px;flex-shrink:0;text-align:right}}
.bc-dot{{width:6px;height:6px;border-radius:50%;flex-shrink:0}}
.bc-nm{{font-size:10px;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;width:200px;flex-shrink:0}}
.bc-bar-wrap{{flex:1;height:12px;background:var(--surface2);border-radius:1px;position:relative;overflow:hidden;min-width:40px}}
.bc-bar{{height:100%;border-radius:1px}}
.bc-val{{font-size:10px;color:var(--text);white-space:nowrap;width:58px;flex-shrink:0;text-align:right}}
.bc-sub{{font-size:9px;color:var(--muted);white-space:nowrap;width:60px;flex-shrink:0}}

/* ── Tab 2: Dispersão ── */
.t2-wrap{{flex:1;display:flex;flex-direction:column;min-height:0;overflow:hidden}}
.ctrl-bar{{display:flex;align-items:center;gap:9px;padding:7px 18px 5px;flex-shrink:0;flex-wrap:wrap;border-bottom:1px solid var(--border)}}
.lbl{{color:var(--muted);font-size:9px;letter-spacing:.13em;text-transform:uppercase;white-space:nowrap}}
select{{background:var(--surface2);color:var(--text);border:1px solid var(--border);
  font-family:var(--font-mono);font-size:10px;padding:4px 20px 4px 7px;border-radius:2px;
  cursor:pointer;outline:none;appearance:none;-webkit-appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='8' height='5'%3E%3Cpath d='M0 0l4 5 4-5z' fill='%23666'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 6px center}}
.vsep{{width:1px;height:16px;background:var(--border)}}
.sbtns{{display:flex;border:1px solid var(--border);border-radius:2px;overflow:hidden}}
.sbtn{{padding:3px 8px;font-family:var(--font-mono);font-size:9px;letter-spacing:.06em;text-transform:uppercase;
  background:none;color:var(--muted);border:none;cursor:pointer;transition:background .1s,color .1s}}
.sbtn.on{{background:var(--accent);color:#fff}}
.chips{{display:flex;gap:4px;flex-wrap:wrap}}
.chip{{display:flex;align-items:center;gap:4px;padding:3px 8px;border-radius:2px;
  cursor:pointer;font-size:9px;background:var(--surface2);transition:opacity .15s;white-space:nowrap;user-select:none}}
.chip.off{{opacity:.25}}
.cdot{{width:6px;height:6px;border-radius:50%;flex-shrink:0}}
.chip-n{{color:var(--muted);margin-left:1px}}
.search-wrap{{position:relative;flex-shrink:0;display:flex;align-items:center}}
.search-in{{background:var(--surface2);color:var(--text);border:1px solid var(--border);
  font-family:var(--font-mono);font-size:10px;padding:4px 28px 4px 8px;
  border-radius:2px;outline:none;width:180px;transition:border-color .15s,width .15s}}
.search-in:focus{{border-color:var(--accent);width:240px}}
.search-in::placeholder{{color:var(--muted)}}
.search-clr{{position:absolute;right:6px;background:none;border:none;color:var(--muted);
  cursor:pointer;font-size:13px;padding:0;display:none}}
.search-clr.show{{display:block}}
.search-clr:hover{{color:var(--text)}}

.canvas-wrap{{flex:1;position:relative;min-height:0;margin:0 18px 3px}}
.canvas-wrap canvas{{position:absolute;inset:0;width:100%!important;height:100%!important;cursor:crosshair}}
.status-bar{{padding:2px 18px 4px;flex-shrink:0;font-size:9px;color:var(--muted);letter-spacing:.04em}}
.status-bar b{{color:var(--text)}}

.det-panel{{flex-shrink:0;background:var(--surface);border-top:1px solid var(--border);
  overflow:hidden;transition:height .22s cubic-bezier(.4,0,.2,1);height:0}}
.det-panel.open{{height:280px}}
.det-inner{{height:100%;display:flex;flex-direction:column;padding:0 18px 10px}}
.det-head{{display:flex;align-items:center;gap:8px;padding:9px 0 7px;flex-shrink:0;
  border-bottom:1px solid var(--border);flex-wrap:wrap;row-gap:5px}}
.det-name{{font-family:var(--font-head);font-size:14px;font-weight:800;color:var(--text);word-break:break-word}}
.det-badge{{font-size:8px;letter-spacing:.1em;text-transform:uppercase;padding:2px 8px;border-radius:2px;color:#fff;white-space:nowrap;flex-shrink:0}}
.det-kpis{{display:flex;gap:12px;margin-left:auto;font-size:10px;color:var(--muted);flex-wrap:wrap;row-gap:3px}}
.det-kpis b{{color:var(--text)}}
.det-close{{background:none;border:none;color:var(--muted);cursor:pointer;font-size:18px;padding:2px 5px;border-radius:2px;flex-shrink:0}}
.det-close:hover{{color:var(--text);background:var(--border)}}
.det-tw{{flex:1;overflow-y:auto;min-height:0;margin-top:7px}}
.det-tw::-webkit-scrollbar{{width:3px}}
.det-tw::-webkit-scrollbar-thumb{{background:var(--border);border-radius:2px}}

/* Tables shared */
table{{width:100%;border-collapse:collapse}}
thead th{{position:sticky;top:0;z-index:1;background:var(--surface);
  padding:4px 8px;text-align:left;font-size:9px;letter-spacing:.08em;text-transform:uppercase;
  color:var(--muted);border-bottom:1px solid var(--border);white-space:nowrap;cursor:pointer;user-select:none}}
thead th:hover{{color:var(--text)}}
thead th.sa::after{{content:' ↑';color:var(--accent)}}
thead th.sd::after{{content:' ↓';color:var(--accent)}}
thead th.r{{text-align:right}}
tbody tr{{border-bottom:1px solid rgba(255,255,255,.035);transition:background .06s}}
tbody tr:hover{{background:var(--surface2)}}
td{{padding:5px 8px;vertical-align:middle;white-space:nowrap;font-size:11px;color:var(--text)}}
td.r{{text-align:right;font-variant-numeric:tabular-nums}}
td.dim{{color:var(--muted)}}
.badge{{font-size:8px;letter-spacing:.08em;text-transform:uppercase;padding:2px 7px;border-radius:2px;color:#fff;white-space:nowrap}}

/* ── Tab 3: Ranking ── */
.t3-wrap{{flex:1;display:flex;flex-direction:column;min-height:0;overflow:hidden}}
.rank-ctrl{{display:flex;align-items:center;gap:9px;padding:7px 18px 5px;flex-shrink:0;flex-wrap:wrap;border-bottom:1px solid var(--border)}}
.rank-count{{font-size:9px;color:var(--muted);margin-left:auto}}
.rank-tw{{flex:1;overflow-y:auto;min-height:0;padding:0 18px}}
.rank-tw::-webkit-scrollbar{{width:4px}}
.rank-tw::-webkit-scrollbar-thumb{{background:var(--border);border-radius:2px}}
.rank-tw thead th{{background:var(--bg)}}
.detail-row td{{padding:0}}
.detail-row-inner{{padding:8px 10px;background:rgba(255,255,255,.025);border-left:2px solid var(--accent)}}
.detail-row-inner table{{margin-top:4px}}
.detail-row-inner thead th{{background:rgba(255,255,255,.03)}}
tr.expanded{{background:rgba(224,90,43,.07)!important}}
.mini-bar{{display:inline-block;height:6px;border-radius:1px;vertical-align:middle;margin-right:4px;opacity:.7}}

#tip{{position:fixed;pointer-events:none;background:#111;border:1px solid #2a2a2a;
  border-radius:3px;padding:8px 12px;max-width:260px;z-index:999;
  opacity:0;transition:opacity .07s;line-height:1.8;font-size:10px}}
#tip.on{{opacity:1}}

/* ── Cluster panel (slide-in de baixo dos cards) ── */
.cl-panel{{display:none;background:var(--surface2);border:1px solid var(--border);border-radius:3px;
  padding:10px 14px;margin-bottom:16px}}
.cl-panel.open{{display:block}}
.cl-panel-head{{display:flex;align-items:center;gap:8px;margin-bottom:8px;
  padding-bottom:6px;border-bottom:1px solid var(--border)}}
.cl-panel-title{{font-family:var(--font-head);font-size:12px;font-weight:700}}
.cl-panel-close{{margin-left:auto;background:none;border:none;color:var(--muted);
  cursor:pointer;font-size:16px;padding:0 4px;border-radius:2px}}
.cl-panel-close:hover{{color:var(--text);background:var(--border)}}
.cl-prod-list{{display:flex;flex-direction:column;gap:2px;max-height:260px;overflow-y:auto}}
.cl-prod-list::-webkit-scrollbar{{width:3px}}
.cl-prod-list::-webkit-scrollbar-thumb{{background:var(--border);border-radius:2px}}
.cl-prod-row{{display:flex;align-items:center;gap:8px;padding:4px 6px;border-radius:2px;
  cursor:pointer;transition:background .08s}}
.cl-prod-row:hover{{background:var(--border)}}
.cl-prod-nm{{font-size:10px;color:var(--text);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.cl-prod-meta{{font-size:9px;color:var(--muted);white-space:nowrap}}

/* ── Obras drawer (Tab 1 bar chart click) ── */
.obras-drawer{{display:none;background:var(--surface);border:1px solid var(--border);border-radius:3px;
  padding:10px 14px;margin-top:10px}}
.obras-drawer.open{{display:block}}
.obras-drawer-head{{display:flex;align-items:center;gap:8px;margin-bottom:8px;
  padding-bottom:6px;border-bottom:1px solid var(--border)}}
.obras-drawer-title{{font-family:var(--font-head);font-size:11px;font-weight:700;
  flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.obras-drawer-close{{background:none;border:none;color:var(--muted);cursor:pointer;
  font-size:16px;padding:0 4px;border-radius:2px}}
.obras-drawer-close:hover{{color:var(--text);background:var(--border)}}
.obras-drawer-tw{{max-height:240px;overflow-y:auto}}
.obras-drawer-tw::-webkit-scrollbar{{width:3px}}
.obras-drawer-tw::-webkit-scrollbar-thumb{{background:var(--border);border-radius:2px}}

/* ── Portfolio size selector ── */
.quad-ctrl{{display:flex;align-items:center;gap:8px;margin-bottom:6px;flex-wrap:wrap}}
.quad-ctrl .lbl{{color:var(--muted);font-size:9px;letter-spacing:.1em;text-transform:uppercase}}
.bc-row{{cursor:pointer;border-radius:2px;padding:2px 4px;transition:background .08s}}
.bc-row:hover{{background:var(--surface2)}}
</style>
</head>
<body>

<div class="hdr">
  <span class="hdr-t">Produtoras Independentes BR</span>
  <span class="hdr-s">ROI doméstico e internacional — obras com fomento público e bilheteria {year_min}–{year_max}</span>
</div>

<div class="tab-bar">
  <button class="tab-btn active" onclick="switchTab(0)">Visão Geral</button>
  <button class="tab-btn" onclick="switchTab(2)">Ranking</button>
  <button class="tab-btn" onclick="switchTab(3)">Filmes</button>
</div>

<!-- ═══ TAB 1: VISÃO GERAL ═══ -->
<div class="tab-panel main-tab active" id="tab0">
  <div class="t1-scroll">
    <div class="kpi-bar" id="kpi-bar"></div>
    <div class="section-title">Distribuição por Cluster</div>
    <div class="cl-grid" id="cl-grid"></div>
    <div class="section-title">Matriz de Portfolio — ROI Dom. vs ROI Intl</div>
    <div class="quad-wrap">
      <div class="quad-ctrl">
        <span class="lbl">Tamanho</span>
        <select id="quad-size" onchange="drawQuad()" style="font-size:10px;padding:2px 18px 2px 6px">
          <option value="n">N Obras</option>
          <option value="inv_def">Investimento Deflac.</option>
          <option value="pub_def">Bilheteria Deflac.</option>
          <option value="rec_def">Receita Deflac.</option>
          <option value="np_uniq">Países Únicos</option>
        </select>
        <span id="quad-size-legend" style="font-size:9px;color:var(--muted);margin-left:4px"></span>
      </div>
      <canvas id="quad-canvas" width="800" height="260" style="width:100%;height:260px"></canvas>
    </div>
    <div id="cl-panel" class="cl-panel">
      <div class="cl-panel-head">
        <div class="cl-dot" id="cl-panel-dot" style="width:10px;height:10px;border-radius:50%;flex-shrink:0"></div>
        <span class="cl-panel-title" id="cl-panel-title"></span>
        <button class="cl-panel-close" onclick="closeClusterPanel()">×</button>
      </div>
      <div class="cl-prod-list" id="cl-prod-list"></div>
    </div>
    <div class="ov-rank-head">
      <div class="section-title" style="margin:0;padding:0;border:none">Ranking Produtoras</div>
      <span class="lbl">Ordenar</span>
      <select id="ov-sort" onchange="renderOvRank()"></select>
      <div class="search-wrap" style="margin-left:auto">
        <input class="search-in" id="ov-search" placeholder="Buscar…" oninput="renderOvRank()" style="width:160px">
        <button class="search-clr" id="ov-clr" onclick="document.getElementById('ov-search').value='';document.getElementById('ov-clr').classList.remove('show');renderOvRank()">×</button>
      </div>
      <span id="ov-count" style="font-size:9px;color:var(--muted)"></span>
    </div>
    <div class="bar-chart" id="bar-chart"></div>
    <div id="obras-drawer" class="obras-drawer">
      <div class="obras-drawer-head">
        <span class="obras-drawer-title" id="obras-drawer-title"></span>
        <button class="obras-drawer-close" onclick="closeObrasDrawer()">×</button>
      </div>
      <div class="obras-drawer-tw">
        <table>
          <thead id="obras-drawer-thead"><tr></tr></thead>
          <tbody id="obras-drawer-tbody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ═══ TAB 2: DISPERSÃO (Plotly) ═══ -->
<div class="tab-panel" id="tab-dispersao-hidden" style="display:none">
  <div style="padding:14px 18px">
    <div class="ctrl-bar" style="margin-bottom:12px">
      <span class="lbl">Eixo X</span>
      <select id="sx" onchange="renderScatterPlotly()"></select>
      <span class="lbl">Eixo Y</span>
      <select id="sy" onchange="renderScatterPlotly()"></select>
      <div class="vsep"></div>
      <span class="lbl">Escala</span>
      <div class="sbtns">
        <button class="sbtn on" id="s-lin" onclick="setScalePlotly('lin')">Lin</button>
        <button class="sbtn" id="s-log" onclick="setScalePlotly('log')">Log</button>
      </div>
    </div>
    <div id="scatter-plotly" style="width:100%;height:560px"></div>
  </div>
</div>

<!-- ═══ TAB 3: RANKING ═══ -->
<div class="tab-panel main-tab" id="tab1">
  <div class="t3-wrap">
    <div class="rank-ctrl">
      <span class="lbl">Ordenar</span>
      <select id="rank-sort" onchange="renderRanking()"></select>
      <div class="vsep"></div>
      <div class="chips" id="rank-chips"></div>
      <div class="vsep"></div>
      <div class="search-wrap">
        <input class="search-in" id="r-search" placeholder="Buscar produtora…" oninput="renderRanking()">
        <button class="search-clr" id="r-clr" onclick="document.getElementById('r-search').value='';renderRanking()">×</button>
      </div>
      <span class="rank-count" id="rank-count"></span>
    </div>
    <div class="rank-tw">
      <table>
        <thead id="rank-thead"></thead>
        <tbody id="rank-tbody"></tbody>
      </table>
    </div>
  </div>
</div>

<div class="tab-panel main-tab" id="tab2">
  <div class="t3-wrap">
    <div class="rank-ctrl">
      <span class="lbl">Ordenar</span>
      <select id="film-sort" onchange="renderFilms()"></select>
      <div class="vsep"></div>
      <span class="lbl" style="margin-left:4px">Cluster</span>
      <div class="chips" id="film-cl-chips"></div>
      <div class="vsep"></div>
      <span class="lbl">Crítica mín. fontes</span>
      <select id="film-crn" onchange="renderFilms()">
        <option value="0">Todos</option>
        <option value="1">≥ 1 fonte</option>
        <option value="2" selected>≥ 2 fontes</option>
        <option value="3">≥ 3 fontes</option>
      </select>
      <div class="vsep"></div>
      <div class="search-wrap">
        <input class="search-in" id="film-search" placeholder="Buscar título ou produtora…" oninput="renderFilms()">
        <button class="search-clr" id="film-clr" onclick="document.getElementById('film-search').value='';renderFilms()">×</button>
      </div>
      <span class="rank-count" id="film-count"></span>
    </div>
    <div class="rank-tw">
      <table>
        <thead id="film-thead"></thead>
        <tbody id="film-tbody"></tbody>
      </table>
    </div>
  </div>
</div>

<div id="tip"></div>

<script id="prod-data">
const PROD = {json_data};
const THRESHOLDS = {thresholds_json};
</script>
<script>
// ── Cluster config ─────────────────────────────────────────────────────────
const CL = {{
  duplo:       {{c:'#f5c842', l:'Duplo Retorno',         desc:'Receita total ≥ R$ 2,5M e ROI Internacional máx ≥ 13 (cauda superior da distribuição)'}},
  dom:         {{c:'#5b8cff', l:'Retorno Doméstico',     desc:'Receita total ≥ R$ 10M, ou ROI doméstico > 0,6 com renda ≥ R$ 2,5M, sem ROI Internacional máx ≥ 13'}},
  intl:        {{c:'#ff80b0', l:'Retorno Internacional', desc:'ROI Internacional máx ≥ 13, sem critério de Retorno Doméstico'}},
  sem_retorno: {{c:'#f87171', l:'Fomento Baixo Retorno',  desc:'Investimento total > R$ 5M e ROI Internacional máx < 13'}},
  pequeno:     {{c:'#7b849a', l:'Pequeno Porte',         desc:'Investimento total ≤ R$ 5M ou receita sem critério de retorno'}},
}};
const CL_ORDER = ['duplo','dom','intl','sem_retorno','pequeno'];


// ── Format helpers ─────────────────────────────────────────────────────────
function fmtMoney(v) {{
  if(v>=1e9) return 'R$ '+(v/1e9).toFixed(1)+'B';
  if(v>=1e6) return 'R$ '+(v/1e6).toFixed(1)+'M';
  if(v>=1e3) return 'R$ '+(v/1e3).toFixed(0)+'K';
  return v>0 ? 'R$ '+Math.round(v).toLocaleString('pt-BR') : '—';
}}
function fmtRatio(v) {{ return v>0 ? v.toFixed(2)+'×' : '—'; }}
function fmtInt(v)   {{ return v>0 ? String(v) : '—'; }}
function fmtVal(v, fmt) {{
  if(fmt==='money') return fmtMoney(v);
  if(fmt==='ratio') return fmtRatio(v);
  return fmtInt(v);
}}
function getVal(r, key) {{
  return r[key] !== undefined ? (r[key] || 0) : 0;
}}
function hex2rgba(hex, a) {{
  const r=parseInt(hex.slice(1,3),16), g=parseInt(hex.slice(3,5),16), b=parseInt(hex.slice(5,7),16);
  return `rgba(${{r}},${{g}},${{b}},${{a}})`;
}}

// ── Tab switching ──────────────────────────────────────────────────────────
const tabBtns = document.querySelectorAll('.tab-btn');
const tabPanels = document.querySelectorAll('.main-tab');
let activeTab = 0;
function switchTab(i) {{
  tabBtns.forEach((b,j) => b.classList.toggle('active', j===i));
  tabPanels.forEach((p,j) => p.classList.toggle('active', j===i));
  activeTab = i;
  if(i===2) renderRanking();
  if(i===3) renderFilms();
}}

// ══════════════════════════════════════════════════════════════════════════
// TAB 1: VISÃO GERAL
// ══════════════════════════════════════════════════════════════════════════
const OV_SORT_OPTS = [
  {{k:'rim',     l:'ROI Intl Máx'}},
  {{k:'ria',     l:'ROI Intl Médio'}},
  {{k:'roi_def', l:'ROI Dom. Deflac.'}},
  {{k:'rda',     l:'ROI Dom. Proporcional'}},
  {{k:'rdm',     l:'ROI Dom. Mediano'}},
  {{k:'inv_def', l:'Investimento Deflac.'}},
  {{k:'inv',     l:'Investimento Nominal'}},
  {{k:'n',       l:'N Obras'}},
  {{k:'pub_def', l:'Bilheteria Deflac.'}},
  {{k:'pub',     l:'Bilheteria Nominal'}},
  {{k:'rec_def', l:'Receita Deflac.'}},
  {{k:'rec',     l:'Receita Nominal'}},
  {{k:'np_uniq', l:'Países Únicos'}},
];

// Formato por chave métrica — usado em renderOvRank e drawQuad
const AXES = {{
  rim:     {{fmt:'ratio'}}, ria:     {{fmt:'ratio'}},
  roi_def: {{fmt:'ratio'}}, rda:     {{fmt:'ratio'}}, rdm: {{fmt:'ratio'}},
  inv_def: {{fmt:'money'}}, inv:     {{fmt:'money'}},
  pub_def: {{fmt:'money'}}, pub:     {{fmt:'money'}},
  rec_def: {{fmt:'money'}}, rec:     {{fmt:'money'}},
  n:       {{fmt:'int'}},   np_uniq: {{fmt:'int'}},
}};

function _med(arr) {{
  if(!arr.length) return 0;
  const s=[...arr].sort((a,b)=>a-b);
  return s[Math.floor(s.length/2)];
}}
function _sum(arr) {{ return arr.reduce((a,b)=>a+b,0); }}

function buildTab1() {{
  const total = PROD.length;
  const totalObras = PROD.reduce((s,r) => s+r.n, 0);
  const totalInv = PROD.reduce((s,r) => s+r.inv, 0);
  const rdaVals = PROD.filter(r=>r.rda>0).map(r=>r.rda);
  const rdaMed = _med(rdaVals);
  const rimValsAll = PROD.filter(r=>r.rim>=13).map(r=>r.rim);
  const rimMedAll = _med(rimValsAll);
  const riaVals = PROD.filter(r=>r.ria>0).map(r=>r.ria);
  const riaAvg = riaVals.length ? _sum(riaVals)/riaVals.length : 0;
  const totalProd = total;
  const withIntl = PROD.filter(r=>r.rim>=13).length;

  // KPI bar
  const totalBilhDef = PROD.reduce((s,r) => s+(r.pub_def||0), 0);
  const totalInvDef  = PROD.reduce((s,r) => s+(r.inv_def||0), 0);
  const totalRecDef  = PROD.reduce((s,r) => s+(r.rec_def||0), 0);
  const roiDefGlobal = totalInvDef>=1000 ? totalRecDef/totalInvDef : 0;
  const kpiBar = document.getElementById('kpi-bar');
  kpiBar.innerHTML = [
    {{l:'Produtoras', v:totalProd, fmt:'int'}},
    {{l:'Obras Totais', v:totalObras, fmt:'int'}},
    {{l:'Invest. Deflac. Total', v:totalInvDef, fmt:'money'}},
    {{l:'Bilheteria Deflac.', v:totalBilhDef, fmt:'money'}},
    {{l:'ROI Deflac. Global', v:roiDefGlobal, fmt:'ratio'}},
    {{l:'ROI Intl Mediano', v:rimMedAll, fmt:'ratio'}},
    {{l:'Com ROI Intl ≥13', v:withIntl, fmt:'int'}},
  ].map(k => `<div class="kpi">
    <div class="kpi-l">${{k.l}}</div>
    <div class="kpi-v">${{fmtVal(k.v, k.fmt)}}</div>
  </div>`).join('');

  // Per-cluster stats for enriched cards
  const clStats = {{}};
  CL_ORDER.forEach(k => {{
    const grp = PROD.filter(r=>r.cl===k);
    const gA0  = grp.filter(r=>r.a0>0).map(r=>r.a0);
    const gA1  = grp.filter(r=>r.a1>0).map(r=>r.a1);
    const topRim = [...grp].sort((a,b)=>b.rim-a.rim).slice(0,3);
    // Aggregate cluster-level ROI deflacionado
    const sumInvDef  = grp.reduce((s,r)=>s+(r.inv_def||0),0);
    const sumFsaDef  = grp.reduce((s,r)=>s+(r.inv_fsa_d||0),0);
    const sumRenDef  = grp.reduce((s,r)=>s+(r.inv_ren_d||0),0);
    const sumRecDef  = grp.reduce((s,r)=>s+(r.rec_def||0),0);
    const sumInv     = grp.reduce((s,r)=>s+(r.inv||0),0);
    const sumObras   = grp.reduce((s,r)=>s+r.n,0);
    const roiDefCluster = sumInvDef>=1000 ? sumRecDef/sumInvDef : 0;
    // Países únicos do cluster = união de todas as listas
    const clPaisesSet = new Set(grp.flatMap(r=>r.paises||[]));
    clStats[k] = {{
      roiDef: roiDefCluster,
      sumInvDef,
      sumFsaDef,
      sumRenDef,
      sumRecDef,
      sumInv,
      totalObras: sumObras,
      totalProd: grp.length,
      npCluster: clPaisesSet.size,
      a0min: gA0.length ? Math.min(...gA0) : 0,
      a1max: gA1.length ? Math.max(...gA1) : 0,
      topRim,
    }};
  }});

  // Cluster cards
  const clGrid = document.getElementById('cl-grid');
  clGrid.innerHTML = CL_ORDER.map(k => {{
    const info=CL[k], st=clStats[k], cnt=st.totalProd;
    const pct = total ? (cnt/total*100).toFixed(0) : 0;
    const topNames = st.topRim.map(r=>{{
      const nm=r.nm.length>26?r.nm.slice(0,24)+'…':r.nm;
      return `<b>${{nm}}</b> ${{fmtRatio(r.rim)}}`;
    }}).join('<br>');
    const anos = st.a0min && st.a1max ? `${{st.a0min}}\u2013${{st.a1max}}` : '\u2014';
    return `<div class="cl-card" style="border-left-color:${{info.c}}" onclick="openClusterPanel('${{k}}')">
      <div class="cl-card-head">
        <div class="cl-dot" style="background:${{info.c}}"></div>
        <span class="cl-name">${{info.l}}</span>
        <span class="cl-n" style="color:${{info.c}};font-size:16px">${{cnt}}</span>
        <span class="cl-pct">${{pct}}%</span>
      </div>
      <div class="cl-desc">${{info.desc}}</div>
      <div class="cl-stats">
        <div class="cl-stat"><span>ROI Deflac. Cluster</span><b style="color:${{info.c}}">${{st.roiDef>0?st.roiDef.toFixed(2)+'&times;':'\u2014'}}</b></div>
        <div class="cl-stat"><span>Renda Est. Deflac.</span><b>${{fmtMoney(st.sumRecDef)}}</b></div>
        <div class="cl-stat"><span>FSA Deflac.</span><b>${{fmtMoney(st.sumFsaDef)}}</b></div>
        <div class="cl-stat"><span>Renu&#769;ncia Deflac.</span><b>${{fmtMoney(st.sumRenDef)}}</b></div>
        <div class="cl-stat"><span>Obras / Produtoras</span><b>${{st.totalObras}} / ${{st.totalProd}}</b></div>
        <div class="cl-stat"><span>Pa&#237;ses &Uacute;nicos</span><b>${{st.npCluster>0?st.npCluster+'p':'\u2014'}}</b></div>
        <div class="cl-stat"><span>Per&#237;odo</span><b>${{anos}}</b></div>
      </div>
      <div style="font-size:8px;color:var(--muted);margin-top:4px;opacity:.6">clique para ver produtoras</div>
    </div>`;
  }}).join('');

  // Sort selector
  const sel = document.getElementById('ov-sort');
  if(!sel.options.length) {{
    OV_SORT_OPTS.forEach(o => {{ sel.innerHTML += `<option value="${{o.k}}">${{o.l}}</option>`; }});
    sel.value = 'rim';
  }}
  renderOvRank();
  drawQuad();
}}

function drawQuad() {{
  const qc = document.getElementById('quad-canvas');
  const qctx = qc.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const W = qc.clientWidth || 800, H = qc.clientHeight || 260;
  qc.width = W * dpr; qc.height = H * dpr;
  qctx.scale(dpr, dpr);

  const PAD_Q = {{l:54, r:20, t:22, b:38}};
  const pw = W - PAD_Q.l - PAD_Q.r, ph = H - PAD_Q.t - PAD_Q.b;

  qctx.clearRect(0,0,W,H);

  // Determine axis ranges
  // X: receita total (rec), escala log — divide em rec >= BILH_THRESHOLD_HIGH
  // Y: ROI intl máx (rim), linear com zona negativa para rim=0
  const recVals = PROD.filter(r=>r.rec>0).map(r=>r.rec);
  const rimVals = PROD.filter(r=>r.rim>=13).map(r=>r.rim);
  const xlo = recVals.length > 0 ? Math.max(Math.min(...recVals) * 0.6, 1000) : 1000;
  const xhi = recVals.length > 0 ? Math.max(...recVals) * 1.2 : 100000000;
  const yhi = rimVals.length > 0 ? Math.max(...rimVals, 5) * 1.12 : 20;
  // Zona negativa abaixo de rim=0 para acomodar empresas sem ROI intl mensurável (rim=0)
  const yNegFrac = 0.15; // 15% do espaço visual reservado para rim=0
  const ylo = -yhi * (yNegFrac / (1 - yNegFrac));
  // Posição Y onde as bolhas rim=0 ficam plotadas (centro da zona negativa)
  const yRim0 = ylo * 0.55;

  function toQX(v) {{
    const lv = v > 0 ? Math.log10(v) : Math.log10(xlo);
    const ll = Math.log10(xlo), lh = Math.log10(xhi);
    return PAD_Q.l + ((lv - ll) / (lh - ll)) * pw;
  }}
  function toQY(v) {{
    return PAD_Q.t + ph - ((v - ylo) / (yhi - ylo)) * ph;
  }}

  // Divisores alinhados com a lógica de cluster:
  // X: receita >= R$2,5M (BILH_THRESHOLD_HIGH)
  // Y: rim == 0  (fronteira visual entre zero e ROI intl mensurável)
  const txLine = toQX(THRESHOLDS.bilh_high);
  const tyLine = toQY(0);

  const quadColors = [
    {{x:PAD_Q.l, y:PAD_Q.t, w:txLine-PAD_Q.l, h:tyLine-PAD_Q.t, c:'rgba(79,163,224,0.06)'}},   // top-left: intl
    {{x:txLine, y:PAD_Q.t, w:PAD_Q.l+pw-txLine, h:tyLine-PAD_Q.t, c:'rgba(245,200,66,0.06)'}},   // top-right: duplo
    {{x:txLine, y:tyLine, w:PAD_Q.l+pw-txLine, h:PAD_Q.t+ph-tyLine, c:'rgba(240,144,32,0.06)'}}, // bottom-right: comercial/dom
    {{x:PAD_Q.l, y:tyLine, w:txLine-PAD_Q.l, h:PAD_Q.t+ph-tyLine, c:'rgba(74,74,96,0.08)'}},    // bottom-left: baixo/pequeno
  ];
  quadColors.forEach(q => {{
    qctx.fillStyle = q.c;
    qctx.fillRect(q.x, q.y, q.w, q.h);
  }});

  // Linha separadora da zona rim=0 (fundo hachurado leve)
  qctx.fillStyle = 'rgba(30,30,30,0.18)';
  qctx.fillRect(PAD_Q.l, tyLine, pw, PAD_Q.t + ph - tyLine);

  // Grid lines (apenas região rim>=13)
  qctx.strokeStyle = '#1e1e1e'; qctx.lineWidth = 1;
  const xTicks = [10000,50000,100000,500000,1000000,2500000,5000000,10000000,50000000,100000000];
  xTicks.forEach(v => {{
    if(v < xlo || v > xhi) return;
    const px = toQX(v);
    qctx.beginPath(); qctx.moveTo(px, PAD_Q.t); qctx.lineTo(px, PAD_Q.t+ph); qctx.stroke();
  }});
  const yStep = yhi / 5;
  for(let i=0; i<=5; i++) {{
    const yv = i * yStep;
    const py = toQY(yv);
    qctx.beginPath(); qctx.moveTo(PAD_Q.l, py); qctx.lineTo(PAD_Q.l+pw, py); qctx.stroke();
  }}

  // Divisores de cluster (linhas tracejadas)
  qctx.setLineDash([4,3]); qctx.lineWidth = 1.5;
  qctx.strokeStyle = 'rgba(200,200,200,0.5)';
  qctx.beginPath(); qctx.moveTo(txLine, PAD_Q.t); qctx.lineTo(txLine, PAD_Q.t+ph); qctx.stroke();
  qctx.beginPath(); qctx.moveTo(PAD_Q.l, tyLine); qctx.lineTo(PAD_Q.l+pw, tyLine); qctx.stroke();
  qctx.setLineDash([]);

  // Axes
  qctx.strokeStyle = '#333'; qctx.lineWidth = 1;
  qctx.beginPath(); qctx.moveTo(PAD_Q.l, PAD_Q.t); qctx.lineTo(PAD_Q.l, PAD_Q.t+ph); qctx.stroke();
  qctx.beginPath(); qctx.moveTo(PAD_Q.l, PAD_Q.t+ph); qctx.lineTo(PAD_Q.l+pw, PAD_Q.t+ph); qctx.stroke();

  // Axis labels
  function fmtRec(v) {{
    if(v >= 1e8) return (v/1e6).toFixed(0)+'M';
    if(v >= 1e6) return (v/1e6).toFixed(1).replace('.0','')+'M';
    if(v >= 1e3) return (v/1e3).toFixed(0)+'k';
    return v.toFixed(0);
  }}
  qctx.fillStyle='#444'; qctx.font="9px 'DM Mono',monospace";
  xTicks.forEach(v => {{
    if(v < xlo || v > xhi) return;
    const px = toQX(v);
    qctx.textAlign='center'; qctx.textBaseline='top';
    qctx.fillText(fmtRec(v), px, PAD_Q.t+ph+4);
  }});
  for(let i=0; i<=5; i++) {{
    const yv = i * yStep;
    const py = toQY(yv);
    qctx.textAlign='right'; qctx.textBaseline='middle';
    qctx.fillText(yv.toFixed(1)+'×', PAD_Q.l-4, py);
  }}
  // Label zona rim=0
  qctx.fillStyle='rgba(120,120,140,0.7)'; qctx.font="8px 'DM Mono',monospace";
  qctx.textAlign='left'; qctx.textBaseline='middle';
  qctx.fillText('ROI intl < 13', PAD_Q.l+4, toQY(yRim0));

  qctx.fillStyle='#555'; qctx.font="9px 'DM Mono',monospace";
  qctx.textAlign='center'; qctx.textBaseline='bottom';
  qctx.fillText('Receita Total (R$, log)', PAD_Q.l+pw/2, H-2);
  qctx.save(); qctx.translate(11, PAD_Q.t+ph/2); qctx.rotate(-Math.PI/2);
  qctx.textAlign='center'; qctx.textBaseline='middle';
  qctx.fillText('ROI Intl Máx', 0, 0); qctx.restore();

  // Quadrant corner labels (apenas zona rim>=13)
  qctx.font="8px 'DM Mono',monospace";
  const qlabels = [
    {{x:PAD_Q.l+4, y:PAD_Q.t+4, a:'left', b:'top', t:'ROI Intl', c:'rgba(79,163,224,0.7)'}},
    {{x:PAD_Q.l+pw-4, y:PAD_Q.t+4, a:'right', b:'top', t:'Duplo Retorno', c:'rgba(245,200,66,0.7)'}},
    {{x:PAD_Q.l+pw-4, y:tyLine-4, a:'right', b:'bottom', t:'ROI Comercial', c:'rgba(240,144,32,0.5)'}},
    {{x:PAD_Q.l+4, y:tyLine-4, a:'left', b:'bottom', t:'Baixo Retorno', c:'rgba(100,100,120,0.5)'}},
  ];
  qlabels.forEach(q => {{
    qctx.fillStyle=q.c; qctx.textAlign=q.a; qctx.textBaseline=q.b;
    qctx.fillText(q.t, q.x, q.y);
  }});

  // Draw dots — 4th dimension: bubble size
  const szEl = document.getElementById('quad-size');
  const szKey = szEl ? szEl.value : 'n';
  const szVals = PROD.map(r=>r[szKey]||0);
  const szMax = Math.max(...szVals, 1);
  const MIN_R = 4, MAX_R = 20;  // range amplo para tornar a 4ª dimensão visível

  const quadPts = [];
  PROD.forEach(r => {{
    // X: receita total (log) — determina dom vs intl/duplo/pequeno
    const xv = r.rec > 0 ? r.rec : xlo;
    const cx = toQX(xv);
    // Y: rim<13 -> zona negativa; rim>=13 -> posição proporcional
    const cy = toQY(r.rim >= 13 ? r.rim : yRim0);
    const clr = (CL[r.cl]||{{c:'#666'}}).c;
    const alpha = r.cl==='pequeno' ? 0.25 : 0.6;
    const szV = r[szKey] || 0;
    const radius = r.cl==='pequeno'
      ? 2
      : MIN_R + (Math.sqrt(szV) / Math.sqrt(szMax)) * (MAX_R - MIN_R);
    qctx.beginPath(); qctx.arc(cx, cy, radius, 0, Math.PI*2);
    qctx.fillStyle = hex2rgba(clr, alpha); qctx.fill();
    if(r.cl!=='pequeno') {{
      qctx.strokeStyle = hex2rgba(clr, 0.8); qctx.lineWidth=1; qctx.stroke();
    }}
    quadPts.push({{r, cx, cy, radius}});
  }});

  // Legenda tamanho da bolha
  const szLegEl = document.getElementById('quad-size-legend');
  if(szLegEl) {{
    const szMin = Math.min(...szVals.filter(v=>v>0), 0);
    const axSz = AXES[szKey] || {{fmt:'int'}};
    szLegEl.textContent = szMax>0
      ? `\u2218 min\u2248${{fmtVal(szMin,axSz.fmt)}} \u25cf max\u2248${{fmtVal(szMax,axSz.fmt)}}`
      : '';
  }}

  // Mousemove tooltip
  qc._pts = quadPts;
  if(!qc._hasListener) {{
    qc._hasListener = true;
    qc.addEventListener('mousemove', e => {{
      const rect = qc.getBoundingClientRect();
      const mx = (e.clientX - rect.left) * (qc.width / rect.width / dpr);
      const my = (e.clientY - rect.top)  * (qc.height / rect.height / dpr);
      let best=-1, bestD=16;
      (qc._pts||[]).forEach((p,i) => {{
        const d=Math.hypot(p.cx-mx, p.cy-my);
        const thresh = Math.max(p.radius||4, 8);
        if(d<thresh && d<bestD) {{ bestD=d; best=i; }}
      }});
      if(best>=0) {{
        const p = (qc._pts||[])[best];
        showTipQuad(e, p.r);
      }} else hideTip();
    }});
    qc.addEventListener('mouseleave', hideTip);
  }}
}}

const tip = document.getElementById('tip');
function moveTip(e) {{ tip.style.left=(e.clientX+14)+'px'; tip.style.top=(e.clientY+14)+'px'; }}
function hideTip() {{ tip.classList.remove('on'); }}

function showTipQuad(e, r) {{
  tip.innerHTML = `<b style="color:#ddd8cc">${{r.nm}}</b><br>
    <span style="color:${{(CL[r.cl]||{{c:'#888'}}).c}}">${{(CL[r.cl]||{{l:r.cl}}).l}}</span><br>
    Obras: ${{r.n}} · Inv: ${{fmtMoney(r.inv)}}<br>
    ROI Dom: ${{fmtRatio(r.rda)}} · ROI Intl: ${{fmtRatio(r.rim)}}<br>
    Bilheteria Deflac: ${{fmtMoney(r.pub_def||0)}}<br>
    Países Únicos: ${{r.np_uniq||0}}`;
  tip.classList.add('on');
  moveTip(e);
}}

// ── Cluster panel ──────────────────────────────────────────────────────────
let _openCluster = null;
function openClusterPanel(clKey) {{
  if(_openCluster === clKey) {{ closeClusterPanel(); return; }}
  _openCluster = clKey;
  const info = CL[clKey] || {{l:clKey, c:'#888'}};
  const panel = document.getElementById('cl-panel');
  document.getElementById('cl-panel-dot').style.background = info.c;
  document.getElementById('cl-panel-title').textContent = info.l;
  const prods = [...PROD.filter(r=>r.cl===clKey)].sort((a,b)=>b.rda-a.rda);
  const list = document.getElementById('cl-prod-list');
  list.innerHTML = prods.map(r => {{
    const nm = r.nm.length>40 ? r.nm.slice(0,38)+'\u2026' : r.nm;
    const idx = PROD.indexOf(r);
    return `<div class="cl-prod-row" onclick="openObrasDrawerByIdx(${{idx}})">
      <span class="cl-prod-nm" title="${{r.nm}}">${{nm}}</span>
      <span class="cl-prod-meta">${{r.n}} obras &middot; ${{fmtRatio(r.rda)}} dom &middot; ${{fmtRatio(r.rim)}} intl</span>
    </div>`;
  }}).join('');
  panel.classList.add('open');
  panel.scrollIntoView({{behavior:'smooth', block:'nearest'}});
}}
function closeClusterPanel() {{
  _openCluster = null;
  document.getElementById('cl-panel').classList.remove('open');
}}

// ── Obras drawer (Tab 1) ───────────────────────────────────────────────────
const OBR_COLS = [
  {{k:'td',  l:'Título',          r:false, fmt:null}},
  {{k:'a',   l:'Ano',             r:true,  fmt:'int'}},
  {{k:'bd',  l:'Bilh. Deflac.',   r:true,  fmt:'money'}},
  {{k:'rec', l:'Renda Total Def.',r:true,  fmt:'money'}},
  {{k:'rd',  l:'ROI Dom.',        r:true,  fmt:'ratio'}},
  {{k:'ri',  l:'ROI Intl.',       r:true,  fmt:'ratio'}},
  {{k:'np',  l:'Países',          r:true,  fmt:'int'}},
  {{k:'inv', l:'Investimento',    r:true,  fmt:'money'}},
  {{k:'cat', l:'Categoria',       r:false, fmt:null}},
];
const DET_COLS = OBR_COLS;
function openObrasDrawerByIdx(idx) {{
  if(idx<0 || idx>=PROD.length) return;
  openObrasDrawer(PROD[idx]);
}}
function openObrasDrawer(r) {{
  if(!r) return;
  const drawer = document.getElementById('obras-drawer');
  document.getElementById('obras-drawer-title').textContent = r.nm;
  const obras = [...r.obras].sort((a,b)=>b.rd-a.rd);
  document.getElementById('obras-drawer-thead').innerHTML =
    '<tr>'+OBR_COLS.map(c=>`<th class="${{c.r?'r':''}}">${{c.l}}</th>`).join('')+'</tr>';
  document.getElementById('obras-drawer-tbody').innerHTML =
    obras.map(o=>'<tr>'+OBR_COLS.map(c=>{{
      const v=o[c.k]; const d=c.fmt?fmtVal(v,c.fmt):(v||'—');
      return `<td class="${{c.r?'r':''}}">${{d}}</td>`;
    }}).join('')+'</tr>').join('');
  drawer.classList.add('open');
  drawer.scrollIntoView({{behavior:'smooth', block:'nearest'}});
}}
function closeObrasDrawer() {{
  document.getElementById('obras-drawer').classList.remove('open');
}}

function renderOvRank() {{
  const sk = document.getElementById('ov-sort').value;
  const q  = document.getElementById('ov-search').value.trim().toLowerCase();
  document.getElementById('ov-clr').classList.toggle('show', q.length>0);
  let data = q ? PROD.filter(r=>r.nm.toLowerCase().includes(q)) : [...PROD];
  data.sort((a,b) => sk==='nm'?a.nm.localeCompare(b.nm):((b[sk]||0)-(a[sk]||0)));
  data = data.slice(0, q ? data.length : 30);
  document.getElementById('ov-count').textContent = data.length + ' produtoras';
  const axInfo = AXES[sk] || {{fmt:'ratio'}};
  const maxV = data.length ? Math.max(...data.map(r=>r[sk]||0), 0.001) : 1;
  const bc = document.getElementById('bar-chart');
  bc.innerHTML = data.map((r,i) => {{
    const v = r[sk] || 0;
    const pct = maxV>0 ? (v/maxV*100).toFixed(1) : 0;
    const clr = (CL[r.cl]||{{c:'#666'}}).c;
    const nm = r.nm.length>32 ? r.nm.slice(0,30)+'…' : r.nm;
    const sub = sk==='rim'||sk==='ria' ? `Dom.${{fmtRatio(r.rda)}}` :
                sk==='rda'||sk==='rdm' ? `Intl.${{fmtRatio(r.rim)}}` :
                sk==='inv'||sk==='rec'||sk==='pub'||sk==='pub_def' ? `${{r.n}} obras` :
                sk==='n' ? `Inv.${{fmtMoney(r.inv)}}` : '';
    const prodIdx = PROD.indexOf(r);
    const tipoBadge = r.tipo==='captador' ? `<span style="font-size:8px;padding:2px 6px;border-radius:2px;background:#2a1a0a;color:#e09a4f;border:1px solid #5a3a10;margin-left:6px;flex-shrink:0;letter-spacing:.06em;font-weight:600">DISTRIBUIDORA</span>` : '';
    return `<div class="bc-row" onclick="openObrasDrawerByIdx(${{prodIdx}})" title="Clique para ver obras">
      <div class="bc-rank">${{i+1}}</div>
      <div class="bc-dot" style="background:${{clr}}"></div>
      <div class="bc-nm" title="${{r.nm}}" style="display:flex;align-items:center">${{nm}}${{tipoBadge}}</div>
      <div class="bc-bar-wrap">
        <div class="bc-bar" style="width:${{pct}}%;background:${{clr}};opacity:.75"></div>
      </div>
      <div class="bc-val">${{fmtVal(v, axInfo.fmt)}}</div>
      <div class="bc-sub">${{sub}}</div>
    </div>`;
  }}).join('');
}}

// ══════════════════════════════════════════════════════════════════════════
// TAB 2: DISPERSÃO (Plotly)
// ══════════════════════════════════════════════════════════════════════════
const SCATTER_AXES = [
  {{k:'inv_def', l:'Investimento Deflac. (R$)'}},
  {{k:'pub_def', l:'Bilheteria Deflac. (R$)'}},
  {{k:'rec_def', l:'Receita Deflac. (R$)'}},
  {{k:'roi_def', l:'ROI Doméstico Deflac.'}},
  {{k:'rim',     l:'ROI Intl Máximo'}},
  {{k:'ria',     l:'ROI Intl Médio'}},
  {{k:'n',       l:'N Obras'}},
  {{k:'np_uniq', l:'Países Únicos'}},
];
let scatterScale = 'log';
let scatterInited = false;

function buildScatterControls() {{
  const sx = document.getElementById('sx');
  const sy = document.getElementById('sy');
  SCATTER_AXES.forEach(a => {{
    sx.innerHTML += `<option value="${{a.k}}">${{a.l}}</option>`;
    sy.innerHTML += `<option value="${{a.k}}">${{a.l}}</option>`;
  }});
  sx.value = 'inv_def';
  sy.value = 'roi_def';
}}

function setScalePlotly(mode) {{
  scatterScale = mode;
  document.getElementById('s-lin').classList.toggle('on', mode==='lin');
  document.getElementById('s-log').classList.toggle('on', mode==='log');
  renderScatterPlotly();
}}

function renderScatterPlotly() {{
  if(!document.getElementById('scatter-plotly')) return;
  const xk = document.getElementById('sx').value;
  const yk = document.getElementById('sy').value;
  const xLabel = SCATTER_AXES.find(a=>a.k===xk).l;
  const yLabel = SCATTER_AXES.find(a=>a.k===yk).l;
  const traces = [];
  CL_ORDER.forEach(cl => {{
    const info = CL[cl];
    const pts = PROD.filter(r => r.cl===cl && getVal(r,xk)>0 && getVal(r,yk)>0);
    if(!pts.length) return;
    traces.push({{
      x: pts.map(r => getVal(r,xk)),
      y: pts.map(r => getVal(r,yk)),
      text: pts.map(r => r.nm),
      customdata: pts.map(r => [r.n, fmtMoney(r.inv_def), fmtRatio(r.rda)]),
      mode: 'markers',
      name: info.l,
      marker: {{
        color: info.c,
        size: pts.map(r => Math.max(6, Math.min(28, 4+r.n*2))),
        opacity: 0.7,
        line: {{width:0.5, color:'rgba(255,255,255,0.4)'}}
      }},
      hovertemplate: '<b>%{{text}}</b><br>'+xLabel+': %{{x:,.0f}}<br>'+yLabel+': %{{y:.3f}}<br>Obras: %{{customdata[0]}}<extra>%{{fullData.name}}</extra>'
    }});
  }});
  const xType = (scatterScale==='log' && ['inv_def','pub_def','rec_def'].includes(xk)) ? 'log' : 'linear';
  const yType = (scatterScale==='log' && ['inv_def','pub_def','rec_def'].includes(yk)) ? 'log' : 'linear';
  const layout = {{
    title: xLabel + ' × ' + yLabel,
    xaxis: {{title:xLabel, type:xType, gridcolor:'#272727', linecolor:'#363636'}},
    yaxis: {{title:yLabel, type:yType, gridcolor:'#272727', linecolor:'#363636'}},
    paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(20,20,20,0.5)',
    font:{{color:'#e2e8f0', family:'Inter, system-ui, sans-serif', size:11}},
    legend:{{orientation:'h', y:-0.15, font:{{size:10}}}},
    height:560, margin:{{l:70,r:20,t:50,b:70}},
    hovermode:'closest'
  }};
  Plotly.react('scatter-plotly', traces, layout, {{responsive:true, displaylogo:false}});
  scatterInited = true;
}}

// ══════════════════════════════════════════════════════════════════════════
// TAB 3: RANKING
// ══════════════════════════════════════════════════════════════════════════
let rankFilter = new Set(CL_ORDER);
let expandedIdx = null;
let rankSortKey = 'rda', rankSortAsc = false;

const RANK_SORT_OPTIONS = [
  {{k:'roi_def', l:'ROI Dom. Deflac.'}},
  {{k:'rda',     l:'ROI Dom. Proporcional'}},
  {{k:'rdm',     l:'ROI Dom. Mediano'}},
  {{k:'rim',     l:'ROI Intl Máx'}},
  {{k:'ria',     l:'ROI Intl Médio'}},
  {{k:'inv_def', l:'Investimento Deflac.'}},
  {{k:'inv',     l:'Investimento Nominal'}},
  {{k:'n',       l:'N Obras'}},
  {{k:'pub_def', l:'Bilheteria Deflac.'}},
  {{k:'pub',     l:'Bilheteria Nominal'}},
  {{k:'rec_def', l:'Receita Deflac.'}},
  {{k:'rec',     l:'Receita Nominal'}},
  {{k:'np_uniq', l:'Países Únicos'}},
  {{k:'nm',      l:'Nome (A→Z)'}},
];

const RANK_COLS = [
  {{k:'#',       l:'#',              r:true,  fmt:null,    w:'30px'}},
  {{k:'nm',      l:'Nome',           r:false, fmt:null,    w:'auto'}},
  {{k:'n',       l:'Obras',          r:true,  fmt:'int',   w:'55px'}},
  {{k:'inv_def', l:'Invest. Deflac.',r:true,  fmt:'money', w:'100px'}},
  {{k:'roi_def', l:'ROI Dom. Def.',  r:true,  fmt:'ratio', w:'90px'}},
  {{k:'rim',     l:'ROI Intl Máx',   r:true,  fmt:'ratio', w:'90px'}},
  {{k:'ria',     l:'ROI Intl Méd',   r:true,  fmt:'ratio', w:'90px'}},
  {{k:'pub_def', l:'Bilh. Deflac.',  r:true,  fmt:'money', w:'90px'}},
  {{k:'rec_def', l:'Renda Total Def.',r:true, fmt:'money', w:'100px'}},
  {{k:'np_uniq', l:'Países Únicos',  r:true,  fmt:'int',   w:'90px'}},
  {{k:'cl',      l:'Cluster',        r:false, fmt:null,    w:'120px'}},
];

function buildRankControls() {{
  const sel = document.getElementById('rank-sort');
  RANK_SORT_OPTIONS.forEach(o => {{
    sel.innerHTML += `<option value="${{o.k}}">${{o.l}}</option>`;
  }});
  sel.value = rankSortKey;

  const chips = document.getElementById('rank-chips');
  const clCounts = {{}};
  PROD.forEach(r => {{ clCounts[r.cl]=(clCounts[r.cl]||0)+1; }});
  chips.innerHTML = CL_ORDER.map(k => {{
    const info=CL[k], cnt=clCounts[k]||0;
    return `<div class="chip" id="rk-chip-${{k}}" onclick="toggleRFilter('${{k}}')">
      <div class="cdot" style="background:${{info.c}}"></div>
      ${{info.l}} <span class="chip-n">${{cnt}}</span>
    </div>`;
  }}).join('');
}}

function toggleRFilter(k) {{
  if(rankFilter.has(k)) rankFilter.delete(k); else rankFilter.add(k);
  document.getElementById('rk-chip-'+k).classList.toggle('off', !rankFilter.has(k));
  expandedIdx = null;
  renderRanking();
}}

function renderRanking() {{
  rankSortKey = document.getElementById('rank-sort').value;
  const q = document.getElementById('r-search').value.trim().toLowerCase();
  document.getElementById('r-clr').classList.toggle('show', q.length>0);

  let data = PROD.filter(r => {{
    if(!rankFilter.has(r.cl)) return false;
    if(q && !r.nm.toLowerCase().includes(q)) return false;
    return true;
  }});

  data.sort((a,b) => {{
    if(rankSortKey==='nm') {{
      return rankSortAsc ? b.nm.localeCompare(a.nm) : a.nm.localeCompare(b.nm);
    }}
    const va=a[rankSortKey]||0, vb=b[rankSortKey]||0;
    return rankSortAsc ? va-vb : vb-va;
  }});

  const nDistRank = data.filter(r=>r.tipo==='captador').length;
  document.getElementById('rank-count').textContent = data.length+' empresas ('+nDistRank+' distribuidoras)';

  // Build thead
  const thead = document.getElementById('rank-thead');
  thead.innerHTML = '<tr>'+RANK_COLS.map(c => {{
    if(c.k==='#') return `<th style="width:${{c.w}}" class="r">#</th>`;
    const sortable = c.k!=='cl';
    const cls = (c.r?' r':'')+( sortable && c.k===rankSortKey ? (rankSortAsc?' sa':' sd') : '');
    const handler = sortable ? `onclick="rankSort('${{c.k}}')"` : '';
    return `<th style="width:${{c.w}}" class="${{cls.trim()}}" ${{handler}}>${{c.l}}</th>`;
  }}).join('')+'</tr>';

  // Compute max rda for mini bars
  const maxRda = Math.max(...data.map(r=>r.rda), 1);

  // Build tbody
  const tbody = document.getElementById('rank-tbody');
  let html = '';
  data.forEach((r, i) => {{
    const isExp = expandedIdx === i;
    const clr = (CL[r.cl]||{{c:'#666'}}).c;
    const pct = maxRda>0 ? (r.rda/maxRda*80).toFixed(1) : 0;
    html += `<tr class="${{isExp?'expanded':''}}" onclick="toggleExpand(${{i}},this)" style="cursor:pointer">`;
    RANK_COLS.forEach(c => {{
      if(c.k==='#') {{
        html += `<td class="r dim">${{i+1}}</td>`;
      }} else if(c.k==='nm') {{
        const _distBadge = r.tipo==='captador' ? `<span style="font-size:8px;padding:2px 6px;border-radius:2px;background:#2a1a0a;color:#e09a4f;border:1px solid #5a3a10;margin-left:7px;vertical-align:middle;letter-spacing:.06em;font-weight:600">DISTRIBUIDORA</span>` : '';
        html += `<td>${{r.nm}}${{_distBadge}}</td>`;
      }} else if(c.k==='rda') {{
        html += `<td class="r">
          <span class="mini-bar" style="width:${{pct}}px;background:${{clr}}"></span>
          ${{fmtRatio(r.rda)}}
        </td>`;
      }} else if(c.k==='cl') {{
        html += `<td><span class="badge" style="background:${{clr}}">${{(CL[r.cl]||{{l:r.cl}}).l}}</span></td>`;
      }} else {{
        const v = r[c.k]||0;
        html += `<td class="${{c.r?'r':''}}">${{c.fmt ? fmtVal(v,c.fmt) : (v||'—')}}</td>`;
      }}
    }});
    html += '</tr>';
    if(isExp) {{
      html += `<tr class="detail-row"><td colspan="${{RANK_COLS.length}}">
        <div class="detail-row-inner">
          <div style="font-size:9px;color:var(--muted);letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px">
            Obras — ${{r.nm}}
          </div>
          ${{buildObrasTableHTML(r.obras)}}
        </div>
      </td></tr>`;
    }}
  }});
  tbody.innerHTML = html;
}}

function buildObrasTableHTML(obras) {{
  const cols = DET_COLS;
  let h = '<table><thead><tr>';
  cols.forEach(c => {{ h += `<th class="${{c.r?'r':''}}">${{c.l}}</th>`; }});
  h += '</tr></thead><tbody>';
  obras.forEach(o => {{
    h += '<tr>';
    cols.forEach(c => {{
      const v = o[c.k];
      const disp = c.fmt ? fmtVal(v,c.fmt) : (v||'—');
      h += `<td class="${{c.r?'r':''}}">${{disp}}</td>`;
    }});
    h += '</tr>';
  }});
  h += '</tbody></table>';
  return h;
}}

function toggleExpand(i, tr) {{
  expandedIdx = expandedIdx===i ? null : i;
  renderRanking();
}}

function rankSort(k) {{
  if(rankSortKey===k) rankSortAsc=!rankSortAsc;
  else {{ rankSortKey=k; rankSortAsc=false; }}
  document.getElementById('rank-sort').value = k;
  expandedIdx=null;
  renderRanking();
}}

// ══════════════════════════════════════════════════════════════════════════
// TAB 4: FILMES — film-level analysis by cluster + category
// ══════════════════════════════════════════════════════════════════════════
// Build flat list of all obras from PROD, enriched with produtora + cluster
let ALL_FILMS = [];
PROD.forEach(p => {{
  (p.obras || []).forEach(o => {{
    ALL_FILMS.push(Object.assign({{}}, o, {{
      nm: p.nm,
      cl: p.cl,
      tipo: p.tipo,
    }}));
  }});
}});

const FILM_SORT_OPTIONS = [
  {{k:'rd',  l:'ROI Dom.'}},
  {{k:'ri',  l:'ROI Intl.'}},
  {{k:'cr',  l:'Crítica (1–5)'}},
  {{k:'bd',  l:'Bilh. Deflac.'}},
  {{k:'p',   l:'Bilh. Nominal'}},
  {{k:'inv', l:'Investimento'}},
  {{k:'a',   l:'Ano'}},
  {{k:'td',  l:'Título (A→Z)'}},
];
const FILM_COLS = [
  {{k:'#',   l:'#',             r:true,  fmt:null,    w:'30px'}},
  {{k:'td',  l:'Título',        r:false, fmt:null,    w:'auto'}},
  {{k:'a',   l:'Ano',           r:true,  fmt:'int',   w:'50px'}},
  {{k:'nm',  l:'Produtora',     r:false, fmt:null,    w:'160px'}},
  {{k:'cl',  l:'Cluster',       r:false, fmt:null,    w:'120px'}},
  {{k:'cat', l:'Categoria',     r:false, fmt:null,    w:'200px'}},
  {{k:'rd',  l:'ROI Dom.',      r:true,  fmt:'ratio', w:'75px'}},
  {{k:'ri',  l:'ROI Intl.',     r:true,  fmt:'ratio', w:'75px'}},
  {{k:'bd',  l:'Bilh. Deflac.', r:true,  fmt:'money', w:'90px'}},
  {{k:'inv', l:'Investimento',  r:true,  fmt:'money', w:'90px'}},
  {{k:'cr',  l:'Crítica',       r:true,  fmt:'critica', w:'70px'}},
  {{k:'crn', l:'N Fontes',      r:true,  fmt:'int',   w:'60px'}},
];

let filmSortKey = 'rd', filmSortAsc = false;
let filmClFilter = new Set(CL_ORDER);

function buildFilmControls() {{
  const sel = document.getElementById('film-sort');
  FILM_SORT_OPTIONS.forEach(o => {{
    sel.innerHTML += `<option value="${{o.k}}">${{o.l}}</option>`;
  }});
  sel.value = filmSortKey;

  const chips = document.getElementById('film-cl-chips');
  const clCounts = {{}};
  ALL_FILMS.forEach(f => {{ clCounts[f.cl] = (clCounts[f.cl]||0)+1; }});
  chips.innerHTML = CL_ORDER.map(k => {{
    const info = CL[k], cnt = clCounts[k]||0;
    return `<div class="chip" id="fm-chip-${{k}}" onclick="toggleFmFilter('${{k}}')">
      <div class="cdot" style="background:${{info.c}}"></div>
      ${{info.l}} <span class="chip-n">${{cnt}}</span>
    </div>`;
  }}).join('');
}}

function toggleFmFilter(k) {{
  if(filmClFilter.has(k)) filmClFilter.delete(k); else filmClFilter.add(k);
  document.getElementById('fm-chip-'+k).classList.toggle('off', !filmClFilter.has(k));
  renderFilms();
}}

function fmtCritica(v) {{ return v>0 ? v.toFixed(1) : '—'; }}

function filmSort(k) {{
  if(filmSortKey===k) filmSortAsc=!filmSortAsc; else {{ filmSortKey=k; filmSortAsc=false; }}
  document.getElementById('film-sort').value = k;
  renderFilms();
}}

function renderFilms() {{
  filmSortKey = document.getElementById('film-sort').value;
  const crnMin = parseInt(document.getElementById('film-crn').value) || 0;
  const q = document.getElementById('film-search').value.trim().toLowerCase();
  document.getElementById('film-clr').classList.toggle('show', q.length>0);

  let data = ALL_FILMS.filter(f => {{
    if(!filmClFilter.has(f.cl)) return false;
    if(crnMin > 0 && (f.crn||0) < crnMin) return false;
    if(q && !f.td.toLowerCase().includes(q) && !f.nm.toLowerCase().includes(q)) return false;
    return true;
  }});

  data.sort((a,b) => {{
    if(filmSortKey==='td') return filmSortAsc ? b.td.localeCompare(a.td) : a.td.localeCompare(b.td);
    const va=a[filmSortKey]||0, vb=b[filmSortKey]||0;
    return filmSortAsc ? va-vb : vb-va;
  }});

  document.getElementById('film-count').textContent = data.length+' filmes';

  const thead = document.getElementById('film-thead');
  thead.innerHTML = '<tr>'+FILM_COLS.map(c => {{
    if(c.k==='#') return `<th style="width:${{c.w}}" class="r">#</th>`;
    const sortable = !['cl','cat','nm','td'].includes(c.k) || c.k==='td';
    const cls = (c.r?' r':'') + (c.k===filmSortKey ? (filmSortAsc?' sa':' sd') : '');
    const handler = sortable ? `onclick="filmSort('${{c.k}}')"` : '';
    return `<th style="width:${{c.w}}" class="${{cls.trim()}}" ${{handler}}>${{c.l}}</th>`;
  }}).join('')+'</tr>';

  const tbody = document.getElementById('film-tbody');
  tbody.innerHTML = data.map((f,i) => '<tr>'+FILM_COLS.map(c => {{
    if(c.k==='#') return `<td class="r dim">${{i+1}}</td>`;
    if(c.k==='cl') {{
      const clr = (CL[f.cl]||{{c:'#666'}}).c;
      return `<td><span class="badge" style="background:${{clr}}">${{(CL[f.cl]||{{l:f.cl}}).l}}</span></td>`;
    }}
    if(c.k==='cr') return `<td class="r">${{fmtCritica(f.cr||0)}}</td>`;
    if(c.k==='td' || c.k==='nm' || c.k==='cat') {{
      const v = f[c.k]||'—';
      const short = v.length>40 ? v.slice(0,38)+'…' : v;
      return `<td title="${{v}}">${{short}}</td>`;
    }}
    const v = f[c.k]||0;
    return `<td class="${{c.r?'r':''}}">${{c.fmt && c.fmt!=='critica' ? fmtVal(v,c.fmt) : (v||'—')}}</td>`;
  }}).join('')+'</tr>').join('');
}}

// ══════════════════════════════════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════════════════════════════════
buildTab1();
buildScatterControls();
buildRankControls();
renderRanking();
buildFilmControls();
renderFilms();
</script>
</body>
</html>"""

out_path = BASE / "resultados" / "painel_produtoras.html"
out_path.write_text(HTML, encoding="utf-8")
print(f"\npainel_produtoras.html gerado: {len(records)} produtoras - {year_min} a {year_max}")
print(f"Tamanho do arquivo: {out_path.stat().st_size / 1024:.1f} KB")
