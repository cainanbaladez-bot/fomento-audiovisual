# -*- coding: utf-8 -*-
import re, unicodedata, pathlib, csv as _csv
import numpy as np
import pandas as pd

BASE = pathlib.Path(__file__).parent.parent

def norm_title(s):
    if not isinstance(s, str): return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().upper()

def norm_cnpj(s):
    if not isinstance(s, str): s = str(s) if pd.notna(s) else ""
    digits = re.sub(r"\D", "", s)
    return digits.lstrip("0").zfill(14)

def _fix_enc(s):
    try: return s.encode('latin-1').decode('utf-8')
    except: return s

# 1. Carrega tabela consolidada
df = pd.read_excel(BASE / "resultados" / "tabela_consolidada_obras.xlsx", sheet_name=0)
df.columns = df.columns.str.strip()
print(f"Obras carregadas: {len(df)}")

COL_MAP = {
    "Projeto": "titulo", "Ano": "ano", "Chamada": "chamada", "Categoria": "categoria",
    "Valor FSA (R$)": "vfsa", "Renúncia Art.3/3-A/39 (R$)": "vren3", "Renúncia Outros Mec. (R$)": "vrenO",
    "Bilheteria Nominal (R$)": "bilheteria", "Bilheteria Deflac. (R$)": "bilh_def",
    "Estimativa Outras Janelas (R$)": "janelas", "Outras Janelas Deflac. (R$2024)": "janelas_def",
    "Investimento Total Deflac. (R$2024)": "inv_def_col", "ROI Internacional (0-100)": "roi_intl",
    "VOD Intl \u2014 N Pa\u00edses": "n_paises",
}
df = df.rename(columns={k: v for k, v in COL_MAP.items() if k in df.columns})
for col in ["vfsa","vren3","vrenO","bilheteria","bilh_def","janelas","janelas_def","inv_def_col","roi_intl","n_paises"]:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
df["ano"] = pd.to_numeric(df.get("ano", 0), errors="coerce").fillna(0).astype(int)
df["titulo_norm"] = df["titulo"].apply(norm_title)
df["inv_linha"] = df["vfsa"] + df["vren3"] + df["vrenO"]

# 2. CPB lookup
def _build_cpb_lookup(base):
    fsa = pd.read_csv(base/"raw"/"obras-nao-pub-brasileiras-investimento-fsa.csv", sep=None, engine="python", encoding="utf-8-sig", dtype=str, usecols=["CPB","TITULO_ORIGINAL"], on_bad_lines="skip").fillna("").drop_duplicates("CPB")
    ind = pd.read_csv(base/"raw"/"obras-nao-pub-brasileiras-fomento-indireto.csv", sep=None, engine="python", encoding="utf-8-sig", dtype=str, usecols=["CPB","TITULO_ORIGINAL"], on_bad_lines="skip").fillna("").drop_duplicates("CPB")
    return pd.concat([fsa, ind]).drop_duplicates("CPB", keep="first")

obras_cpb = _build_cpb_lookup(BASE)
obras_cpb["titulo_norm"] = obras_cpb["TITULO_ORIGINAL"].apply(norm_title)
obras_cpb = obras_cpb.drop_duplicates("titulo_norm")
if "CPB" in df.columns: df = df.drop(columns=["CPB"])
df = df.merge(obras_cpb[["titulo_norm","CPB","TITULO_ORIGINAL"]], on="titulo_norm", how="left")

# 3. Produtores
prod = pd.read_csv(BASE/"raw"/"produtores-de-obras-nao-publicitarias-brasileiras.csv", sep=";", encoding="utf-8", usecols=["CPB","CNPJ_PRODUTOR"], on_bad_lines="skip").dropna(subset=["CPB","CNPJ_PRODUTOR"])
prod["cnpj_norm"] = prod["CNPJ_PRODUTOR"].apply(norm_cnpj)

# 4. Agentes independentes
agentes = pd.read_csv(BASE/"raw"/"agentes-economicos-regulares.csv", sep=";", encoding="utf-8", usecols=["RAZAO_SOCIAL","CNPJ","BRASILEIRO_INDEPENDENTE"], on_bad_lines="skip")
agentes["cnpj_norm"] = agentes["CNPJ"].apply(norm_cnpj)
ind_ag = agentes[agentes["BRASILEIRO_INDEPENDENTE"].str.strip().str.upper()=="SIM"].drop_duplicates("cnpj_norm")
ind_cnpjs = set(ind_ag["cnpj_norm"])

cpb_to_cnpj_ind = prod[prod["cnpj_norm"].isin(ind_cnpjs)][["CPB","cnpj_norm"]].drop_duplicates()
cpb_map = df[["titulo_norm","CPB"]].dropna(subset=["CPB"]).drop_duplicates("titulo_norm")
obra_prod_map = cpb_to_cnpj_ind.merge(cpb_map, on="CPB", how="inner")[["titulo_norm","cnpj_norm"]].drop_duplicates()
obra_prod_map = obra_prod_map.merge(ind_ag[["cnpj_norm","RAZAO_SOCIAL"]], on="cnpj_norm", how="left")
df_ind = df.merge(obra_prod_map, on="titulo_norm", how="inner")
print(f"Linhas com produtora independente: {len(df_ind)} | Produtoras: {df_ind['cnpj_norm'].nunique()}")

# 5a. Agrega por produtora
def _fsa_ren(g):
    inv_tot_nom = (g["vfsa"]+g["vren3"]+g["vrenO"]).sum()
    fsa_nom = g["vfsa"].sum()
    frac_fsa = fsa_nom/inv_tot_nom if inv_tot_nom>0 else 0.0
    inv_d = g["inv_def_col"].sum()
    return inv_d*frac_fsa, inv_d*(1-frac_fsa)

raw_records = []
for cnpj, grp in df_ind.groupby("cnpj_norm"):
    nm = grp["RAZAO_SOCIAL"].iloc[0]
    obra_dedup = grp.groupby("titulo_norm").agg(bilh_def=("bilh_def","max"), janelas_def=("janelas_def","max"), roi_intl=("roi_intl","max")).reset_index()
    n = len(obra_dedup)
    inv_total = float(grp["inv_def_col"].sum())
    rec_total = float(obra_dedup["bilh_def"].sum()) + float(obra_dedup["janelas_def"].sum())
    inv_fsa_d, inv_ren_d = _fsa_ren(grp)
    rda = rec_total/inv_total if inv_total>=1000 else 0.0
    rim = float(obra_dedup["roi_intl"].max())
    raw_records.append({"nm":nm,"n":n,"rec":rec_total,"inv":inv_total,"inv_fsa":inv_fsa_d,"inv_ren":inv_ren_d,"bil":float(obra_dedup["bilh_def"].sum()),"rim":rim,"rda":rda})

print(f"  Produtoras raw: {len(raw_records)}")

# 5b. Captadores/distribuidoras FSA (mesmo bloco do script 04)
cap_map_fsa = {}
try:
    with open(BASE/"raw"/"projetos-fsa.csv", encoding='latin-1') as _f:
        for _row in _csv.DictReader(_f, delimiter=';'):
            _tit = norm_title(_fix_enc(_row.get('TITULO_PROJETO','') or ''))
            _prop = _fix_enc(_row.get('RAZAO_SOCIAL_PROPONENTE','') or '').strip()
            _prod_r = _fix_enc(_row.get('RAZAO_SOCIAL_PRODUTORA','') or '').strip()
            if not _tit: continue
            _captador = _prop if _prop else _prod_r
            _tipo = 'produtora' if (not _prop) or _prop.upper()==_prod_r.upper() else 'captador'
            cap_map_fsa[_tit] = {'captador': _captador, 'tipo': _tipo}
    print(f"  cap_map_fsa: {len(cap_map_fsa)} titulos")
except Exception as e:
    print(f"  AVISO cap_map_fsa: {e}")

if cap_map_fsa:
    df["_captador"] = df["titulo_norm"].map(lambda t: cap_map_fsa.get(t,{}).get("captador",""))
    df["_captador_tipo"] = df["titulo_norm"].map(lambda t: cap_map_fsa.get(t,{}).get("tipo","produtora"))
    _cat_col = "categoria" if "categoria" in df.columns else None
    if _cat_col:
        df_cap = df[(df["vfsa"]>0) & (df["_captador"]!="") & (~df[_cat_col].str.upper().str.contains("TV_EXCLUIR",na=False))].copy()
    else:
        df_cap = df[(df["vfsa"]>0) & (df["_captador"]!="")].copy()

    cap_raw = []
    for cap_nm, grp in df_cap.groupby("_captador"):
        if not cap_nm: continue
        if grp["_captador_tipo"].iloc[0] != "captador": continue
        obra_dedup2 = grp.groupby("titulo_norm").agg(bilh_def=("bilh_def","max"), janelas_def=("janelas_def","max"), roi_intl=("roi_intl","max")).reset_index()
        n2 = len(obra_dedup2)
        inv2 = float(grp["inv_def_col"].sum())
        rec2 = float(obra_dedup2["bilh_def"].sum()) + float(obra_dedup2["janelas_def"].sum())
        inv_fsa2, inv_ren2 = _fsa_ren(grp)
        rda2 = rec2/inv2 if inv2>=1000 else 0.0
        rim2 = float(obra_dedup2["roi_intl"].max())
        cap_raw.append({"nm":cap_nm,"n":n2,"rec":rec2,"inv":inv2,"inv_fsa":inv_fsa2,"inv_ren":inv_ren2,"bil":float(obra_dedup2["bilh_def"].sum()),"rim":rim2,"rda":rda2})

    cap_names_up = {r["nm"].upper() for r in cap_raw}
    raw_records = [r for r in raw_records if r["nm"].upper() not in cap_names_up]
    raw_records.extend(cap_raw)
    print(f"  Captadores adicionados: {len(cap_raw)}")

print(f"Total final: {len(raw_records)} entidades")

# 6. Clusters
BILH_TH    = 2_500_000
INV_TH     = 5_000_000
REC_NEW_TH = 10_000_000
ROI_DOM_TH = 0.6
REC_PISO   = 2_500_000

def cluster_atual(r):
    rec, rim, inv = r["rec"], r["rim"], r["inv"]
    if rec >= BILH_TH and rim >= 13: return "Duplo Retorno"
    if rec > BILH_TH:                return "Retorno Domestico"
    if rim >= 13:                    return "Retorno Internacional"
    if inv > INV_TH:                 return "Fomento Baixo Retorno"
    return "Pequeno Porte"

def cluster_novo(r):
    rec, rim, inv, rda = r["rec"], r["rim"], r["inv"], r["rda"]
    if rec >= BILH_TH and rim >= 13: return "Duplo Retorno"
    is_dom = (rec >= REC_NEW_TH) or (rda > ROI_DOM_TH and rec >= REC_PISO)
    if is_dom and rim < 13:          return "Retorno Domestico"
    if rim >= 13:                    return "Retorno Internacional"
    if inv > INV_TH:                 return "Fomento Baixo Retorno"
    return "Pequeno Porte"

for r in raw_records:
    r["cl_atual"] = cluster_atual(r)
    r["cl_novo"]  = cluster_novo(r)

df_r = pd.DataFrame(raw_records)

print("\n=== Distribuicao: Atual vs Novo ===")
comp = pd.DataFrame({"Atual": df_r["cl_atual"].value_counts(), "Novo": df_r["cl_novo"].value_counts()}).fillna(0).astype(int)
print(comp)

def agg(df_r, col):
    return df_r.groupby(col).agg(
        produtoras=("nm","count"), obras=("n","sum"),
        inv_fsa=("inv_fsa","sum"), inv_ren=("inv_ren","sum"),
        inv_total=("inv","sum"), bilheteria=("bil","sum"), renda_total=("rec","sum"),
    ).reset_index()

def fmt(v): return f"R$ {v/1e6:,.1f}M"
ORDEM = ["Duplo Retorno","Retorno Domestico","Retorno Internacional","Fomento Baixo Retorno","Pequeno Porte"]

print("\n=== TABELA ATUAL ===")
ag = agg(df_r,"cl_atual").set_index("cl_atual")
for cl in ORDEM:
    if cl not in ag.index: continue
    r = ag.loc[cl]
    print(f"{cl}: {int(r['produtoras'])} prod | {int(r['obras'])} obras | FSA {fmt(r['inv_fsa'])} | Ren {fmt(r['inv_ren'])} | Tot {fmt(r['inv_total'])} | Bil {fmt(r['bilheteria'])} | Renda {fmt(r['renda_total'])}")

print("\n=== TABELA NOVA (rec>=10M ou ROI Dom>0.6 e rec>=2.5M) ===")
ag2 = agg(df_r,"cl_novo").set_index("cl_novo")
for cl in ORDEM:
    if cl not in ag2.index: continue
    r = ag2.loc[cl]
    print(f"{cl}: {int(r['produtoras'])} prod | {int(r['obras'])} obras | FSA {fmt(r['inv_fsa'])} | Ren {fmt(r['inv_ren'])} | Tot {fmt(r['inv_total'])} | Bil {fmt(r['bilheteria'])} | Renda {fmt(r['renda_total'])}")

mudou = df_r[df_r["cl_atual"]!=df_r["cl_novo"]][["nm","cl_atual","cl_novo","rec","rda","rim","inv"]].copy()
mudou["rec"] = (mudou["rec"]/1e6).round(2)
mudou["inv"] = (mudou["inv"]/1e6).round(2)
mudou["rda"] = mudou["rda"].round(3)
mudou.columns = ["Produtora","Atual","Novo","Renda(M)","ROI Dom","ROI Intl","Inv(M)"]
print(f"\n=== MUDANCAS ({len(mudou)} produtoras) ===")
print(mudou.sort_values(["Novo","Atual"]).to_string(index=False))
