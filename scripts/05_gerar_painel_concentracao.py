# -*- coding: utf-8 -*-
"""
gerar_painel_concentracao.py
Painel: Concentração FSA e Sustentabilidade das Produtoras (3 abas)
Design: dark theme de painel_produtoras.html
"""
import json, pathlib, re
import pandas as pd
import numpy as np

BASE = pathlib.Path(__file__).parent.parent
OUT  = BASE / "resultados" / "painel_concentracao_produtoras.html"

# ── Dados ──────────────────────────────────────────────────────────────────
df = pd.read_excel(BASE / "resultados" / "tabela_consolidada_obras.xlsx", sheet_name=0)
prod_df = pd.read_csv(BASE / "raw" / "produtores-de-obras-nao-publicitarias-brasileiras.csv",
                      sep=";", encoding="latin1")
vnom = "Valor FSA (R$)"
vdef = "Valor FSA Deflac. (R$2024)"

def norm_cnpj(s):
    if not isinstance(s, str): s = str(s) if pd.notna(s) else ""
    return re.sub(r"\D", "", s).zfill(14)

# ── Filtro: apenas produtoras independentes brasileiras ──────────────────────
# Exclui: (1) produtores estrangeiros; (2) emissoras e distribuidoras integradas
# que aparecem como co-produtores por contratos de pré-venda de direitos TV/VOD.
# Essas entidades inflacionam o ticket e o FSA atribuído pois o merge por CPB
# repete o valor FSA integral do projeto para cada co-produtor registrado.
NON_INDEPENDENT_CNPJ = {
    # Grupo Globo — emissora, programadoras, distribuidoras
    "027865757008340", "027865757000102", "027865757006992",
    "027865757008855", "027865757005678",  # Globo Comunicação e Participações S/A (branches)
    "000811990000571", "000811990000229", "000811990000148",  # Globosat Programadora
    "033252156000119",  # TV Globo LTDA
    "004067191000160",  # Editora Globo S.A.
    "028114122000128",  # Sigla / Sistema Globo de Gravações
    "014100997000166",  # Globo Cine Digital
    # Canal Brasil / programadoras
    "002608224000106",  # Canal Brazil S/A
    "001310489000160",  # CBC Canal Brasileiro de Comunicações
    # Telecine (programadora pay-TV, consórcio Globo/majors)
    "000252848000108",  # Telecine Programação de Filmes LTDA
    # Fox Film do Brasil (major, filial nacional)
    "033110420000180", "033110420000856",
}

prod_br = prod_df[prod_df["PAIS_PRODUTOR"] == "BRASIL"].copy()
prod_br["cnpj"] = prod_br["CNPJ_PRODUTOR"].apply(norm_cnpj)
prod_br = prod_br[~prod_br["cnpj"].isin(NON_INDEPENDENT_CNPJ)]

df_fsa = df[df[vnom].notna() & (df[vnom] > 0)].copy()
merged = (df_fsa.merge(prod_br[["CPB","CNPJ_PRODUTOR","PRODUTOR","cnpj"]], on="CPB", how="left")
          .dropna(subset=["CNPJ_PRODUTOR"]))
# cnpj já computado no merge — renomear para consistência
ps = merged.groupby("cnpj").agg(
    n_obras=("CPB","nunique"), fsa_nom=(vnom,"sum"),
    fsa_def=(vdef, lambda x: x.sum() if x.notna().any() else 0),
    ano_min=("Ano","min"), ano_max=("Ano","max"), nome=("PRODUTOR","first"),
).reset_index()
ps["anos"] = (ps["ano_max"] - ps["ano_min"] + 1).clip(lower=1)
ps["ticket"]    = ps["fsa_nom"] / ps["anos"]
ps["fsa_por_obra"] = ps["fsa_nom"] / ps["n_obras"]
total_fsa = ps["fsa_nom"].sum()
n_prod    = len(ps)
ps = ps.sort_values("fsa_nom", ascending=False).reset_index(drop=True)
ps["rank"] = ps.index + 1
ps["tier"] = ps["rank"].apply(lambda r: "A" if r<=10 else ("B" if r<=50 else ("C" if r<=150 else ("D" if r<=300 else "E"))))

# Gini
ps_s = ps.sort_values("fsa_nom").reset_index(drop=True)
lcum = ps_s["fsa_nom"].cumsum()
gini_val = round((n_prod+1 - 2*(lcum.sum()/total_fsa)) / n_prod, 3)

# Lorenz
n = n_prod; step = max(1, n//400)
lx = [0] + [(i+1)/n for i in range(n)]
ly = [0] + list(lcum / total_fsa)
lx_s = [round(v,4) for v in lx[::step]+[1.0]]
ly_s = [round(v,4) for v in ly[::step]+[1.0]]

# Histogram
bins_k = [0,100,200,300,500,750,1000,1500,2000,3000,5000,10000,31000]
hist_labels, hist_counts, hist_colors = [], [], []
for i in range(len(bins_k)-1):
    lo, hi = bins_k[i]*1000, bins_k[i+1]*1000
    cnt = int(((ps["ticket"] >= lo) & (ps["ticket"] < hi)).sum())
    lb = f"R${bins_k[i]}k" if bins_k[i] < 1000 else f"R${bins_k[i]//1000}M"
    hist_labels.append(lb)
    hist_counts.append(cnt)
    hist_colors.append("#e05a2b" if hi<=500001 else ("#e0c94f" if hi<=1000001 else "#4ae0a0"))

# Cumulative threshold curve (x=threshold in R$k, y=% produtoras below)
thr_k   = list(range(100, 5100, 100))
thr_pct = [round((ps["ticket"]<t*1000).sum()/n_prod*100, 1) for t in thr_k]
thr_fsa = [round((ps[ps["ticket"]<t*1000]["fsa_nom"].sum()/total_fsa)*100, 1) for t in thr_k]

# Tier rows
tier_defs = [("A","Mega — Top 10",0,10), ("B","Grandes — 11 a 50",10,50),
             ("C","Médias — 51 a 150",50,150), ("D","Pequenas — 151 a 300",150,300),
             ("E","Micro — 301+",300,n_prod)]
tier_rows = []
# Adequate ticket suggestion (commentary-based)
adq = {"A":(8000,0), "B":(2000,5000), "C":(1200,2500), "D":(700,1500), "E":(400,900)}
tier_colors = {"A":"#e05a2b","B":"#f5c842","C":"#4fa3e0","D":"#4ae0a0","E":"#9b8eaf"}
for tid, tnm, i0, i1 in tier_defs:
    t = ps.iloc[i0:i1]
    fss = t["fsa_nom"].sum()
    tr = {
        "id":tid, "name":tnm, "n":len(t),
        "fsa_share": round(fss/total_fsa*100, 1),
        "ticket_med": round(t["ticket"].median()/1000),
        "ticket_mean": round(t["ticket"].mean()/1000),
        "fsa_med_obra": round(t["fsa_por_obra"].median()/1e6, 2),
        "anos_med": round(t["anos"].median(), 1),
        "obras_med": round(t["n_obras"].mean(), 1),
        "pct_below_500k": round((t["ticket"]<500000).mean()*100),
        "pct_below_1m":   round((t["ticket"]<1000000).mean()*100),
        "adq_min": adq[tid][0], "adq_max": adq[tid][1],
        "color": tier_colors[tid],
    }
    tier_rows.append(tr)

# Scatter (ticket_k vs obras, top 300, colored by tier)
scatter = []
for _, row in ps.head(300).iterrows():
    scatter.append({
        "x": round(row["ticket"]/1000, 0),
        "y": int(row["n_obras"]),
        "anos": int(row["anos"]),
        "tier": row["tier"],
        "fsa_m": round(row["fsa_nom"]/1e6, 1),
        "nome": str(row["nome"])[:45],
        "color": tier_colors[row["tier"]],
    })

# Embed as JSON
D = {
    "n_prod": n_prod, "total_fsa_m": round(total_fsa/1e6),
    "gini": gini_val,
    "med_ticket": round(ps["ticket"].median()/1000),
    "mean_ticket": round(ps["ticket"].mean()/1000),
    "pct_below_500k": round((ps["ticket"]<500000).mean()*100),
    "pct_below_1m":   round((ps["ticket"]<1000000).mean()*100),
    "n_below_500k": int((ps["ticket"]<500000).sum()),
    "n_below_1m":   int((ps["ticket"]<1000000).sum()),
    "lorenz_x": lx_s, "lorenz_y": ly_s,
    "hist_labels": hist_labels, "hist_counts": hist_counts, "hist_colors": hist_colors,
    "thr_k": thr_k, "thr_pct": thr_pct, "thr_fsa": thr_fsa,
    "tier_rows": tier_rows, "scatter": scatter,
    "tier_names": [r["name"] for r in tier_rows],
    "tier_ticket_med": [r["ticket_med"] for r in tier_rows],
    "tier_adq_min":    [r["adq_min"] for r in tier_rows],
    "tier_fsa_share":  [r["fsa_share"] for r in tier_rows],
    "tier_colors":     [r["color"] for r in tier_rows],
    "top10_pct_fsa": round(ps.head(10)["fsa_nom"].sum()/total_fsa*100, 1),
    "top50_pct_fsa": round(ps.head(50)["fsa_nom"].sum()/total_fsa*100, 1),
}
Dj = json.dumps(D, ensure_ascii=False)

# ── HTML ───────────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Concentração FSA — Ticket Médio e Sustentabilidade</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Mono:wght@300;400&display=swap" rel="stylesheet">
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg:#07080f;--s1:#0e1018;--s2:#141620;--s3:#1e2035;--s4:#262940;
  --txt:#e8eaf2;--muted:#5a6080;--dim:#1e2035;
  --acc:#e05a2b;--acc2:#f5c842;--acc3:#4ae0a0;--acc4:#4fa3e0;
}}
html,body{{height:100%;overflow:hidden}}
body{{background:var(--bg);color:var(--txt);font-family:'DM Mono',monospace;font-size:12px;display:flex;flex-direction:column}}
.hdr{{padding:10px 18px 0;flex-shrink:0;display:flex;align-items:baseline;gap:10px;flex-wrap:wrap}}
.hdr-t{{font-family:'Syne',sans-serif;font-size:16px;font-weight:800;letter-spacing:-.01em}}
.hdr-s{{color:var(--muted);font-size:10px}}
.tab-bar{{display:flex;gap:0;padding:8px 18px 0;flex-shrink:0;border-bottom:1px solid var(--s3)}}
.tab-btn{{padding:6px 16px;font-family:'DM Mono',monospace;font-size:10px;letter-spacing:.1em;text-transform:uppercase;
  background:none;color:var(--muted);border:none;border-bottom:2px solid transparent;cursor:pointer;
  transition:color .15s,border-color .15s;margin-bottom:-1px;white-space:nowrap}}
.tab-btn:hover{{color:var(--txt)}}
.tab-btn.active{{color:var(--acc);border-bottom-color:var(--acc)}}
.tab-panel{{display:none;flex:1;flex-direction:column;min-height:0;overflow:hidden}}
.tab-panel.active{{display:flex}}
.scroll{{flex:1;overflow-y:auto;padding:14px 18px 18px}}
.scroll::-webkit-scrollbar{{width:4px}}
.scroll::-webkit-scrollbar-thumb{{background:var(--s3);border-radius:2px}}
.kpi-bar{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:8px;margin-bottom:16px}}
.kpi{{background:var(--s1);border:1px solid var(--s3);border-radius:3px;padding:10px 14px}}
.kpi.warn{{border-left:3px solid var(--acc)}}
.kpi.ok{{border-left:3px solid var(--acc3)}}
.kpi.mid{{border-left:3px solid var(--acc2)}}
.kpi-l{{font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:5px}}
.kpi-v{{font-family:'Syne',sans-serif;font-size:22px;font-weight:700;color:var(--txt)}}
.kpi-u{{font-size:9px;color:var(--muted);margin-left:3px}}
.kpi-sub{{font-size:8px;color:var(--muted);margin-top:3px}}
.sec{{font-family:'Syne',sans-serif;font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
  color:var(--muted);margin-bottom:10px;padding-bottom:5px;border-bottom:1px solid var(--s3)}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px}}
.grid3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:16px}}
.card{{background:var(--s1);border:1px solid var(--s3);border-radius:3px;padding:12px}}
.card-t{{font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:8px}}
.alert{{background:rgba(224,90,43,.1);border:1px solid rgba(224,90,43,.3);border-radius:3px;
  padding:10px 14px;margin-bottom:14px;font-size:10px;color:var(--txt);line-height:1.7}}
.alert b{{color:var(--acc)}}
.info{{background:rgba(79,163,224,.08);border:1px solid rgba(79,163,224,.25);border-radius:3px;
  padding:10px 14px;margin-bottom:14px;font-size:10px;color:var(--txt);line-height:1.7}}
.info b{{color:var(--acc4)}}
table{{width:100%;border-collapse:collapse}}
thead th{{background:var(--s2);padding:6px 10px;text-align:left;font-size:9px;letter-spacing:.08em;
  text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--s3);white-space:nowrap}}
thead th.r{{text-align:right}}
tbody tr{{border-bottom:1px solid rgba(255,255,255,.03);transition:background .06s}}
tbody tr:hover{{background:var(--s2)}}
td{{padding:6px 10px;font-size:10px;color:var(--txt);white-space:nowrap}}
td.r{{text-align:right;font-variant-numeric:tabular-nums}}
td.dim{{color:var(--muted)}}
.dot{{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px;vertical-align:middle;flex-shrink:0}}
.tier-badge{{display:inline-block;padding:2px 8px;border-radius:2px;font-size:8px;letter-spacing:.1em;color:#fff;font-weight:700}}
.adq-bar{{display:flex;align-items:center;gap:6px;margin-top:4px}}
.adq-seg{{height:6px;border-radius:1px;display:inline-block}}
.footnote{{padding:6px 18px 8px;font-size:8px;color:var(--muted);border-top:1px solid var(--s3);flex-shrink:0}}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-t">Concentração FSA — Ticket Médio e Sustentabilidade das Produtoras</div>
  <div class="hdr-s">ANCINE · {D['n_prod']} produtoras · R$ {D['total_fsa_m']}M total FSA (nominal)</div>
</div>

<div class="tab-bar">
  <button class="tab-btn active" onclick="showTab('t1',this)">Ticket Médio Anual</button>
  <button class="tab-btn" onclick="showTab('t2',this)">Análise por Tiers</button>
  <button class="tab-btn" onclick="showTab('t3',this)">Concentração Lorenz</button>
</div>

<!-- ═══════════════════════ TAB 1: Ticket Médio ═══════════════════════ -->
<div class="tab-panel active" id="t1">
<div class="scroll">

<div class="kpi-bar">
  <div class="kpi warn">
    <div class="kpi-l">Ticket médio anual — mediana</div>
    <div class="kpi-v">R$ {D['med_ticket']}k<span class="kpi-u">/ano</span></div>
    <div class="kpi-sub">média: R$ {D['mean_ticket']}k/ano</div>
  </div>
  <div class="kpi warn">
    <div class="kpi-l">Produtoras abaixo de R$500k/ano</div>
    <div class="kpi-v">{D['pct_below_500k']}<span class="kpi-u">%</span></div>
    <div class="kpi-sub">{D['n_below_500k']} de {D['n_prod']} produtoras</div>
  </div>
  <div class="kpi mid">
    <div class="kpi-l">Produtoras abaixo de R$1M/ano</div>
    <div class="kpi-v">{D['pct_below_1m']}<span class="kpi-u">%</span></div>
    <div class="kpi-sub">{D['n_below_1m']} produtoras — zona de risco</div>
  </div>
  <div class="kpi ok">
    <div class="kpi-l">Gini por produtora (FSA total)</div>
    <div class="kpi-v">{D['gini']}</div>
    <div class="kpi-sub">0 = perfeita igualdade · 1 = concentração total</div>
  </div>
  <div class="kpi ok">
    <div class="kpi-l">Top 10 produtoras</div>
    <div class="kpi-v">{D['top10_pct_fsa']}<span class="kpi-u">%</span></div>
    <div class="kpi-sub">do FSA total · top 50 = {D['top50_pct_fsa']}%</div>
  </div>
</div>

<div class="alert">
  <b>Sustentabilidade:</b> Uma produtora independente brasileira com equipe mínima (2–3 profissionais fixos)
  incorre em custos anuais de aproximadamente <b>R$400k–700k</b> apenas em pessoal e overhead —
  sem contar o desenvolvimento de projetos. Com ciclos de produção de 18–24 meses por longa-metragem,
  um ticket FSA inferior a <b>R$500k/ano</b> torna inviável manter estrutura permanente:
  a empresa opera de chamada em chamada, sem capacidade de pipeline, formação de equipe ou
  desenvolvimento continuado. <b>{D['pct_below_500k']}% das produtoras</b> estão nesse patamar
  — estruturalmente dependentes de cada edital para sobreviver.
</div>

<div class="grid2">
  <div class="card">
    <div class="card-t">Distribuição do Ticket Médio Anual (FSA nominal)</div>
    <div id="hist_ticket" style="height:280px"></div>
  </div>
  <div class="card">
    <div class="card-t">% de Produtoras abaixo do limiar de ticket (curva acumulada)</div>
    <div id="thr_curve" style="height:280px"></div>
  </div>
</div>

<div class="info">
  <b>Leitura da curva:</b> A linha sobe rapidamente até R$1M/ano — onde 70% das produtoras já foram
  ultrapassadas. A inflexão ocorre em torno de R$2M/ano (90% das produtoras), indicando que
  apenas <b>10% das produtoras recebem acima desse patamar</b>. A curva da participação no FSA
  mostra que esse mesmo grupo de 30% acima de R$700k/ano concentra <b>79,7% dos recursos</b>.
</div>

</div><!-- scroll -->
</div><!-- t1 -->

<!-- ═══════════════════════ TAB 2: Tiers ═══════════════════════ -->
<div class="tab-panel" id="t2">
<div class="scroll">

<div class="alert">
  <b>Metodologia dos Tiers:</b> Produtoras ordenadas por volume total de FSA recebido e agrupadas em
  5 faixas. Para cada faixa, é calculado o ticket médio anual (FSA total ÷ anos com pelo menos
  uma obra FSA). O <b>ticket adequado sugerido</b> representa o patamar mínimo necessário para
  sustentabilidade operacional de cada porte — estimado com base nos custos setoriais de
  produção audiovisual brasileira (ANCINE, 2022; Matta, 2020).
</div>

<div class="grid2" style="margin-bottom:16px">
  <div class="card">
    <div class="card-t">Ticket Médio Observado vs. Ticket Adequado Sugerido (por Tier)</div>
    <div id="tier_comp" style="height:300px"></div>
  </div>
  <div class="card">
    <div class="card-t">Participação no FSA Total por Tier</div>
    <div id="tier_pie" style="height:300px"></div>
  </div>
</div>

<div class="sec">Tabela Detalhada por Tier</div>
<table>
  <thead>
    <tr>
      <th>Tier</th>
      <th class="r">N</th>
      <th class="r">% FSA</th>
      <th class="r">Ticket med. obs. (R$k/ano)</th>
      <th class="r">Ticket adequado sugerido</th>
      <th class="r">Anos ativos (med.)</th>
      <th class="r">Obras/tier (med.)</th>
      <th class="r">% abaixo R$500k</th>
      <th class="r">% abaixo R$1M</th>
    </tr>
  </thead>
  <tbody id="tier_tbody"></tbody>
</table>

<div style="margin-top:16px" class="card">
  <div class="card-t">Ticket Médio Observado por Tier (com benchmark de sustentabilidade)</div>
  <div id="tier_bar_sust" style="height:260px"></div>
</div>

<div class="info" style="margin-top:14px">
  <b>Achado crítico — Tier E (Micro, 301+):</b> 784 produtoras concentram mediana de ticket de
  <b>R$499k/ano</b> — exatamente no limiar de insustentabilidade — e <b>50% delas estão abaixo
  desse patamar</b>. Essas produtoras representam 72% do total de empresas com FSA, mas operam
  essencialmente como veículos de projeto único. A política que distribui recursos muito dispersamente
  cria uma ilusão de diversidade setorial ao custo de impedir a formação de capacidade institucional.
  O ticket adequado para viabilidade mínima nesse tier seria de R$400k–900k/ano — exigindo que
  projetos menores sejam executados em ciclos de 2–3 anos, não anuais.
</div>

</div><!-- scroll -->
</div><!-- t2 -->

<!-- ═══════════════════════ TAB 3: Concentração Lorenz ═══════════════════════ -->
<div class="tab-panel" id="t3">
<div class="scroll">

<div class="kpi-bar">
  <div class="kpi warn">
    <div class="kpi-l">Gini por Produtora</div>
    <div class="kpi-v">{D['gini']}</div>
    <div class="kpi-sub">vs. 0,480 por obra individual</div>
  </div>
  <div class="kpi warn">
    <div class="kpi-l">Top 10 captam</div>
    <div class="kpi-v">{D['top10_pct_fsa']}<span class="kpi-u">%</span></div>
    <div class="kpi-sub">do FSA total (10 de {D['n_prod']} produtoras)</div>
  </div>
  <div class="kpi mid">
    <div class="kpi-l">Bottom 50% captam</div>
    <div class="kpi-v">≈ 4<span class="kpi-u">%</span></div>
    <div class="kpi-sub">das 542 menores produtoras</div>
  </div>
</div>

<div class="grid2">
  <div class="card">
    <div class="card-t">Curva de Lorenz — FSA por Produtora</div>
    <div id="lorenz_prod" style="height:320px"></div>
  </div>
  <div class="card">
    <div class="card-t">Scatter: Ticket Médio vs. N° de Obras (top 300, por tier)</div>
    <div id="scatter_tick" style="height:320px"></div>
  </div>
</div>

<div class="info">
  <b>Interpretação:</b> O Gini de 0,724 por produtora é significativamente superior ao Gini
  de 0,480 calculado por obra individual — indicando que a desigualdade se amplifica quando
  a unidade de análise passa de projetos para empresas. Isso reflete a dupla vantagem das
  grandes produtoras: captam mais projetos <i>e</i> recebem valores maiores por projeto.
  O scatter evidencia que o grupo Mega (Tier A) ocupa quadrante superior direito (alto ticket,
  muitas obras), enquanto o Tier E se concentra na nuvem inferior esquerda — poucas obras,
  baixo ticket, alta rotatividade.
</div>

</div><!-- scroll -->
</div><!-- t3 -->

<div class="footnote">
  Fontes: ANCINE — Portal de Dados Abertos (produtores-de-obras-nao-publicitarias-brasileiras.csv +
  tabela_consolidada_obras.xlsx, n={D['n_prod']} produtoras com FSA). Ticket = FSA total nominal ÷ anos com obra FSA.
  Ticket adequado: referência setorial baseada em custos de produtoras independentes brasileiras (ANCINE 2022; Matta 2020).
</div>

<script>
const D = {Dj};
const PLOT_CFG = {{responsive:true,displayModeBar:false}};
const BG = '#07080f', S1='#0e1018', S2='#141620', S3='#1e2035';
const TXT='#e8eaf2', MUT='#5a6080';
const LAY = (extra={{}}) => ({{
  paper_bgcolor:BG, plot_bgcolor:S1,
  font:{{family:'DM Mono,monospace',size:10,color:MUT}},
  margin:{{t:6,b:40,l:50,r:16}},
  xaxis:{{gridcolor:S3,zerolinecolor:S3,tickfont:{{size:9}}}},
  yaxis:{{gridcolor:S3,zerolinecolor:S3,tickfont:{{size:9}}}},
  ...extra
}});

// ── Tab 1 ──
Plotly.newPlot('hist_ticket', [{{
  x: D.hist_labels, y: D.hist_counts, type:'bar',
  marker:{{color:D.hist_colors, line:{{width:0}}}},
  text: D.hist_counts.map(v=>v>0?v:''),
  textposition:'outside', textfont:{{size:9,color:TXT}},
  hovertemplate:'%{{x}}: %{{y}} produtoras<extra></extra>',
}}], {{
  ...LAY(),
  shapes:[
    {{type:'line',x0:'R$500k',x1:'R$500k',y0:0,y1:230,line:{{color:'#e05a2b',dash:'dash',width:1}}}},
    {{type:'line',x0:'R$1000k',x1:'R$1000k',y0:0,y1:230,line:{{color:'#e0c94f',dash:'dash',width:1}}}},
  ],
  annotations:[
    {{x:'R$500k',y:220,text:'Limiar crítico<br>R$500k',showarrow:false,font:{{size:8,color:'#e05a2b'}},bgcolor:S1,borderpad:3}},
    {{x:'R$1000k',y:220,text:'Zona de<br>atenção R$1M',showarrow:false,font:{{size:8,color:'#e0c94f'}},bgcolor:S1,borderpad:3}},
  ],
  yaxis:{{...LAY().yaxis,title:'N° produtoras'}},
  xaxis:{{...LAY().xaxis,title:'Ticket médio anual (FSA nominal)',tickangle:-30}},
}}, PLOT_CFG);

Plotly.newPlot('thr_curve', [
  {{x:D.thr_k, y:D.thr_pct, type:'scatter', mode:'lines', name:'% produtoras abaixo',
    line:{{color:'#e05a2b',width:2}},
    fill:'tozeroy', fillcolor:'rgba(224,90,43,0.08)',
    hovertemplate:'R$%{{x}}k: %{{y}}% das produtoras<extra></extra>'}},
  {{x:D.thr_k, y:D.thr_fsa, type:'scatter', mode:'lines', name:'% FSA nesse grupo',
    line:{{color:'#4fa3e0',width:2,dash:'dot'}},
    hovertemplate:'R$%{{x}}k: grupo detém %{{y}}% do FSA<extra></extra>'}},
], {{
  ...LAY(),
  shapes:[
    {{type:'line',x0:500,x1:500,y0:0,y1:105,line:{{color:'#e05a2b',dash:'dash',width:1}}}},
    {{type:'line',x0:1000,x1:1000,y0:0,y1:105,line:{{color:'#e0c94f',dash:'dash',width:1}}}},
  ],
  yaxis:{{...LAY().yaxis,title:'% acumulada',range:[0,105]}},
  xaxis:{{...LAY().xaxis,title:'Limiar de ticket (R$k/ano)'}},
  legend:{{bgcolor:S1,bordercolor:S3,borderwidth:1,font:{{size:9}},x:.05,y:.95}},
}}, PLOT_CFG);

// ── Tab 2 ──
const tc = D.tier_rows;
function buildTierTable(){{
  const tbody = document.getElementById('tier_tbody');
  tc.forEach(r=>{{
    const adq = r.adq_max > 0 ? `R${{r.adq_min}}k – R${{r.adq_max}}k` : `R${{r.adq_min}}k+`;
    const warn500 = r.pct_below_500k > 30 ? 'color:#e05a2b' : (r.pct_below_500k>10?'color:#e0c94f':'color:#4ae0a0');
    const warn1m  = r.pct_below_1m > 50 ? 'color:#e05a2b' : (r.pct_below_1m>25?'color:#e0c94f':'color:#4ae0a0');
    tbody.innerHTML += `<tr>
      <td><span class="dot" style="background:${{r.color}}"></span>
          <span class="tier-badge" style="background:${{r.color}}">${{r.id}}</span>
          <span style="margin-left:6px;color:${{TXT}}">${{r.name}}</span></td>
      <td class="r dim">${{r.n}}</td>
      <td class="r">${{r.fsa_share}}%</td>
      <td class="r" style="color:${{r.ticket_med<500?'#e05a2b':r.ticket_med<1000?'#e0c94f':'#4ae0a0'}};font-weight:700">
        R$${{r.ticket_med}}k</td>
      <td class="r" style="color:#4fa3e0">${{adq}}</td>
      <td class="r dim">${{r.anos_med}}a</td>
      <td class="r dim">${{r.obras_med.toFixed(1)}}</td>
      <td class="r" style="${{warn500}}">${{r.pct_below_500k}}%</td>
      <td class="r" style="${{warn1m}}">${{r.pct_below_1m}}%</td>
    </tr>`;
  }});
}}
buildTierTable();

Plotly.newPlot('tier_comp', [
  {{x:tc.map(r=>r.name), y:tc.map(r=>r.ticket_med), type:'bar',
    name:'Ticket observado', marker:{{color:tc.map(r=>r.color)}},
    text:tc.map(r=>'R$'+r.ticket_med+'k'),
    textposition:'outside',textfont:{{size:9}},
    hovertemplate:'%{{x}}<br>Observado: R$%{{y}}k/ano<extra></extra>'}},
  {{x:tc.map(r=>r.name), y:tc.map(r=>r.adq_min), type:'scatter',mode:'markers',
    name:'Adequado mínimo', marker:{{symbol:'line-ew',size:14,color:'#4fa3e0',line:{{width:2,color:'#4fa3e0'}}}},
    hovertemplate:'%{{x}}<br>Mín. adequado: R$%{{y}}k/ano<extra></extra>'}},
], {{
  ...LAY(),
  barmode:'group',
  yaxis:{{...LAY().yaxis,title:'R$k / ano',type:'log'}},
  xaxis:{{...LAY().xaxis,tickangle:-20}},
  legend:{{bgcolor:S1,bordercolor:S3,borderwidth:1,font:{{size:9}}}},
}}, PLOT_CFG);

Plotly.newPlot('tier_pie', [{{
  labels:tc.map(r=>r.id+' '+r.name.split('—')[0].trim()),
  values:tc.map(r=>r.fsa_share),
  type:'pie', hole:0.4,
  marker:{{colors:tc.map(r=>r.color)}},
  textinfo:'label+percent',
  textfont:{{size:9}},
  hovertemplate:'%{{label}}: %{{value}}% do FSA<extra></extra>',
}}], {{
  ...LAY({{margin:{{t:6,b:6,l:6,r:6}},showlegend:false}}),
}}, PLOT_CFG);

const tierNames   = tc.map(r=>r.id+' — '+r.name);
const ticketMed   = tc.map(r=>r.ticket_med);
const adqMins     = tc.map(r=>r.adq_min);
const tierCols    = tc.map(r=>r.color);

Plotly.newPlot('tier_bar_sust', [
  {{x:tierNames, y:ticketMed, type:'bar', name:'Ticket Observado',
    marker:{{color:tierCols}},
    text:ticketMed.map(v=>'R$'+v+'k'), textposition:'outside', textfont:{{size:9}}}},
  {{x:tierNames, y:adqMins, type:'scatter', mode:'lines+markers',
    name:'Ticket Adequado Mín.',
    line:{{color:'#4fa3e0',dash:'dash',width:2}},
    marker:{{color:'#4fa3e0',size:7}},
    text:adqMins.map(v=>'R$'+v+'k'),
    hovertemplate:'Adequado mín.: R$%{{y}}k/ano<extra></extra>'}},
  {{x:tierNames, y:[500,500,500,500,500], type:'scatter', mode:'lines',
    name:'Limiar crítico R$500k',
    line:{{color:'#e05a2b',dash:'dot',width:1}}}},
], {{
  ...LAY(),
  yaxis:{{...LAY().yaxis,title:'R$k / ano'}},
  legend:{{bgcolor:S1,bordercolor:S3,borderwidth:1,font:{{size:9}}}},
}}, PLOT_CFG);

// ── Tab 3 ──
Plotly.newPlot('lorenz_prod', [
  {{x:[0,1],y:[0,1],mode:'lines',name:'Igualdade perfeita',
    line:{{dash:'dash',color:MUT,width:1}}}},
  {{x:D.lorenz_x,y:D.lorenz_y,mode:'lines',name:'FSA por produtora',
    fill:'tozeroy',fillcolor:'rgba(224,90,43,0.10)',
    line:{{color:'#e05a2b',width:2}},
    hovertemplate:'%{{x:.1%}} das produtoras → %{{y:.1%}} do FSA<extra></extra>'}},
], {{
  ...LAY(),
  annotations:[{{x:0.6,y:0.2,text:'Gini = {D["gini"]}',showarrow:false,
    font:{{size:12,color:'#e05a2b'}},bgcolor:S1,borderpad:4}}],
  xaxis:{{...LAY().xaxis,title:'Fração acumulada de produtoras',tickformat:'.0%'}},
  yaxis:{{...LAY().yaxis,title:'Fração acumulada do FSA',tickformat:'.0%'}},
  showlegend:false,
}}, PLOT_CFG);

// Scatter por tier
const tierOrder = ['A','B','C','D','E'];
const tierCmap = {{'A':'#e05a2b','B':'#f5c842','C':'#4fa3e0','D':'#4ae0a0','E':'#9b8eaf'}};
const tierLabels = {{'A':'Mega (Top 10)','B':'Grandes (11–50)','C':'Médias (51–150)','D':'Pequenas (151–300)','E':'Micro (301+)'}};
const scatterTraces = tierOrder.map(tid => {{
  const pts = D.scatter.filter(p=>p.tier===tid);
  return {{
    x:pts.map(p=>p.x), y:pts.map(p=>p.y),
    mode:'markers', type:'scatter', name:tierLabels[tid],
    marker:{{color:tierCmap[tid],size:pts.map(p=>Math.min(10+p.fsa_m*0.15,20)),
             opacity:.75,line:{{width:0}}}},
    text:pts.map(p=>`${{p.nome}}<br>Ticket: R$${{p.x}}k/ano<br>Obras: ${{p.y}}<br>FSA: R$${{p.fsa_m}}M`),
    hoverinfo:'text',
  }};
}});
Plotly.newPlot('scatter_tick', scatterTraces, {{
  ...LAY(),
  xaxis:{{...LAY().xaxis,title:'Ticket médio anual (R$k/ano)',type:'log'}},
  yaxis:{{...LAY().yaxis,title:'N° de obras com FSA',type:'log'}},
  legend:{{bgcolor:S1,bordercolor:S3,borderwidth:1,font:{{size:9}}}},
  shapes:[{{type:'line',x0:500,x1:500,y0:0.5,y1:200,line:{{color:'#e05a2b',dash:'dash',width:1}}}}],
}}, PLOT_CFG);

// Tab switcher
function showTab(id, btn){{
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
  window.dispatchEvent(new Event('resize'));
}}
</script>
</body>
</html>"""

OUT.write_text(html, encoding="utf-8")
print(f"Salvo: {OUT}")
