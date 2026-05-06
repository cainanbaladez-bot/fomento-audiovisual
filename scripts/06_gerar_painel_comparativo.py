# -*- coding: utf-8 -*-
"""
FASE 3 — Painel Comparativo de Mecanismos de Fomento
Compara: Renúncia Pura | FSA Puro | FSA+Renúncia (FSA Maj.) | FSA+Renúncia (Ren. Maj.)
"""

import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
from pathlib import Path
import json
import unicodedata
import re
import warnings
warnings.filterwarnings("ignore")

# ── Caminhos ──────────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent.parent
TABELA      = ROOT / "resultados" / "tabela_consolidada_obras.xlsx"
BASE_PROD   = ROOT / "resultados" / "datasets" / "base_nivel_produtora.csv"
BASE_OBRA   = ROOT / "resultados" / "datasets" / "base_nivel_obra.csv"
BASE_FEST_ATA = ROOT / "resultados" / "dataset" / "base_festivais_obras_ata.csv"
BRDE_FEST_LIST = ROOT / "dados" / "brde_festivais_lista.csv"
VOD_RAW     = ROOT / "raw" / "lumiere_vod_search.xlsx"
LUM_RAW     = ROOT / "raw" / "lumiere_search.xlsx"
OUTPUT = ROOT / "resultados" / "painel_comparativo.html"

# ── Paleta ────────────────────────────────────────────────────────────────────
CORES = {
    "Renúncia Pura":               "#E07B54",
    "FSA Puro":                    "#4C72B0",
    "FSA + Renúncia — FSA Maj.":   "#55A868",
    "FSA + Renúncia — Ren. Maj.":  "#8172B2",
}
COR_LISTA = list(CORES.values())

# ── Leitura ───────────────────────────────────────────────────────────────────
print("Carregando tabela consolidada...")
df = pd.read_excel(TABELA)

# Colunas financeiras (deflacionadas)
COL_FSA = "Valor FSA Deflac. (R$2024)"
COL_REN = "Renúncia Total Deflac. (R$2024)"
COL_INV = "Investimento Total Deflac. (R$2024)"
COL_BIL = "Bilheteria Deflac. (R$)"
COL_JAN = "Outras Janelas Deflac. (R$2024)"
COL_ROI_DOM = "ROI Dom. Total (deflac)"
COL_ROI_FSA = "ROI Dom. FSA (deflac)"
COL_ROI_INT = "ROI Internacional (0-100)"
COL_PAISES = "Total Países Alcançados"

# Garantir numérico
for col in [COL_FSA, COL_REN, COL_INV, COL_BIL, COL_JAN,
            COL_ROI_DOM, COL_ROI_FSA, COL_ROI_INT, COL_PAISES]:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# ── Classificação dos grupos ───────────────────────────────────────────────────
tem_fsa = df[COL_FSA] > 0
tem_ren = df[COL_REN] > 0

# Dentro de FSA+Renúncia: quem é majoritário?
soma = df[COL_FSA] + df[COL_REN]
fsa_share = np.where(soma > 0, df[COL_FSA] / soma, 0.5)

def classificar(row_idx):
    f = tem_fsa.iloc[row_idx]
    r = tem_ren.iloc[row_idx]
    if r and not f:
        return "Renúncia Pura"
    if f and not r:
        return "FSA Puro"
    if f and r:
        if fsa_share[row_idx] >= 0.5:
            return "FSA + Renúncia — FSA Maj."
        else:
            return "FSA + Renúncia — Ren. Maj."
    return None  # sem fomento — excluir

df["grupo"] = [classificar(i) for i in range(len(df))]
df = df[df["grupo"].notna()].copy()

# ── Exclusão global FASE 3: SAV/MINC, Arranjos Regionais e TV/PRODAV ─────────
# Removidas de toda a análise da Fase 3. O escopo da Fase 3 é PRODECINE
# (cinema), portanto a categoria "_tv_excluir" (conteúdo de TV/PRODAV, séries)
# também é removida — do contrário o "FSA Puro" seria dominado por séries que
# não têm lançamento comercial em sala por design.
_cat_str = df["Categoria"].fillna("").astype(str)
_mask_excluir_sav = _cat_str.str.contains("SAV", case=False, na=False) | \
                    _cat_str.str.contains("Arranjos", case=False, na=False)
_mask_excluir_tv  = _cat_str.str.contains("_tv_excluir", case=False, na=False)
_n_removidas = int((_mask_excluir_sav | _mask_excluir_tv).sum())
df = df[~(_mask_excluir_sav | _mask_excluir_tv)].copy()
print(f"Excluídas {_n_removidas} obras (SAV/MINC + Arranjos + _tv_excluir) — Fase 3")

# ── Recorte temporal FASE 3: janela de 10 anos (2014–2023) ────────────────────
# Antes de 2014 o FSA tem volume marginal de obras finalizadas, o que torna a
# comparação contra Renúncia desequilibrada. Restringimos a análise a um
# decênio onde todos os grupos coexistem com massa estatística.
ANO_INI, ANO_FIM = 2014, 2023
df["Ano"] = pd.to_numeric(df["Ano"], errors="coerce")
_n_antes = len(df)
df = df[df["Ano"].between(ANO_INI, ANO_FIM)].copy()
print(f"Recorte temporal {ANO_INI}–{ANO_FIM}: {len(df)} obras "
      f"(removidas {_n_antes - len(df)} fora da janela)")

# ── Filtro de exibição comercial: só obras que estrearam em cartaz no Brasil ──
# Escopo da análise é retorno cinematográfico; obras sem bilheteria registrada
# (lançamento direto em TV/streaming ou sem lançamento) são excluídas do painel.
_n_antes_bil = len(df)
df = df[df[COL_BIL] > 0].copy()
print(f"Filtro bilheteria>0: {len(df)} obras "
      f"(removidas {_n_antes_bil - len(df)} sem lançamento em cartaz)")

GRUPOS = list(CORES.keys())
ORDEM = {g: i for i, g in enumerate(GRUPOS)}
df["grupo_ord"] = df["grupo"].map(ORDEM)

print(f"Total obras classificadas: {len(df)}")
print(df["grupo"].value_counts())

# ── Receita total ─────────────────────────────────────────────────────────────
df["receita_total"] = df[COL_BIL]

# ── Funções de agregação ──────────────────────────────────────────────────────
def stats(g):
    s = {}
    s["n_obras"]          = len(g)
    s["inv_total_bi"]     = g[COL_INV].sum() / 1e9
    s["fsa_total_bi"]     = g[COL_FSA].sum() / 1e9
    s["ren_total_bi"]     = g[COL_REN].sum() / 1e9
    s["bil_total_bi"]     = g[COL_BIL].sum() / 1e9
    s["receita_total_bi"] = g["receita_total"].sum() / 1e9
    # ROI doméstico: média ponderada pelo investimento (todas as obras têm BIL>0)
    mask_dom = g[COL_INV] > 0
    if mask_dom.sum() > 0:
        w = g.loc[mask_dom, COL_INV]
        s["roi_dom_med"] = float((g.loc[mask_dom, COL_ROI_DOM] * w).sum() / w.sum())
    else:
        s["roi_dom_med"] = 0
    # ROI internacional: média simples INCONDICIONAL (todas as obras, incluindo zeros)
    s["roi_int_med"]      = float(g[COL_ROI_INT].mean())
    s["pct_com_bil"]      = (g[COL_BIL] > 0).mean() * 100
    s["pct_com_intl"]     = (g[COL_ROI_INT] >= 13).mean() * 100
    # Países: média simples incondicional
    s["paises_med"]       = float(g[COL_PAISES].mean())
    s["inv_medio_mi"]     = g[COL_INV].mean() / 1e6
    s["bil_medio_mi"]     = g[COL_BIL].mean() / 1e6
    # ROI agregado (receita total / inv total)
    inv_sum = g[COL_INV].sum()
    rec_sum = g["receita_total"].sum()
    s["roi_agregado"]     = rec_sum / inv_sum if inv_sum > 0 else 0
    return s

resumo = {g: stats(df[df["grupo"] == g]) for g in GRUPOS}

# ── Layout HTML ───────────────────────────────────────────────────────────────
def fmt_bi(v):  return f"R$ {v:.2f} bi"
def fmt_mi(v):  return f"R$ {v:.1f} mi"
def fmt_pct(v): return f"{v:.1f}%"
def fmt_x(v):   return f"{v:.2f}x"
def fmt_n(v):   return f"{int(v):,}".replace(",", ".")

# ── Tabela resumo ─────────────────────────────────────────────────────────────
def hex_to_rgba(hex_color, alpha=0.15):
    """Converte '#RRGGBB' para 'rgba(r,g,b,alpha)'."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def make_tabela_resumo():
    headers = [
        "Grupo", "N obras", "Inv. Total", "Bilheteria",
        "ROI Dom.<br>(média pond.)", "ROI Intl<br>(média)",
        "% c/ Intl",
    ]
    rows = []
    for g in GRUPOS:
        s = resumo[g]
        rows.append([
            g,
            fmt_n(s["n_obras"]),
            fmt_bi(s["inv_total_bi"]),
            fmt_bi(s["bil_total_bi"]),
            fmt_x(s["roi_dom_med"]),
            f"{s['roi_int_med']:.1f}",
            fmt_pct(s["pct_com_intl"]),
        ])

    col_vals = list(zip(*rows))
    n_grupos = len(GRUPOS)
    row_colors = [hex_to_rgba(CORES[g]) for g in GRUPOS]
    fill_colors = [row_colors] + [["#f8f9fa"] * n_grupos] * (len(headers) - 1)

    fig = go.Figure(go.Table(
        columnwidth=[220, 80, 110, 110, 130, 110, 90],
        header=dict(
            values=[f"<b>{h}</b>" for h in headers],
            fill_color="#2c3e50",
            font=dict(color="white", size=12),
            align="center",
            height=36,
        ),
        cells=dict(
            values=col_vals,
            fill_color=fill_colors,
            font=dict(size=12),
            align=["left"] + ["center"] * (len(headers) - 1),
            height=30,
        ),
    ))
    fig.update_layout(
        margin=dict(l=0, r=0, t=10, b=0),
        height=210,
    )
    return fig


# ── Gráfico: Investimento x Bilheteria (barras empilhadas) ─────────────────────
def make_inv_bil():
    fig = go.Figure()
    inv_vals = [resumo[g]["inv_total_bi"] for g in GRUPOS]
    bil_vals = [resumo[g]["bil_total_bi"] for g in GRUPOS]
    rec_vals = [resumo[g]["receita_total_bi"] for g in GRUPOS]

    fig.add_trace(go.Bar(
        name="Investimento Total", x=GRUPOS, y=inv_vals,
        marker_color=[CORES[g] for g in GRUPOS],
        opacity=0.5,
        text=[fmt_bi(v) for v in inv_vals],
        textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name="Bilheteria", x=GRUPOS, y=bil_vals,
        marker_color=[CORES[g] for g in GRUPOS],
        marker_pattern_shape="/",
        text=[fmt_bi(v) for v in bil_vals],
        textposition="outside",
    ))
    fig.add_trace(go.Scatter(
        name="Bilheteria (cinema)", x=GRUPOS, y=rec_vals,
        mode="markers+lines",
        marker=dict(size=12, symbol="diamond", color="black"),
        line=dict(color="black", dash="dot"),
    ))
    fig.update_layout(
        barmode="group",
        title="Investimento vs. Retorno por Grupo (R$ bilhões, R$2024)",
        xaxis_title="", yaxis_title="R$ bilhões",
        legend=dict(orientation="h", y=-0.15),
        height=420, margin=dict(l=50, r=20, t=50, b=80),
    )
    return fig


# ── Gráfico: ROI doméstico — box plot (log) ──────────────────────────────────
def make_roi_dom_box():
    fig = go.Figure()
    for g in GRUPOS:
        vals = df[(df["grupo"] == g) & (df[COL_ROI_DOM] > 0)][COL_ROI_DOM]
        fig.add_trace(go.Box(
            y=vals, name=g,
            marker_color=CORES[g],
            boxmean="sd",
        ))
    fig.update_layout(
        title="Distribuição do ROI Doméstico por Grupo<br><sup>(obras com ROI > 0 · escala log)</sup>",
        yaxis=dict(title="ROI Doméstico (receita/invest.)", type="log"),
        showlegend=False,
        height=420, margin=dict(l=60, r=20, t=70, b=50),
    )
    return fig


# ── Gráfico: ROI Internacional — box plot (log, base 10) ─────────────────────
def make_roi_int_box():
    fig = go.Figure()
    for g in GRUPOS:
        vals = df[(df["grupo"] == g) & (df[COL_ROI_INT] >= 13)][COL_ROI_INT]
        fig.add_trace(go.Box(
            y=vals, name=g,
            marker_color=CORES[g],
            boxmean="sd",
        ))
    fig.update_layout(
        title="Distribuição do ROI Internacional por Grupo<br><sup>(obras com ROI Intl ≥ 13 · escala log)</sup>",
        yaxis=dict(title="ROI Internacional (0–100)", type="log"),
        showlegend=False,
        height=420, margin=dict(l=60, r=20, t=70, b=50),
    )
    return fig


# ── Gráfico: Dispersão Investimento × ROI (Dom ou Intl) — nível obra ──────────
def make_scatter_inv_roi():
    fig = go.Figure()
    inv_mi = df[COL_INV] / 1e6  # R$ milhões

    # Traços ROI Doméstico (visíveis por padrão)
    for g in GRUPOS:
        mask = df["grupo"] == g
        sub = df[mask]
        fig.add_trace(go.Scatter(
            x=sub[COL_INV] / 1e6,
            y=sub[COL_ROI_DOM],
            mode="markers",
            name=g,
            legendgroup=g,
            marker=dict(color=CORES[g], size=6, opacity=0.7,
                        line=dict(width=0.4, color="white")),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Ano: %{customdata[1]}<br>"
                "Inv: R$ %{x:.1f} mi<br>"
                "ROI Dom: %{y:.2f}x<extra></extra>"
            ),
            customdata=sub[["Projeto", "Ano"]].values,
            visible=True,
        ))

    # Traços ROI Internacional (ocultos por padrão)
    for g in GRUPOS:
        mask = df["grupo"] == g
        sub = df[mask]
        fig.add_trace(go.Scatter(
            x=sub[COL_INV] / 1e6,
            y=sub[COL_ROI_INT],
            mode="markers",
            name=g,
            legendgroup=g,
            showlegend=False,
            marker=dict(color=CORES[g], size=6, opacity=0.7,
                        line=dict(width=0.4, color="white")),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Ano: %{customdata[1]}<br>"
                "Inv: R$ %{x:.1f} mi<br>"
                "ROI Intl: %{y:.1f}<extra></extra>"
            ),
            customdata=sub[["Projeto", "Ano"]].values,
            visible=False,
        ))

    n = len(GRUPOS)
    dom_vis  = [True] * n + [False] * n
    intl_vis = [False] * n + [True] * n

    fig.update_layout(
        title="Investimento × ROI por Obra e Grupo (R$2024)",
        xaxis=dict(title="Investimento Total (R$ milhões, R$2024)", type="log"),
        yaxis=dict(title="ROI Doméstico (receita/invest.)", type="log"),
        legend=dict(orientation="h", y=-0.12),
        height=460, margin=dict(l=70, r=20, t=40, b=80),
    )
    return fig


# ── Gráfico: Composição do investimento (FSA vs Renúncia) ─────────────────────
def make_composicao_inv():
    fig = go.Figure()
    fsa_vals = [resumo[g]["fsa_total_bi"] for g in GRUPOS]
    ren_vals = [resumo[g]["ren_total_bi"] for g in GRUPOS]

    fig.add_trace(go.Bar(
        name="FSA", x=GRUPOS, y=fsa_vals,
        marker_color="#4C72B0",
        text=[fmt_bi(v) for v in fsa_vals],
        textposition="inside", textfont_color="white",
    ))
    fig.add_trace(go.Bar(
        name="Renúncia Fiscal", x=GRUPOS, y=ren_vals,
        marker_color="#E07B54",
        text=[fmt_bi(v) for v in ren_vals],
        textposition="inside", textfont_color="white",
    ))
    fig.update_layout(
        barmode="stack",
        title="Composição do Investimento por Grupo (R$ bilhões, R$2024)",
        yaxis_title="R$ bilhões",
        legend=dict(orientation="h", y=-0.15),
        height=380, margin=dict(l=50, r=20, t=50, b=80),
    )
    return fig


# ── Gráfico: ROI agregado — gauge / barras ────────────────────────────────────
def make_roi_agregado():
    rois = [resumo[g]["roi_agregado"] for g in GRUPOS]
    fig = go.Figure(go.Bar(
        x=GRUPOS, y=rois,
        marker_color=[CORES[g] for g in GRUPOS],
        text=[f"{v:.3f}x" for v in rois],
        textposition="outside",
    ))
    fig.add_hline(y=1.0, line_dash="dash", line_color="red",
                  annotation_text="Break-even (1x)", annotation_position="top right")
    fig.update_layout(
        title="ROI Agregado por Grupo<br><sup>(Receita Total / Investimento Total — R$2024)</sup>",
        yaxis_title="ROI Agregado",
        showlegend=False,
        height=380, margin=dict(l=50, r=20, t=70, b=50),
    )
    return fig


# ── Gráfico: N países alcançados (log) ────────────────────────────────────────
def make_paises():
    fig = go.Figure()
    for g in GRUPOS:
        vals = df[(df["grupo"] == g) & (df[COL_PAISES] > 0)][COL_PAISES]
        fig.add_trace(go.Box(
            y=vals, name=g, marker_color=CORES[g],
            boxmean="sd",
        ))
    fig.update_layout(
        title="Distribuição de Países Alcançados por Grupo<br><sup>(obras com presença intl · escala log)</sup>",
        yaxis=dict(title="Total de Países", type="log"),
        showlegend=False,
        height=420, margin=dict(l=60, r=20, t=70, b=50),
    )
    return fig


# ── Gráfico: Scatter ROI Dom x ROI Intl (log-log) ────────────────────────────
def make_scatter_roi():
    df_plot = df[(df[COL_ROI_DOM] > 0) & (df[COL_ROI_INT] >= 13)].copy()

    fig = go.Figure()
    for g in GRUPOS:
        sub = df_plot[df_plot["grupo"] == g]
        fig.add_trace(go.Scatter(
            x=sub[COL_ROI_INT], y=sub[COL_ROI_DOM],
            mode="markers",
            name=g,
            marker=dict(color=CORES[g], size=7, opacity=0.65,
                        line=dict(width=0.4, color="white")),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Ano: %{customdata[1]}<br>"
                "ROI Dom: %{y:.2f}x<br>"
                "ROI Intl: %{x:.1f}<extra></extra>"
            ),
            customdata=sub[["Projeto", "Ano"]].values,
        ))
    fig.add_hline(y=1, line_dash="dot", line_color="gray",
                  annotation_text="ROI dom = 1x (break-even)", annotation_position="right")
    fig.update_layout(
        title="ROI Doméstico × ROI Internacional por Obra e Grupo<br><sup>(ambos os eixos em escala log)</sup>",
        xaxis=dict(title="ROI Internacional (0–100, log)", type="log"),
        yaxis=dict(title="ROI Doméstico (log)", type="log"),
        legend=dict(orientation="h", y=-0.15),
        height=560, margin=dict(l=70, r=20, t=70, b=80),
    )
    return fig


# ── Gráfico: Investimento médio por obra ──────────────────────────────────────
def make_inv_medio():
    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=["Investimento médio por obra (R$ mi)", "Bilheteria média por obra (R$ mi)"])
    inv_med = [resumo[g]["inv_medio_mi"] for g in GRUPOS]
    bil_med = [resumo[g]["bil_medio_mi"] for g in GRUPOS]

    for vals, col in [(inv_med, 1), (bil_med, 2)]:
        fig.add_trace(go.Bar(
            x=GRUPOS, y=vals,
            marker_color=[CORES[g] for g in GRUPOS],
            text=[f"R$ {v:.1f}mi" for v in vals],
            textposition="outside",
            showlegend=False,
        ), row=1, col=col)

    fig.update_layout(
        title="Escala por Obra (valores médios, R$2024)",
        height=380, margin=dict(l=50, r=20, t=60, b=100),
        xaxis_tickangle=-30, xaxis2_tickangle=-30,
    )
    return fig


# ── Análise por Categoria de Chamada ──────────────────────────────────────────
CAT_EXCLUIR = {"_tv_excluir", "sem_categoria"}
df_cat = df[
    ~df["Categoria"].isin(CAT_EXCLUIR) &
    df["Categoria"].notna() &
    ~df["Categoria"].fillna("").str.startswith("Renúncia")
].copy()

# Paleta por família (FSA em tons azul/verde, Renúncia em tons laranja/vermelho)
def cor_categoria(cat):
    if cat.startswith("Renúncia"):
        return "#E07B54" if "Art.3" in cat else "#D95F3A"
    if "SAV" in cat or "Arranjos" in cat:
        return "#8172B2"
    if "Bilheteria e Roteiro" in cat and "Produtora" in cat:
        return "#4C72B0"
    if "Bilheteria e Roteiro" in cat and "Distribuidora" in cat:
        return "#6B8FC7"
    if "Festivais e Roteiro" in cat:
        return "#55A868"
    if "Complementação" in cat:
        return "#77B8A0"
    if "Automático Bilheteria" in cat:
        return "#3A7CA8"
    if "Automático Festivais" in cat:
        return "#2E8B57"
    if "Coprodução Internacional" in cat:
        return "#1F6091"
    if "Comercialização" in cat or "Distribuição" in cat:
        return "#9CAFC2"
    if "Apenas roteiro" in cat:
        return "#B0B0B0"
    return "#888888"

def stats_cat(g):
    inv_sum = g[COL_INV].sum()
    rec_sum = g["receita_total"].sum()
    # ROI doméstico: média ponderada pelo investimento (mesma convenção de stats())
    mask_dom = g[COL_INV] > 0
    if mask_dom.sum() > 0:
        w = g.loc[mask_dom, COL_INV]
        roi_dom_med = float((g.loc[mask_dom, COL_ROI_DOM] * w).sum() / w.sum())
    else:
        roi_dom_med = 0
    return {
        "n_obras":        len(g),
        "n_com_intl":     int((g[COL_ROI_INT] >= 13).sum()),
        "inv_total_bi":   inv_sum / 1e9,
        "inv_medio_mi":   g[COL_INV].mean() / 1e6,
        "bil_total_bi":   g[COL_BIL].sum() / 1e9,
        "receita_bi":     rec_sum / 1e9,
        "roi_agregado":   rec_sum / inv_sum if inv_sum > 0 else 0,
        "roi_dom_med":    roi_dom_med,
        # ROI internacional: média incondicional (mesma convenção de stats())
        "roi_int_medio":  float(g[COL_ROI_INT].mean()),
        # ROI internacional: soma absoluta dos pontos de alcance intl (volume total)
        "roi_int_total":  g[COL_ROI_INT].sum(),
        "pct_com_bil":    (g[COL_BIL] > 0).mean() * 100,
        "pct_com_intl":   (g[COL_ROI_INT] >= 13).mean() * 100,
    }

cat_resumo = {c: stats_cat(df_cat[df_cat["Categoria"] == c]) for c in df_cat["Categoria"].unique()}
# Ordenar por n_obras desc
CATS_ORD = sorted(cat_resumo.keys(), key=lambda c: -cat_resumo[c]["n_obras"])
CAT_CORES = {c: cor_categoria(c) for c in CATS_ORD}

def make_cat_volumes():
    """Barras horizontais: N obras + Investimento total por categoria."""
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=["Nº de obras por categoria", "Investimento total (R$ bi, 2024)"],
        horizontal_spacing=0.18,
    )
    cats = CATS_ORD[::-1]  # inverter p/ maior no topo
    n_obras = [cat_resumo[c]["n_obras"] for c in cats]
    inv_bi  = [cat_resumo[c]["inv_total_bi"] for c in cats]
    cores   = [CAT_CORES[c] for c in cats]

    fig.add_trace(go.Bar(
        y=cats, x=n_obras, orientation="h",
        marker_color=cores,
        text=[fmt_n(v) for v in n_obras], textposition="outside",
        showlegend=False,
    ), row=1, col=1)

    fig.add_trace(go.Bar(
        y=cats, x=inv_bi, orientation="h",
        marker_color=cores,
        text=[f"R$ {v:.2f}bi" for v in inv_bi], textposition="outside",
        showlegend=False,
    ), row=1, col=2)

    fig.update_layout(
        title="Volume de obras e investimento por Categoria de Chamada",
        height=max(420, 32 * len(cats) + 120),
        margin=dict(l=20, r=40, t=70, b=40),
    )
    fig.update_yaxes(automargin=True)
    return fig

def make_cat_roi():
    """Barras: ROI agregado + ROI doméstico mediano + ROI intl médio + ROI intl total."""
    fig = make_subplots(
        rows=1, cols=4,
        subplot_titles=[
            "ROI agregado (receita / invest.)",
            "ROI doméstico mediano",
            "ROI intl médio <br><sub>(obras c/ alcance intl, n em cima)</sub>",
            "ROI intl TOTAL (soma pontos)",
        ],
        horizontal_spacing=0.12,
    )
    cats = CATS_ORD[::-1]
    roi_agr     = [cat_resumo[c]["roi_agregado"]  for c in cats]
    roi_dom     = [cat_resumo[c]["roi_dom_med"]   for c in cats]
    roi_int_med = [cat_resumo[c]["roi_int_medio"] for c in cats]
    roi_int_sum = [cat_resumo[c]["roi_int_total"] for c in cats]
    n_intl      = [cat_resumo[c]["n_com_intl"]    for c in cats]
    cores       = [CAT_CORES[c] for c in cats]

    fig.add_trace(go.Bar(
        y=cats, x=roi_agr, orientation="h", marker_color=cores,
        text=[f"{v:.2f}x" for v in roi_agr], textposition="outside",
        showlegend=False,
    ), row=1, col=1)

    fig.add_trace(go.Bar(
        y=cats, x=roi_dom, orientation="h", marker_color=cores,
        text=[f"{v:.2f}x" for v in roi_dom], textposition="outside",
        showlegend=False,
    ), row=1, col=2)

    fig.add_trace(go.Bar(
        y=cats, x=roi_int_med, orientation="h", marker_color=cores,
        text=[f"{v:.1f} (n={n})" for v, n in zip(roi_int_med, n_intl)],
        textposition="outside", showlegend=False,
    ), row=1, col=3)

    fig.add_trace(go.Bar(
        y=cats, x=roi_int_sum, orientation="h", marker_color=cores,
        text=[f"{v:.0f}" for v in roi_int_sum], textposition="outside",
        showlegend=False,
    ), row=1, col=4)

    fig.update_layout(
        title="Retorno por Categoria de Chamada",
        height=max(460, 34 * len(cats) + 160),
        margin=dict(l=20, r=40, t=90, b=40),
    )
    fig.update_yaxes(automargin=True)
    return fig

def make_cat_tabela():
    """Tabela resumo completa por categoria."""
    cats = CATS_ORD
    header = ["Categoria", "Nº Obras", "Inv. Total", "Inv. Médio",
              "Receita Tot.", "ROI Agr.", "ROI Dom. Med.",
              "N c/ Intl.", "ROI Intl Méd.", "ROI Intl TOTAL",
              "% c/ Bilh.", "% c/ Intl."]
    cells = [
        cats,
        [fmt_n(cat_resumo[c]["n_obras"]) for c in cats],
        [fmt_bi(cat_resumo[c]["inv_total_bi"]) for c in cats],
        [fmt_mi(cat_resumo[c]["inv_medio_mi"]) for c in cats],
        [fmt_bi(cat_resumo[c]["receita_bi"]) for c in cats],
        [fmt_x(cat_resumo[c]["roi_agregado"]) for c in cats],
        [fmt_x(cat_resumo[c]["roi_dom_med"]) for c in cats],
        [fmt_n(cat_resumo[c]["n_com_intl"]) for c in cats],
        [f"{cat_resumo[c]['roi_int_medio']:.1f}" for c in cats],
        [f"{cat_resumo[c]['roi_int_total']:.0f}" for c in cats],
        [fmt_pct(cat_resumo[c]["pct_com_bil"]) for c in cats],
        [fmt_pct(cat_resumo[c]["pct_com_intl"]) for c in cats],
    ]
    cores_row = [hex_to_rgba(CAT_CORES[c], 0.15) for c in cats]
    fig = go.Figure(go.Table(
        header=dict(
            values=[f"<b>{h}</b>" for h in header],
            fill_color="#2c3e50", font=dict(color="white", size=11),
            align="center", height=34,
        ),
        cells=dict(
            values=cells,
            fill_color=[cores_row] * len(header),
            align=["left"] + ["right"] * (len(header) - 1),
            font=dict(size=10), height=28,
        ),
        columnwidth=[280, 60, 80, 80, 80, 60, 80, 60, 75, 85, 70, 70],
    ))
    fig.update_layout(
        title="Tabela detalhada por Categoria de Chamada",
        height=max(360, 30 * len(cats) + 140),
        margin=dict(l=10, r=10, t=60, b=10),
    )
    return fig

def make_cat_roi_box():
    """Box plot de ROI doméstico por categoria (apenas obras com bilheteria, log)."""
    fig = go.Figure()
    for c in CATS_ORD:
        g = df_cat[(df_cat["Categoria"] == c) & (df_cat[COL_ROI_DOM] > 0)]
        if len(g) < 3:
            continue
        fig.add_trace(go.Box(
            y=g[COL_ROI_DOM],
            name=c,
            marker_color=CAT_CORES[c],
            boxmean="sd",
        ))
    fig.update_layout(
        title="Distribuição do ROI Doméstico por Categoria<br><sup>(obras com bilheteria · escala log)</sup>",
        yaxis=dict(title="ROI doméstico (receita/investimento)", type="log"),
        height=520, margin=dict(l=60, r=20, t=80, b=200),
        xaxis_tickangle=-35, showlegend=False,
    )
    return fig

def make_cat_roi_int_box():
    """Box plot de ROI internacional por categoria (apenas obras com alcance intl, log)."""
    fig = go.Figure()
    for c in CATS_ORD:
        g = df_cat[(df_cat["Categoria"] == c) & (df_cat[COL_ROI_INT] >= 13)]
        if len(g) < 3:
            continue
        fig.add_trace(go.Box(
            y=g[COL_ROI_INT],
            name=c,
            marker_color=CAT_CORES[c],
            boxmean="sd",
        ))
    fig.update_layout(
        title="Distribuição do ROI Internacional por Categoria<br><sup>(obras com alcance intl · escala log)</sup>",
        yaxis=dict(title="ROI internacional (0–100)", type="log"),
        height=520, margin=dict(l=60, r=20, t=80, b=200),
        xaxis_tickangle=-35, showlegend=False,
    )
    return fig


def make_scatter_cat():
    """Scatter 3 modos (sem Renúncia Pura):
       Modo 0 (default): ROI Dom Total Deflac × ROI Internacional
       Modo 1: Investimento Total × ROI Dom Total Deflac
       Modo 2: Investimento Total × ROI Internacional
    """
    GRUPOS_CAT = [g for g in GRUPOS if not g.startswith("Renúncia")]
    df_s = df[df["grupo"].isin(GRUPOS_CAT)].copy()
    n = len(GRUPOS_CAT)
    fig = go.Figure()

    # Modo 0: x=ROI Intl, y=ROI Dom (visível por padrão)
    for g in GRUPOS_CAT:
        sub = df_s[(df_s["grupo"] == g) & (df_s[COL_ROI_INT] >= 13) & (df_s[COL_ROI_DOM] > 0)]
        fig.add_trace(go.Scatter(
            x=sub[COL_ROI_INT], y=sub[COL_ROI_DOM],
            mode="markers", name=g, legendgroup=g,
            marker=dict(color=CORES[g], size=7, opacity=0.65,
                        line=dict(width=0.4, color="white")),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>Ano: %{customdata[1]}<br>"
                "ROI Dom: %{y:.2f}x<br>ROI Intl: %{x:.1f}<extra></extra>"
            ),
            customdata=sub[["Projeto", "Ano"]].values,
            visible=True,
        ))

    # Modo 1: x=Inv Total, y=ROI Dom (oculto por padrão)
    for g in GRUPOS_CAT:
        sub = df_s[(df_s["grupo"] == g) & (df_s[COL_INV] > 0) & (df_s[COL_ROI_DOM] > 0)]
        fig.add_trace(go.Scatter(
            x=sub[COL_INV] / 1e6, y=sub[COL_ROI_DOM],
            mode="markers", name=g, legendgroup=g, showlegend=False,
            marker=dict(color=CORES[g], size=7, opacity=0.65,
                        line=dict(width=0.4, color="white")),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>Ano: %{customdata[1]}<br>"
                "Inv: R$ %{x:.1f}mi<br>ROI Dom: %{y:.2f}x<extra></extra>"
            ),
            customdata=sub[["Projeto", "Ano"]].values,
            visible=False,
        ))

    # Modo 2: x=Inv Total, y=ROI Intl (oculto por padrão)
    for g in GRUPOS_CAT:
        sub = df_s[(df_s["grupo"] == g) & (df_s[COL_INV] > 0) & (df_s[COL_ROI_INT] >= 13)]
        fig.add_trace(go.Scatter(
            x=sub[COL_INV] / 1e6, y=sub[COL_ROI_INT],
            mode="markers", name=g, legendgroup=g, showlegend=False,
            marker=dict(color=CORES[g], size=7, opacity=0.65,
                        line=dict(width=0.4, color="white")),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>Ano: %{customdata[1]}<br>"
                "Inv: R$ %{x:.1f}mi<br>ROI Intl: %{y:.1f}<extra></extra>"
            ),
            customdata=sub[["Projeto", "Ano"]].values,
            visible=False,
        ))

    vis0 = [True]*n + [False]*n + [False]*n
    vis1 = [False]*n + [True]*n + [False]*n
    vis2 = [False]*n + [False]*n + [True]*n

    fig.update_layout(
        title="Dispersão por Grupo — FSA (exclui Renúncia Pura)",
        updatemenus=[dict(
            type="buttons", direction="right",
            x=0.0, y=1.13, xanchor="left",
            showactive=True,
            buttons=[
                dict(
                    label="ROI Dom × ROI Intl",
                    method="update",
                    args=[
                        {"visible": vis0},
                        {"xaxis.title.text": "ROI Internacional (0–100, log)",
                         "xaxis.type": "log",
                         "yaxis.title.text": "ROI Doméstico (log)",
                         "yaxis.type": "log"},
                    ],
                ),
                dict(
                    label="Inv. Total × ROI Dom",
                    method="update",
                    args=[
                        {"visible": vis1},
                        {"xaxis.title.text": "Investimento Total (R$ milhões, log)",
                         "xaxis.type": "log",
                         "yaxis.title.text": "ROI Doméstico (log)",
                         "yaxis.type": "log"},
                    ],
                ),
                dict(
                    label="Inv. Total × ROI Intl",
                    method="update",
                    args=[
                        {"visible": vis2},
                        {"xaxis.title.text": "Investimento Total (R$ milhões, log)",
                         "xaxis.type": "log",
                         "yaxis.title.text": "ROI Internacional (0–100, log)",
                         "yaxis.type": "log"},
                    ],
                ),
            ],
        )],
        xaxis=dict(title="ROI Internacional (0–100, log)", type="log"),
        yaxis=dict(title="ROI Doméstico (log)", type="log"),
        legend=dict(orientation="h", y=-0.15),
        height=560, margin=dict(l=70, r=20, t=90, b=80),
    )
    fig.add_hline(y=1, line_dash="dot", line_color="gray",
                  annotation_text="ROI dom = 1x (break-even)", annotation_position="right")
    return fig


# ── Análise por Cluster de Produtora ──────────────────────────────────────────
print("Carregando base nível produtora...")
df_prod = pd.read_csv(BASE_PROD, sep=";")

# Ordenação lógica dos clusters (maior retorno → menor)
CLUSTERS_ORD = [
    "Duplo Retorno",
    "Retorno Doméstico",
    "Retorno Internacional",
    "Fomento Baixo Retorno",
    "Outros",
]
# Normaliza acentos (o CSV pode ter vindo em codificação diferente)
CLUSTER_ALIAS = {
    "Retorno Dom\u00e9stico": "Retorno Doméstico",
    "Retorno Dom\ufffdstico": "Retorno Doméstico",
}
df_prod["cluster"] = df_prod["cluster"].replace(CLUSTER_ALIAS)
df_prod = df_prod[df_prod["cluster"].isin(CLUSTERS_ORD)].copy()

CLUSTER_CORES = {
    "Duplo Retorno":         "#2E8B57",  # verde escuro — melhor
    "Retorno Doméstico":     "#4C72B0",  # azul
    "Retorno Internacional": "#8172B2",  # roxo
    "Fomento Baixo Retorno": "#D95F3A",  # vermelho
    "Outros":                "#A0A0A0",  # cinza — residual
}

def stats_cluster(g):
    inv_sum = g["investimento_total_deflac"].sum()
    rec_sum = g["receita_total_deflac"].sum()
    fsa_sum = g["investimento_fsa_deflac"].sum()
    return {
        "n_prod":         len(g),
        "n_obras":        int(g["n_obras"].sum()),
        "inv_total_bi":   inv_sum / 1e9,
        "fsa_total_bi":   fsa_sum / 1e9,
        "rec_total_bi":   rec_sum / 1e9,
        "roi_agregado":   rec_sum / inv_sum if inv_sum > 0 else 0,
        "inv_medio_prod_mi": (inv_sum / len(g) / 1e6) if len(g) > 0 else 0,
        "obras_por_prod": g["n_obras"].mean(),
        "roi_dom_med":    float((g["roi_dom_total_deflac"] * g["investimento_total_deflac"]).sum() / g["investimento_total_deflac"].sum()) if g["investimento_total_deflac"].sum() > 0 else 0,
        "roi_intl_medio": float(g["roi_intl_medio"].mean()),
        "roi_intl_total": g["roi_intl_medio"].sum(),
        "pct_com_intl":   (g["roi_intl_max"] >= 13).mean() * 100,
        "critica_med":    g["critica_media"].dropna().mean() if g["critica_media"].notna().any() else 0,
        "pct_fem_med":    g["pct_obras_genero_feminino"].dropna().mean() if g["pct_obras_genero_feminino"].notna().any() else 0,
    }

cluster_resumo = {c: stats_cluster(df_prod[df_prod["cluster"] == c]) for c in CLUSTERS_ORD}

def make_cluster_kpi():
    """Barras: n produtoras + n obras + investimento + receita por cluster."""
    fig = make_subplots(
        rows=1, cols=4,
        subplot_titles=[
            "Nº de produtoras",
            "Nº de obras (soma)",
            "Investimento total (R$ bi)",
            "Receita total (R$ bi)",
        ],
        horizontal_spacing=0.08,
    )
    cats = CLUSTERS_ORD
    cores = [CLUSTER_CORES[c] for c in cats]

    vals_n  = [cluster_resumo[c]["n_prod"]      for c in cats]
    vals_ob = [cluster_resumo[c]["n_obras"]     for c in cats]
    vals_iv = [cluster_resumo[c]["inv_total_bi"] for c in cats]
    vals_rc = [cluster_resumo[c]["rec_total_bi"] for c in cats]

    for vals, col, fmt in [
        (vals_n,  1, lambda v: fmt_n(v)),
        (vals_ob, 2, lambda v: fmt_n(v)),
        (vals_iv, 3, lambda v: f"R$ {v:.2f}bi"),
        (vals_rc, 4, lambda v: f"R$ {v:.2f}bi"),
    ]:
        fig.add_trace(go.Bar(
            x=cats, y=vals, marker_color=cores,
            text=[fmt(v) for v in vals], textposition="outside",
            showlegend=False,
        ), row=1, col=col)

    fig.update_layout(
        title="Volume por Cluster de Produtora",
        height=440, margin=dict(l=40, r=20, t=80, b=120),
    )
    fig.update_xaxes(tickangle=-25)
    return fig

def make_cluster_roi():
    """ROI agregado + ROI dom mediano + ROI intl médio + ROI intl total."""
    fig = make_subplots(
        rows=1, cols=4,
        subplot_titles=[
            "ROI agregado (rec/inv)",
            "ROI doméstico mediano",
            "ROI intl médio (prod. c/ intl)",
            "ROI intl TOTAL (soma)",
        ],
        horizontal_spacing=0.1,
    )
    cats = CLUSTERS_ORD
    cores = [CLUSTER_CORES[c] for c in cats]

    roi_agr  = [cluster_resumo[c]["roi_agregado"]  for c in cats]
    roi_dom  = [cluster_resumo[c]["roi_dom_med"]   for c in cats]
    roi_im   = [cluster_resumo[c]["roi_intl_medio"] for c in cats]
    roi_it   = [cluster_resumo[c]["roi_intl_total"] for c in cats]

    for vals, col, fmt in [
        (roi_agr, 1, lambda v: f"{v:.2f}x"),
        (roi_dom, 2, lambda v: f"{v:.2f}x"),
        (roi_im,  3, lambda v: f"{v:.2f}"),
        (roi_it,  4, lambda v: f"{v:.1f}"),
    ]:
        fig.add_trace(go.Bar(
            x=cats, y=vals, marker_color=cores,
            text=[fmt(v) for v in vals], textposition="outside",
            showlegend=False,
        ), row=1, col=col)

    fig.update_layout(
        title="Retorno por Cluster de Produtora",
        height=460, margin=dict(l=40, r=20, t=80, b=140),
    )
    fig.update_xaxes(tickangle=-25)
    return fig

def make_cluster_box():
    """Box plot do ROI doméstico por cluster (produtoras com receita, log)."""
    fig = go.Figure()
    for c in CLUSTERS_ORD:
        g = df_prod[(df_prod["cluster"] == c) & (df_prod["roi_dom_total_deflac"] > 0)]
        if len(g) < 3:
            continue
        fig.add_trace(go.Box(
            y=g["roi_dom_total_deflac"],
            name=c,
            marker_color=CLUSTER_CORES[c],
            boxmean="sd",
        ))
    fig.update_layout(
        title="Distribuição do ROI Doméstico por Cluster<br><sup>(produtoras com receita · escala log)</sup>",
        yaxis=dict(title="ROI doméstico (receita/invest.)", type="log"),
        height=460, margin=dict(l=60, r=20, t=80, b=80),
        showlegend=False,
    )
    return fig

def make_cluster_box_intl():
    """Box plot do ROI internacional por cluster (produtoras com alcance intl, log)."""
    fig = go.Figure()
    for c in CLUSTERS_ORD:
        g = df_prod[(df_prod["cluster"] == c) & (df_prod["roi_intl_medio"] > 0)]
        if len(g) < 3:
            continue
        fig.add_trace(go.Box(
            y=g["roi_intl_medio"],
            name=c,
            marker_color=CLUSTER_CORES[c],
            boxmean="sd",
        ))
    fig.update_layout(
        title="Distribuição do ROI Internacional por Cluster<br><sup>(produtoras com alcance intl · escala log)</sup>",
        yaxis=dict(title="ROI intl médio das obras (0–100)", type="log"),
        height=460, margin=dict(l=60, r=20, t=80, b=80),
        showlegend=False,
    )
    return fig

def _cluster_marker_size(n_obras_series):
    """Tamanho robusto para marker size (aceita NaN)."""
    n = pd.to_numeric(n_obras_series, errors="coerce").fillna(1).clip(lower=1)
    return np.clip(np.sqrt(n) * 4, 5, 25)

def make_cluster_scatter():
    """Produtoras: investimento × ROI doméstico (log-log)."""
    fig = go.Figure()
    for c in CLUSTERS_ORD:
        g = df_prod[(df_prod["cluster"] == c) &
                    (df_prod["investimento_total_deflac"] > 0) &
                    (df_prod["roi_dom_total_deflac"] > 0)]
        if len(g) == 0:
            continue
        fig.add_trace(go.Scatter(
            x=g["investimento_total_deflac"] / 1e6,
            y=g["roi_dom_total_deflac"],
            mode="markers", name=c,
            marker=dict(
                color=CLUSTER_CORES[c],
                size=_cluster_marker_size(g["n_obras"]),
                opacity=0.65,
                line=dict(width=0.5, color="white"),
            ),
            text=g["razao_social"],
            hovertemplate="<b>%{text}</b><br>Inv: R$ %{x:.1f} mi<br>ROI dom: %{y:.2f}x<extra></extra>",
        ))
    fig.update_layout(
        title="Produtoras: Investimento × ROI Doméstico<br><sup>(tamanho = nº de obras · ambos os eixos em log)</sup>",
        xaxis=dict(title="Investimento total (R$ mi, log)", type="log"),
        yaxis=dict(title="ROI doméstico (log)", type="log"),
        height=560, margin=dict(l=70, r=20, t=80, b=80),
        legend=dict(orientation="h", y=-0.15),
    )
    return fig

def make_cluster_scatter_intl():
    """Produtoras: investimento × ROI internacional (log x, linear y)."""
    fig = go.Figure()
    for c in CLUSTERS_ORD:
        g = df_prod[(df_prod["cluster"] == c) &
                    (df_prod["investimento_total_deflac"] > 0) &
                    (df_prod["roi_intl_medio"] > 0)]
        if len(g) == 0:
            continue
        fig.add_trace(go.Scatter(
            x=g["investimento_total_deflac"] / 1e6,
            y=g["roi_intl_medio"],
            mode="markers", name=c,
            marker=dict(
                color=CLUSTER_CORES[c],
                size=_cluster_marker_size(g["n_obras"]),
                opacity=0.7,
                line=dict(width=0.5, color="white"),
            ),
            text=g["razao_social"],
            hovertemplate="<b>%{text}</b><br>Inv: R$ %{x:.1f} mi<br>ROI intl: %{y:.2f}<extra></extra>",
        ))
    fig.update_layout(
        title="Produtoras: Investimento × ROI Internacional<br><sup>(tamanho = nº de obras · eixo X em log)</sup>",
        xaxis=dict(title="Investimento total (R$ mi, log)", type="log"),
        yaxis=dict(title="ROI internacional médio das obras (0–100)"),
        height=560, margin=dict(l=70, r=20, t=80, b=80),
        legend=dict(orientation="h", y=-0.15),
    )
    return fig

def make_cluster_tabela():
    cats = CLUSTERS_ORD
    header = ["Cluster", "Nº Prod.", "Nº Obras", "Inv. Total",
              "Receita Tot.", "ROI Agr.", "Obras/Prod.",
              "ROI Dom. Med.", "ROI Intl Méd.", "ROI Intl TOTAL",
              "% c/ Intl", "Crítica Med."]
    cells = [
        cats,
        [fmt_n(cluster_resumo[c]["n_prod"]) for c in cats],
        [fmt_n(cluster_resumo[c]["n_obras"]) for c in cats],
        [fmt_bi(cluster_resumo[c]["inv_total_bi"]) for c in cats],
        [fmt_bi(cluster_resumo[c]["rec_total_bi"]) for c in cats],
        [fmt_x(cluster_resumo[c]["roi_agregado"]) for c in cats],
        [f"{cluster_resumo[c]['obras_por_prod']:.1f}" for c in cats],
        [fmt_x(cluster_resumo[c]["roi_dom_med"]) for c in cats],
        [f"{cluster_resumo[c]['roi_intl_medio']:.2f}" for c in cats],
        [f"{cluster_resumo[c]['roi_intl_total']:.1f}" for c in cats],
        [fmt_pct(cluster_resumo[c]["pct_com_intl"]) for c in cats],
        [f"{cluster_resumo[c]['critica_med']:.2f}" for c in cats],
    ]
    cores_row = [hex_to_rgba(CLUSTER_CORES[c], 0.18) for c in cats]
    fig = go.Figure(go.Table(
        header=dict(
            values=[f"<b>{h}</b>" for h in header],
            fill_color="#2c3e50", font=dict(color="white", size=11),
            align="center", height=34,
        ),
        cells=dict(
            values=cells,
            fill_color=[cores_row] * len(header),
            align=["left"] + ["right"] * (len(header) - 1),
            font=dict(size=11), height=30,
        ),
        columnwidth=[160, 60, 60, 85, 85, 60, 75, 80, 80, 90, 60, 70],
    ))
    fig.update_layout(
        title="Tabela detalhada por Cluster de Produtora",
        height=300, margin=dict(l=10, r=10, t=60, b=10),
    )
    return fig


# ── Ticket médio & Capital FSA parado ─────────────────────────────────────────
# Definição (tiers de severidade) de "sem retorno":
#   TIER1 (severíssimo):  receita total == 0              AND ROI intl == 0
#   TIER2 (estrito):      receita total < R$ 100 mil       AND ROI intl == 0
#   TIER3 (moderado):     bilheteria < R$ 100 mil          AND ROI intl == 0
LIMITE = 100_000
# Exclui obras de TV (categoria "_tv_excluir") de todo o escopo do capital parado
_nao_tv   = df["Categoria"] != "_tv_excluir"
_tem_fsa  = (df[COL_FSA] > 0) & _nao_tv
_intl_zero = df[COL_ROI_INT] == 0
_rec_zero  = df["receita_total"] == 0
_rec_baixa = df["receita_total"] < LIMITE
_bil_baixa = df[COL_BIL] < LIMITE

mask_tier1 = _tem_fsa & _rec_zero  & _intl_zero
mask_tier2 = _tem_fsa & _rec_baixa & _intl_zero
mask_tier3 = _tem_fsa & _bil_baixa & _intl_zero

fsa_universo = df.loc[_tem_fsa, COL_FSA].sum()

def _tier_stats(mask):
    return {
        "n_obras": int(mask.sum()),
        "fsa_bi":  df.loc[mask, COL_FSA].sum() / 1e9,
        "pct_fsa": (df.loc[mask, COL_FSA].sum() / fsa_universo * 100) if fsa_universo > 0 else 0,
    }
tiers = {
    "Tier 1 — Zero total":       _tier_stats(mask_tier1),
    "Tier 2 — Receita < R$100k": _tier_stats(mask_tier2),
    "Tier 3 — Bilh. < R$100k":   _tier_stats(mask_tier3),
}
# Para breakdowns por ano/categoria/grupo usamos o Tier 2 (definição-base)
df_parado = df[mask_tier2].copy()

# ── Ticket médio por obra (grupo & categoria) ────────────────────────────────
ticket_grupo = {g: df[df["grupo"] == g][COL_INV].mean() / 1e6 for g in GRUPOS}
ticket_grupo_mediano = {g: df[df["grupo"] == g][COL_INV].median() / 1e6 for g in GRUPOS}
ticket_cat = {c: df_cat[df_cat["Categoria"] == c][COL_INV].mean() / 1e6 for c in CATS_ORD}
ticket_cat_mediano = {c: df_cat[df_cat["Categoria"] == c][COL_INV].median() / 1e6 for c in CATS_ORD}

# Ticket médio por produtora, por cluster
ticket_cluster = {
    c: df_prod[df_prod["cluster"] == c]["investimento_total_deflac"].mean() / 1e6
    for c in CLUSTERS_ORD
}
ticket_cluster_mediano = {
    c: df_prod[df_prod["cluster"] == c]["investimento_total_deflac"].median() / 1e6
    for c in CLUSTERS_ORD
}

def make_ticket_obra_grupo():
    """Ticket (média + mediana) por obra — por Grupo de Mecanismo."""
    med = [ticket_grupo[g]         for g in GRUPOS]
    mdn = [ticket_grupo_mediano[g] for g in GRUPOS]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=GRUPOS, y=med, name="Média",
        marker_color=[CORES[g] for g in GRUPOS],
        text=[f"R$ {v:.2f}mi" for v in med], textposition="outside",
    ))
    fig.add_trace(go.Scatter(
        x=GRUPOS, y=mdn, name="Mediana",
        mode="markers",
        marker=dict(symbol="diamond", size=16, color="#222",
                    line=dict(width=1, color="white")),
    ))
    fig.update_layout(
        title="Ticket por Obra — por Grupo de Mecanismo<br><sup>(barras = média · diamantes = mediana · R$ mi 2024, log)</sup>",
        yaxis=dict(title="R$ mi por obra", type="log"),
        xaxis_tickangle=-25,
        height=440, margin=dict(l=60, r=20, t=80, b=120),
        legend=dict(orientation="h", y=1.1),
    )
    return fig

def make_ticket_obra_categoria():
    """Ticket (média + mediana) por obra — por Categoria de Chamada."""
    cats = CATS_ORD[::-1]
    med = [ticket_cat[c]         for c in cats]
    mdn = [ticket_cat_mediano[c] for c in cats]
    cores_c = [CAT_CORES[c] for c in cats]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=cats, x=med, orientation="h", name="Média",
        marker_color=cores_c,
        text=[f"R$ {v:.2f}mi" for v in med], textposition="outside",
    ))
    fig.add_trace(go.Scatter(
        y=cats, x=mdn, mode="markers", name="Mediana",
        marker=dict(symbol="diamond", size=12, color="#222",
                    line=dict(width=1, color="white")),
    ))
    fig.update_layout(
        title="Ticket por Obra — por Categoria de Chamada<br><sup>(barras = média · diamantes = mediana · R$ mi 2024, log)</sup>",
        xaxis=dict(title="R$ mi por obra", type="log"),
        height=max(500, 28 * len(cats) + 160),
        margin=dict(l=20, r=60, t=80, b=60),
        legend=dict(orientation="h", y=1.04),
    )
    fig.update_yaxes(automargin=True)
    return fig

def make_ticket_produtora():
    """Ticket médio POR PRODUTORA — por cluster (log)."""
    cats = CLUSTERS_ORD
    med = [ticket_cluster[c]         for c in cats]
    mdn = [ticket_cluster_mediano[c] for c in cats]
    cores = [CLUSTER_CORES[c] for c in cats]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=cats, y=med, name="Média",
        marker_color=cores,
        text=[f"R$ {v:.2f}mi" for v in med], textposition="outside",
    ))
    fig.add_trace(go.Scatter(
        x=cats, y=mdn, name="Mediana",
        mode="markers",
        marker=dict(symbol="diamond", size=16, color="#222",
                    line=dict(width=1, color="white")),
    ))
    fig.update_layout(
        title="Ticket Médio por Produtora<br><sup>(investimento total acumulado · R$ mi, log)</sup>",
        yaxis=dict(title="R$ mi", type="log"),
        height=460, margin=dict(l=60, r=20, t=80, b=120),
        xaxis_tickangle=-25,
        legend=dict(orientation="h", y=1.1),
    )
    return fig

def make_ticket_histograma():
    """Distribuição do ticket por obra (log) sobreposto por grupo."""
    fig = go.Figure()
    for g in GRUPOS:
        vals = df[(df["grupo"] == g) & (df[COL_INV] > 0)][COL_INV] / 1e6
        fig.add_trace(go.Histogram(
            x=vals, name=g, marker_color=CORES[g],
            opacity=0.55, nbinsx=50,
        ))
    fig.update_layout(
        title="Distribuição do ticket por obra (R$ mi, 2024)",
        xaxis=dict(title="Investimento por obra (R$ mi, log)", type="log"),
        yaxis=dict(title="Nº de obras"),
        barmode="overlay",
        height=420, margin=dict(l=60, r=20, t=60, b=60),
        legend=dict(orientation="h", y=-0.2),
    )
    return fig

def make_capital_parado_kpi():
    """Três KPIs — barras comparativas contra o universo FSA (tiers de severidade)."""
    nomes = list(tiers.keys())
    vals_bi  = [tiers[n]["fsa_bi"]  for n in nomes]
    vals_n   = [tiers[n]["n_obras"] for n in nomes]
    vals_pct = [tiers[n]["pct_fsa"] for n in nomes]
    cores_tier = ["#8B0000", "#D95F3A", "#E07B54"]

    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=[
            f"FSA sem retorno (R$ bi) — universo: R$ {fsa_universo/1e9:.2f} bi",
            "% do FSA total parado",
        ],
        horizontal_spacing=0.15,
    )

    # Barra de valor absoluto + linha do universo
    fig.add_trace(go.Bar(
        x=nomes, y=vals_bi, marker_color=cores_tier,
        text=[f"R$ {v:.2f} bi<br>({n:,} obras)".replace(",", ".")
              for v, n in zip(vals_bi, vals_n)],
        textposition="outside",
        showlegend=False,
    ), row=1, col=1)
    fig.add_hline(y=fsa_universo/1e9, line_dash="dash", line_color="#2c3e50",
                  annotation_text=f"Universo FSA: R$ {fsa_universo/1e9:.2f} bi",
                  annotation_position="top right", row=1, col=1)

    # Barra de percentual
    fig.add_trace(go.Bar(
        x=nomes, y=vals_pct, marker_color=cores_tier,
        text=[f"{v:.1f}%" for v in vals_pct],
        textposition="outside",
        showlegend=False,
    ), row=1, col=2)

    fig.update_yaxes(title_text="R$ bi", row=1, col=1,
                     range=[0, fsa_universo/1e9 * 1.15])
    fig.update_yaxes(title_text="% do FSA", row=1, col=2, range=[0, 100])

    fig.update_layout(
        title=f"Capital FSA investido sem retorno (3 tiers de severidade) — {int(_tem_fsa.sum()):,} obras FSA no universo".replace(",", "."),
        height=440, margin=dict(l=60, r=30, t=100, b=60),
    )
    return fig

def make_capital_parado_breakdown():
    """Breakdown do Tier 2 (definição-base) por ano, categoria e grupo."""
    fig = make_subplots(
        rows=1, cols=3,
        subplot_titles=[
            "FSA parado por ano",
            "FSA parado por categoria",
            "FSA parado por grupo",
        ],
        horizontal_spacing=0.12,
        column_widths=[0.33, 0.42, 0.25],
    )

    # Por ano
    ano_ser = df_parado.groupby("Ano")[COL_FSA].sum().sort_index() / 1e6
    fig.add_trace(go.Bar(
        x=ano_ser.index.astype(int), y=ano_ser.values,
        marker_color="#D95F3A",
        text=[f"{v:.0f}" for v in ano_ser.values], textposition="outside",
        showlegend=False,
    ), row=1, col=1)

    # Por categoria (ordenado desc)
    cat_ser = df_parado.groupby("Categoria")[COL_FSA].sum().sort_values(ascending=True) / 1e6
    cat_ser = cat_ser[cat_ser > 0]
    cores_cat = [CAT_CORES.get(c, "#888") for c in cat_ser.index]
    fig.add_trace(go.Bar(
        y=cat_ser.index, x=cat_ser.values, orientation="h",
        marker_color=cores_cat,
        text=[f"R$ {v:.0f}mi" for v in cat_ser.values], textposition="outside",
        showlegend=False,
    ), row=1, col=2)

    # Por grupo
    grp_ser = df_parado.groupby("grupo")[COL_FSA].sum() / 1e6
    grp_ser = grp_ser.reindex([g for g in GRUPOS if g in grp_ser.index])
    fig.add_trace(go.Bar(
        x=grp_ser.index, y=grp_ser.values,
        marker_color=[CORES[g] for g in grp_ser.index],
        text=[f"R$ {v:.0f}mi" for v in grp_ser.values], textposition="outside",
        showlegend=False,
    ), row=1, col=3)

    fig.update_layout(
        title=f"Capital parado (Tier 2: receita < R$100k E sem alcance intl) — R$ {tiers['Tier 2 — Receita < R$100k']['fsa_bi']:.2f} bi em {tiers['Tier 2 — Receita < R$100k']['n_obras']:,} obras".replace(",", "."),
        height=540, margin=dict(l=20, r=40, t=80, b=180),
    )
    fig.update_xaxes(tickangle=-30, row=1, col=3)
    fig.update_yaxes(automargin=True, row=1, col=2)
    return fig


# ── Internacionalização ────────────────────────────────────────────────────────
_FEST_EU = {
    "Cannes":        (43.5528,  7.0174, "Festival \u2014 Cannes"),
    "Berlim":        (52.5200, 13.4050, "Festival \u2014 Berlim"),
    "Veneza":        (45.4341, 12.3388, "Festival \u2014 Veneza"),
    "Rotterdam":     (51.9244,  4.4777, "Festival \u2014 Rotterdam"),
    "San Sebastián": (43.3183, -1.9812, "Festival \u2014 San Sebasti\u00e1n"),
    "Locarno":       (46.1706,  8.7981, "Festival \u2014 Locarno"),
    "Annecy":        (45.8992,  6.1294, "Festival \u2014 Annecy"),
    "BFI London":    (51.5074, -0.1278, "Festival \u2014 BFI London"),
}
_FEST_WORLD = {
    "Sundance":  (40.6461,-111.4980, "Festival \u2014 Sundance"),
    "TIFF":      (43.6532, -79.3832, "Festival \u2014 TIFF"),
    "Havana":    (23.1136, -82.3666, "Festival \u2014 Havana"),
    "NYFF":      (40.7812, -73.9800, "Festival \u2014 NYFF"),
    "Oscar":     (34.0522,-118.2437, "Festival \u2014 Oscar"),
}
_FEST_ALL = {**_FEST_EU, **_FEST_WORLD}


def _norm_title(s):
    if not isinstance(s, str): return ""
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_COUNTRY_TO_ISO2 = {
    "alemanha": "DE", "argentina": "AR", "australia": "AU", "austria": "AT",
    "belgica": "BE", "bosnia": "BA", "brasil": "BR", "bulgaria": "BG",
    "burkina faso": "BF", "canada": "CA", "chile": "CL", "china": "CN",
    "colombia": "CO", "coreia do sul": "KR", "croacia": "HR", "cuba": "CU",
    "dinamarca": "DK", "egito": "EG", "espanha": "ES", "estados unidos": "US",
    "eua": "US", "finlandia": "FI", "franca": "FR", "grecia": "GR",
    "holanda": "NL", "hong kong": "HK", "hungria": "HU", "india": "IN",
    "inglaterra": "GB", "irlanda": "IE", "italia": "IT", "japao": "JP",
    "marrocos": "MA", "mexico": "MX", "noruega": "NO", "nova zelandia": "NZ",
    "paises baixos": "NL", "polonia": "PL", "portugal": "PT",
    "reino unido": "GB", "republica checa": "CZ", "romenia": "RO",
    "servia": "RS", "suecia": "SE", "suica": "CH", "taiwan": "TW",
    "tchequia": "CZ", "turquia": "TR", "ucrania": "UA", "uruguai": "UY",
}

_FESTIVAL_ISO_HINTS = [
    ("academia brasileira", "BR"), ("grande premio do cinema brasileiro", "BR"),
    ("anima mundi", "BR"), ("cine ceara", "BR"), ("cine pe", "BR"),
    ("cine vitoria", "BR"), ("festival de brasilia", "BR"), ("festival do rio", "BR"),
    ("festival de gramado", "BR"), ("janela internacional", "BR"),
    ("mostra internacional de cinema de sao paulo", "BR"), ("tiradentes", "BR"),
    ("ouro preto", "BR"), ("curitiba", "BR"), ("guarnice", "BR"),
    ("semana de cinema", "BR"), ("e tudo verdade", "BR"), ("mix brasil", "BR"),
    ("forumdoc", "BR"), ("olhar de cinema", "BR"), ("panorama internacional", "BR"),
    ("amazon", "BR"), ("brasil africa", "BR"),
    ("cannes", "FR"), ("paris", "FR"), ("toulouse", "FR"), ("amiens", "FR"),
    ("annecy", "FR"), ("belfort", "FR"), ("biarritz", "FR"), ("marseille", "FR"),
    ("nantes", "FR"), ("cinema du reel", "FR"), ("film de femmes", "FR"),
    ("berlim", "DE"), ("berlin", "DE"), ("munique", "DE"), ("mannheim", "DE"),
    ("leipzig", "DE"), ("lucas internationales", "DE"),
    ("veneza", "IT"), ("roma", "IT"), ("turim", "IT"), ("giffoni", "IT"),
    ("giornate degli autori", "IT"),
    ("san sebastian", "ES"), ("donostia", "ES"), ("malaga", "ES"),
    ("huelva", "ES"), ("gijon", "ES"), ("valladolid", "ES"), ("sitges", "ES"),
    ("valencia", "ES"), ("cinema jove", "ES"),
    ("londres", "GB"), ("london", "GB"), ("bfi", "GB"), ("raindance", "GB"),
    ("sheffield", "GB"),
    ("locarno", "CH"), ("fribourg", "CH"), ("visions du reel", "CH"),
    ("nyon", "CH"), ("zurich", "CH"),
    ("roterda", "NL"), ("rotterdam", "NL"), ("amsterdam", "NL"),
    ("idfa", "NL"), ("cinekid", "NL"),
    ("lisboa", "PT"), ("doclisboa", "PT"), ("porto", "PT"), ("fantasporto", "PT"),
    ("indielisboa", "PT"), ("queer lisboa", "PT"),
    ("sundance", "US"), ("chicago", "US"), ("los angeles", "US"), ("nova york", "US"),
    ("new york", "US"), ("san francisco", "US"), ("miami", "US"), ("denver", "US"),
    ("tribeca", "US"), ("palm springs", "US"), ("seattle", "US"), ("sxsw", "US"),
    ("south by southwest", "US"), ("nashville", "US"), ("hamptons", "US"),
    ("cine las americas", "US"), ("newfest", "US"), ("art of the real", "US"),
    ("toronto", "CA"), ("montreal", "CA"), ("vancouver", "CA"), ("hot docs", "CA"),
    ("ottawa", "CA"), ("nouveau cinema", "CA"),
    ("havana", "CU"), ("mar del plata", "AR"), ("bafici", "AR"), ("cine politico", "AR"),
    ("bogota", "CO"), ("cartagena", "CO"), ("guadalajara", "MX"), ("morelia", "MX"),
    ("docs mx", "MX"), ("ficunam", "MX"), ("femcine", "CL"), ("valdivia", "CL"),
    ("vina del mar", "CL"), ("punta del este", "UY"), ("cinema do uruguai", "UY"),
    ("assuncao", "PY"), ("fespaco", "BF"),
    ("marrakesh", "MA"), ("cairo", "EG"), ("durban", "ZA"),
    ("busan", "KR"), ("jeonju", "KR"), ("toquio", "JP"), ("tokyo", "JP"),
    ("hong kong", "HK"), ("pingyao", "CN"), ("beijing", "CN"), ("shangai", "CN"),
    ("taipei", "TW"), ("golden horse", "TW"), ("goa", "IN"), ("kerala", "IN"),
    ("kolkata", "IN"), ("mumbai", "IN"),
    ("melbourne", "AU"), ("sydney", "AU"), ("athens", "GR"), ("thessaloniki", "GR"),
    ("sarajevo", "BA"), ("sofia", "BG"), ("transilvania", "RO"), ("varsovia", "PL"),
    ("cracovia", "PL"), ("moscou", "RU"), ("kiev", "UA"), ("molodist", "UA"),
    ("cork", "IE"), ("galway", "IE"), ("estocolmo", "SE"), ("tampere", "FI"),
    ("istambul", "TR"), ("bosforo", "TR"), ("black nights", "EE"), ("tallinn", "EE"),
    ("karlovy vary", "CZ"),
]


def _festival_country_lookup():
    lookup = {}
    if BRDE_FEST_LIST.exists():
        try:
            ref = pd.read_csv(BRDE_FEST_LIST, dtype=str).fillna("")
            for _, row in ref.iterrows():
                pais = str(row.get("PAIS", "")).split("/")[0].strip()
                iso2 = _COUNTRY_TO_ISO2.get(_norm_title(pais))
                if not iso2:
                    continue
                for col in ["FESTIVAL", "LINHA_ORIGINAL"]:
                    name = _norm_title(row.get(col, ""))
                    if name:
                        lookup[name] = iso2
        except Exception:
            pass
    return lookup


def _festival_iso2(name, lookup):
    key = _norm_title(name)
    if not key:
        return ""
    if key in lookup:
        return lookup[key]
    for ref, iso2 in lookup.items():
        if len(ref) >= 8 and (ref in key or key in ref):
            return iso2
    for token, iso2 in _FESTIVAL_ISO_HINTS:
        if token in key:
            return iso2
    return ""


def make_intl_festival_map():
    """Retorna (html_str, js_str) — HTML sem <script> + JS separado."""

    # ISO2 → ISO numeric (world-atlas)
    ISO2_NUM = {
        "AT":40,"BE":56,"BG":100,"CH":756,"CZ":203,"DE":276,"DK":208,
        "EE":233,"ES":724,"FI":246,"FR":250,"GB":826,"GR":300,"HR":191,
        "HU":348,"IE":372,"IT":380,"LT":440,"LU":442,"LV":428,"MK":807,
        "MT":470,"NL":528,"NO":578,"PL":616,"PT":620,"RO":642,"RS":688,
        "SE":752,"SI":705,"SK":703,"TR":792,"UA":804,"AL":8,"BA":70,
        "IS":352,"ME":499,"MD":498,"BY":112,
    }
    ISO2_NUM.update({
        "AR":32,"AU":36,"BF":854,"BR":76,"CA":124,"CL":152,"CN":156,
        "CO":170,"CU":192,"EG":818,"HK":344,"IN":356,"JP":392,"KR":410,
        "MA":504,"MX":484,"NZ":554,"PY":600,"RU":643,"TW":158,"US":840,
        "UY":858,"ZA":710,
    })
    COUNTRY_NAMES = {
        "FR":"França","DE":"Alemanha","GB":"Reino Unido","AT":"Áustria",
        "IE":"Irlanda","ES":"Espanha","PT":"Portugal","BE":"Bélgica",
        "CH":"Suíça","SE":"Suécia","NL":"Países Baixos","NO":"Noruega",
        "IT":"Itália","FI":"Finlândia","PL":"Polônia","DK":"Dinamarca",
        "SK":"Eslováquia","GR":"Grécia","HU":"Hungria","CZ":"Tchéquia",
        "BG":"Bulgária","SI":"Eslovênia","RO":"Romênia","HR":"Croácia",
        "LV":"Letônia","LT":"Lituânia","EE":"Estônia","RS":"Sérvia",
        "MK":"Macedônia do Norte","MT":"Malta","TR":"Turquia","LU":"Luxemburgo",
        "AL":"Albânia","BA":"Bósnia","IS":"Islândia","ME":"Montenegro",
        "MD":"Moldávia","BY":"Bielorrússia","UA":"Ucrânia",
    }

    # ── Festivais por país — obras do df ativo ─────────────────────────────────
    COUNTRY_NAMES.update({
        "AR":"Argentina","AU":"Austrália","BF":"Burkina Faso","BR":"Brasil",
        "CA":"Canadá","CL":"Chile","CN":"China","CO":"Colômbia","CU":"Cuba",
        "EG":"Egito","HK":"Hong Kong","IN":"Índia","JP":"Japão",
        "KR":"Coreia do Sul","MA":"Marrocos","MX":"México",
        "NZ":"Nova Zelândia","PY":"Paraguai","RU":"Rússia","TW":"Taiwan",
        "US":"Estados Unidos","UY":"Uruguai","ZA":"África do Sul",
    })

    fest_by_country = {}   # iso2 -> title -> set(festivals)
    festival_lookup = _festival_country_lookup()
    if BASE_FEST_ATA.exists():
        df_fest_ata = pd.read_csv(BASE_FEST_ATA, dtype=str).fillna("")
        cpbs_ativos = set(df["CPB"].astype(str)) if "CPB" in df.columns else set()
        for _, row in df_fest_ata.iterrows():
            cpb = str(row.get("cpb", "")).strip()
            if cpbs_ativos and cpb not in cpbs_ativos:
                continue
            festival = str(row.get("festival", "")).strip()
            iso2 = _festival_iso2(festival, festival_lookup)
            if iso2 not in ISO2_NUM or iso2 == "BR":
                continue
            titulo = str(row.get("titulo", "") or cpb).strip()
            fest_by_country.setdefault(iso2, {}).setdefault(titulo, set()).add(festival)
    elif BASE_OBRA.exists():
        df_obra = pd.read_csv(BASE_OBRA, sep=";", encoding="utf-8-sig",
                              usecols=["CPB","titulo","paises_festivais"])
        cpbs_ativos = set(df["CPB"].astype(str)) if "CPB" in df.columns else set()
        for _, row in df_obra.iterrows():
            if cpbs_ativos and str(row["CPB"]) not in cpbs_ativos:
                continue
            val = str(row.get("paises_festivais", "") or "").strip()
            if val and val != "nan":
                titulo = str(row.get("titulo", row["CPB"]))
                for iso2 in val.split("|"):
                    iso2 = iso2.strip()
                    if iso2 in ISO2_NUM:
                        fest_by_country.setdefault(iso2, {}).setdefault(titulo, set()).add("Base por obra")

    # ── VOD por país — lumiere_vod_search (todas obras BR) ────────────────────
    vod_by_country = {}   # iso2 → set of titles
    if VOD_RAW.exists():
        df_vod = pd.read_excel(VOD_RAW, usecols=["Original title","Country","Producing country"])
        df_vod = df_vod[df_vod["Producing country"].str.contains("BR", na=False)]
        for _, row in df_vod.iterrows():
            iso2 = str(row["Country"]).strip()
            if iso2 in ISO2_NUM:
                vod_by_country.setdefault(iso2, set()).add(str(row["Original title"]))

    # ── Lumière teatral: total EU ──────────────────────────────────────────────
    n_lum_total, adm_lum_total = 0, 0
    if LUM_RAW.exists():
        lum_df = pd.read_excel(LUM_RAW)
        n_lum_total   = len(lum_df)
        adm_col = [c for c in lum_df.columns if "Total" in c or "since" in c]
        if adm_col:
            adm_lum_total = int(lum_df[adm_col[0]].sum())

    # ── Converter sets → sorted lists para JSON ────────────────────────────────
    fest_data = {
        iso2: [
            f"{titulo} — {', '.join(sorted(festivals))}"
            for titulo, festivals in sorted(title_map.items())
        ]
        for iso2, title_map in fest_by_country.items()
    }
    vod_data  = {k: sorted(v) for k, v in vod_by_country.items()}

    # ── Dados por país para mapa ────────────────────────────────────────────────
    # map_data: numeric_id → {iso2, fest:[...], vod:[...]}
    all_iso2 = sorted(set(fest_data) | set(vod_data))
    map_data = {}
    for iso2 in all_iso2:
        num = ISO2_NUM.get(iso2)
        if num:
            map_data[num] = {
                "iso2": iso2,
                "fest": fest_data.get(iso2, []),
                "vod":  vod_data.get(iso2, []),
            }

    # Top países bar chart data (VOD, que tem mais dados)
    top_vod = sorted(all_iso2, key=lambda iso: len(fest_data.get(iso, [])) + len(vod_data.get(iso, [])), reverse=True)[:15]
    max_vod = max([len(fest_data.get(iso, [])) + len(vod_data.get(iso, [])) for iso in top_vod] or [1])
    bar_parts = []
    for iso2 in top_vod:
        films = vod_data.get(iso2, [])
        n_fest = len(fest_data.get(iso2, []))
        total_country = len(films) + n_fest
        name = COUNTRY_NAMES.get(iso2, iso2)
        w = round(total_country / max_vod * 100)
        fest_badge = f' <span style="font-size:8px;color:#E8702A">+{n_fest} fest.</span>' if n_fest else ""
        bar_parts.append(
            f'<div style="display:flex;align-items:center;gap:6px;margin-bottom:4px">'
            f'<div style="width:28px;flex-shrink:0;font-size:8px;color:#666;text-align:right">{iso2}</div>'
            f'<div style="flex:1;height:16px;background:#eef0f4;border-radius:2px;overflow:hidden;position:relative">'
            f'<div style="width:{w}%;height:100%;background:#5B6BB5;opacity:.75;border-radius:2px"></div>'
            f'<div style="position:absolute;left:5px;top:0;height:100%;display:flex;align-items:center;font-size:9px;color:#222">'
            f'{name}{fest_badge}</div></div>'
            f'<div style="width:22px;flex-shrink:0;font-size:9px;font-weight:700;color:#5B6BB5;text-align:right">{total_country}</div>'
            f'</div>'
        )
    bar_html = "".join(bar_parts)

    map_data_js    = json.dumps(map_data,    ensure_ascii=False)
    cnames_js      = json.dumps(COUNTRY_NAMES, ensure_ascii=False)
    bar_html_js    = json.dumps(bar_html,    ensure_ascii=False)
    n_fest_total   = sum(len(v) for v in fest_data.values())
    n_vod_total    = len(set(t for v in vod_data.values() for t in v))
    n_fest_paises  = len(fest_data)
    n_vod_paises   = len(vod_data)

    adm_str = f"{adm_lum_total/1e6:.1f}M" if adm_lum_total > 0 else "—"

    html_body = f"""
<div style="font-family:'DM Sans',sans-serif;color:#222;background:#fff;padding:4px 0">

  <!-- KPIs -->
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px">
    <div style="background:#f7f8fb;border:1px solid #dde0e8;border-top:3px solid #E8702A;border-radius:4px;padding:12px 14px">
      <div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:#888;margin-bottom:4px">Festivais internacionais</div>
      <div style="font-size:28px;font-weight:800;line-height:1;color:#E8702A;margin-bottom:2px">{n_fest_paises} <span style="font-size:12px;color:#888">países</span></div>
      <div style="font-size:10px;color:#666">{n_fest_total} participações · fonte: ANCINE/base</div>
    </div>
    <div style="background:#f7f8fb;border:1px solid #dde0e8;border-top:3px solid #5B6BB5;border-radius:4px;padding:12px 14px">
      <div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:#888;margin-bottom:4px">VOD Europa</div>
      <div style="font-size:28px;font-weight:800;line-height:1;color:#5B6BB5;margin-bottom:2px">{n_vod_paises} <span style="font-size:12px;color:#888">países</span></div>
      <div style="font-size:10px;color:#666">{n_vod_total} títulos · fonte: Lumière VOD</div>
    </div>
    <div style="background:#f7f8fb;border:1px solid #dde0e8;border-top:3px solid #B8860B;border-radius:4px;padding:12px 14px">
      <div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:#888;margin-bottom:4px">Salas de cinema (Europa)</div>
      <div style="font-size:28px;font-weight:800;line-height:1;color:#B8860B;margin-bottom:2px">{n_lum_total} <span style="font-size:12px;color:#888">obras</span></div>
      <div style="font-size:10px;color:#666">{adm_str} admissões · fonte: Lumière/CNC</div>
    </div>
  </div>

  <!-- Layout: mapa + ranking -->
  <div style="display:grid;grid-template-columns:1fr 260px;gap:14px;align-items:start">

    <!-- Mapa -->
    <div>
      <!-- Legenda + filtros -->
      <div style="display:flex;align-items:center;gap:14px;margin-bottom:8px;flex-wrap:wrap">
        <span style="font-size:11px;font-weight:600;color:#444">Filtrar:</span>
        <label style="display:flex;align-items:center;gap:5px;font-size:10px;cursor:pointer">
          <input type="checkbox" id="chk-fest" checked onchange="intlMapUpdate()">
          <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:#E8702A;vertical-align:middle"></span>
          Festivais
        </label>
        <label style="display:flex;align-items:center;gap:5px;font-size:10px;cursor:pointer">
          <input type="checkbox" id="chk-vod" checked onchange="intlMapUpdate()">
          <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:#5B6BB5;vertical-align:middle"></span>
          VOD Europa
        </label>
        <span style="font-size:9px;color:#888;margin-left:auto">passe o mouse para ver os filmes · scroll para zoom · arraste para mover</span>
      </div>
      <div style="background:#f0f4f8;border:1px solid #ccd0da;border-radius:4px;overflow:hidden;position:relative">
        <svg id="intl-map-svg" viewBox="0 0 960 500" preserveAspectRatio="xMidYMid meet"
             style="width:100%;display:block;cursor:grab"></svg>
        <!-- Lumière badge -->
        <div style="position:absolute;bottom:8px;left:8px;background:rgba(255,255,255,.9);border:1px solid #ddd;border-radius:3px;padding:5px 8px;font-size:9px;color:#555">
          <span style="color:#B8860B;font-weight:700">★ Lumière/CNC</span> · {n_lum_total} obras em salas europeias · {adm_str} admissões
        </div>
      </div>
    </div>

    <!-- Ranking VOD -->
    <div>
      <div style="font-size:10px;font-weight:600;color:#444;margin-bottom:8px;letter-spacing:.05em;text-transform:uppercase">Top países — nº de títulos em VOD</div>
      <div id="intl-country-bars">{bar_html}</div>
    </div>
  </div>

  <!-- Tooltip -->
  <div id="intl-tooltip" style="position:fixed;background:#fff;border:1px solid #5B6BB5;border-radius:5px;padding:10px 13px;font-size:10px;pointer-events:none;display:none;z-index:9999;max-width:520px;max-height:70vh;overflow:auto;box-shadow:0 4px 20px rgba(0,0,0,.18)">
    <div id="intl-tt-title" style="font-size:11px;font-weight:700;color:#333;margin-bottom:6px"></div>
    <div id="intl-tt-body"></div>
  </div>
  <div id="intl-country-detail" style="margin-top:14px;background:#f7f8fb;border:1px solid #dde0e8;border-radius:5px;padding:12px 14px;font-size:10px;color:#333">
    Clique em um país no mapa para ver todas as obras identificadas em festivais e VOD.
  </div>
</div>
"""

    js_body = f"""
(function() {{
  var MAP_DATA = {map_data_js};
  var CNAMES   = {cnames_js};

  var gG = null, gPath = null, gWorld = null;

  function loadD3AndDraw() {{
    // Se D3 já está carregado (ex: mega painel), vai direto
    if (typeof d3 !== 'undefined' && typeof topojson !== 'undefined') {{
      fetchWorld();
      return;
    }}
    // Senão carrega dinamicamente
    function loadScript(src, cb) {{
      var s = document.createElement('script');
      s.src = src; s.onload = cb; document.head.appendChild(s);
    }}
    loadScript('https://d3js.org/d3.v7.min.js', function() {{
      loadScript('https://cdn.jsdelivr.net/npm/topojson@3/dist/topojson.min.js', fetchWorld);
    }});
  }}

  function fetchWorld() {{
    fetch('https://cdn.jsdelivr.net/npm/world-atlas@2/countries-110m.json')
      .then(function(r) {{ return r.json(); }})
      .then(function(world) {{ initMap(world); }});
  }}

  function initMap(world) {{
    gWorld = world;
    var svgEl = document.getElementById('intl-map-svg');
    if (!svgEl) return;
    var svg = d3.select(svgEl);
    var W = 960, H = 500;
    var proj = d3.geoNaturalEarth1().scale(165).translate([W/2, H/2 + 20]);
    gPath = d3.geoPath().projection(proj);
    gG    = svg.append('g');

    var zoom = d3.zoom().scaleExtent([1, 14]).on('zoom', function(ev) {{
      gG.attr('transform', ev.transform);
      gG.selectAll('path').attr('stroke-width', 0.5 / ev.transform.k);
    }});

    svg.call(zoom)
       .on('mousedown.zoom', function() {{ svgEl.style.cursor = 'grabbing'; }})
       .on('mouseup.zoom',   function() {{ svgEl.style.cursor = 'grab'; }});

    drawMap();
  }}

  function inEurope(f) {{
    try {{
      const [[x0,y0],[x1,y1]] = d3.geoBounds(f);
      return x1 > -30 && x0 < 50 && y1 > 32 && y0 < 73;
    }} catch(e) {{ return false; }}
  }}

  function countryColor(d) {{
    const id = +d.id;
    const info = MAP_DATA[id];
    if (!info) return '#dde0e8';
    const showFest = document.getElementById('chk-fest').checked;
    const showVod  = document.getElementById('chk-vod').checked;
    const hasFest  = showFest && info.fest && info.fest.length > 0;
    const hasVod   = showVod  && info.vod  && info.vod.length  > 0;
    if (hasFest && hasVod) return '#7B5EA7';   // ambos — roxo
    if (hasFest)           return '#E8702A';   // só festivais — laranja
    if (hasVod)            return '#5B6BB5';   // só VOD — azul
    return '#dde0e8';                          // sem dados visíveis
  }}

  function esc(s) {{
    return String(s || '').replace(/[&<>"']/g, function(ch) {{
      return ({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}})[ch];
    }});
  }}

  function renderCountryDetail(info) {{
    const box = document.getElementById('intl-country-detail');
    if (!box || !info) return;
    const name = CNAMES[info.iso2] || info.iso2;
    const fest = info.fest || [];
    const vod = info.vod || [];
    let html = '<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:8px">';
    html += '<div><div style="font-size:12px;font-weight:800;color:#222">' + esc(name) + ' (' + esc(info.iso2) + ')</div>';
    html += '<div style="font-size:9px;color:#666;margin-top:2px">Fontes: festivais na base ATA BRDE/FSA 2024; VOD na base Lumière VOD.</div></div>';
    html += '<div style="font-size:10px;color:#555;white-space:nowrap"><b style="color:#E8702A">' + fest.length + '</b> obras em festival · <b style="color:#5B6BB5">' + vod.length + '</b> títulos VOD</div></div>';
    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;align-items:start">';
    html += '<div><div style="font-size:9px;font-weight:700;color:#E8702A;margin-bottom:5px;text-transform:uppercase;letter-spacing:.06em">Festivais</div>';
    html += fest.length ? fest.map(function(f) {{ return '<div style="padding:4px 0;border-bottom:1px solid #e3e5eb;line-height:1.25">' + esc(f) + '</div>'; }}).join('') : '<div style="color:#888">Sem obra identificada em festival neste país.</div>';
    html += '</div><div><div style="font-size:9px;font-weight:700;color:#5B6BB5;margin-bottom:5px;text-transform:uppercase;letter-spacing:.06em">VOD Europa</div>';
    html += vod.length ? vod.map(function(f) {{ return '<div style="padding:4px 0;border-bottom:1px solid #e3e5eb;line-height:1.25">' + esc(f) + '</div>'; }}).join('') : '<div style="color:#888">Sem título identificado em VOD neste país.</div>';
    html += '</div></div>';
    box.innerHTML = html;
  }}

  function drawMap() {{
    if (!gWorld) return;
    gG.selectAll('*').remove();
    const countries = topojson.feature(gWorld, gWorld.objects.countries);
    const tt  = document.getElementById('intl-tooltip');
    const ttT = document.getElementById('intl-tt-title');
    const ttB = document.getElementById('intl-tt-body');

    gG.selectAll('path.country')
      .data(countries.features)
      .enter().append('path')
      .attr('class', 'country')
      .attr('fill', d => countryColor(d))
      .attr('stroke', '#fff')
      .attr('stroke-width', .5)
      .style('cursor', d => MAP_DATA[+d.id] ? 'pointer' : 'default')
      .style('transition', 'fill .12s')
      .on('mousemove', function(event, d) {{
        const id = +d.id;
        const info = MAP_DATA[id];
        if (!info) return;
        const showFest = document.getElementById('chk-fest').checked;
        const showVod  = document.getElementById('chk-vod').checked;
        const fest = showFest ? (info.fest || []) : [];
        const vod  = showVod  ? (info.vod  || []) : [];
        if (!fest.length && !vod.length) return;
        const name = CNAMES[info.iso2] || info.iso2;
        ttT.textContent = name;
        let html = '';
        if (fest.length) {{
          html += '<div style="font-size:9px;font-weight:600;color:#E8702A;margin-bottom:3px">🎬 Festivais (' + fest.length + ' obras)</div>';
          html += fest.map(f => '<div style="font-size:9px;padding:1px 0;border-bottom:1px solid #eee">' + esc(f) + '</div>').join('');
        }}
        if (vod.length) {{
          html += '<div style="font-size:9px;font-weight:600;color:#5B6BB5;margin:' + (fest.length?'8px':'0') + ' 0 3px">📺 VOD (' + vod.length + ' títulos)</div>';
          html += vod.map(f => '<div style="font-size:9px;padding:1px 0;border-bottom:1px solid #eee">' + esc(f) + '</div>').join('');
        }}
        ttB.innerHTML = html;
        tt.style.display = 'block';
        tt.style.left = (event.clientX + 14) + 'px';
        tt.style.top  = Math.max(10, event.clientY - 20) + 'px';
      }})
      .on('click', function(event, d) {{
        const info = MAP_DATA[+d.id];
        if (info) renderCountryDetail(info);
      }})
      .on('mouseleave', function() {{ tt.style.display = 'none'; }});

    // Bordas entre países
    gG.append('path')
      .datum(topojson.mesh(gWorld, gWorld.objects.countries, (a,b) => a !== b))
      .attr('fill', 'none')
      .attr('stroke', '#fff')
      .attr('stroke-width', .5)
      .attr('d', gPath);
  }}

  window.intlMapUpdate = function() {{ if (gWorld) drawMap(); }};
  // Garante renderização mesmo quando o painel estava oculto no carregamento
  window.intlMapEnsure = function() {{
    if (gWorld) {{ drawMap(); }} else {{ loadD3AndDraw(); }}
  }};

  // Inicia
  if (document.readyState === 'loading') {{
    document.addEventListener('DOMContentLoaded', loadD3AndDraw);
  }} else {{
    loadD3AndDraw();
  }}
}})();
"""
    return html_body, js_body


def make_intl_lumiere():
    """Top 20 obras por admissões europeias (Lumière)."""
    COL_LUM = "Adm. EU \u2014 Lumi\u00e8re"
    if COL_LUM not in df.columns:
        return go.Figure()
    sub = df[df[COL_LUM] > 0][["Projeto", "grupo", COL_LUM]].copy()
    sub = sub.nlargest(20, COL_LUM).sort_values(COL_LUM)
    fig = go.Figure(go.Bar(
        y=sub["Projeto"],
        x=sub[COL_LUM] / 1e3,
        orientation="h",
        marker_color=[CORES.get(g, "#888") for g in sub["grupo"]],
        text=[f"{v/1e3:.0f}k" for v in sub[COL_LUM]],
        textposition="outside",
    ))
    fig.update_layout(
        title="Top 20 — Admissões na Europa (Lumière, mil ingressos)",
        xaxis_title="Admissões (mil)",
        height=520, margin=dict(l=280, r=60, t=50, b=40),
        showlegend=False,
    )
    return fig


# ── Renderização HTML ──────────────────────────────────────────────────────────
print("Gerando figuras...")

fig_tabela    = make_tabela_resumo()
fig_inv_bil   = make_inv_bil()
fig_comp_inv  = make_composicao_inv()
fig_roi_agr   = make_roi_agregado()
fig_roi_dom   = make_roi_dom_box()
fig_roi_int   = make_roi_int_box()
fig_scatter_inv  = make_scatter_inv_roi()
fig_paises       = make_paises()
html_intl_map, js_intl_map = make_intl_festival_map()
fig_intl_lumiere = make_intl_lumiere()
fig_scatter   = make_scatter_roi()
fig_inv_med   = make_inv_medio()
fig_cat_vol      = make_cat_volumes()
fig_cat_roi      = make_cat_roi()
fig_cat_tab      = make_cat_tabela()
fig_cat_box      = make_cat_roi_box()
fig_cat_box_intl = make_cat_roi_int_box()
fig_cat_scatter  = make_scatter_cat()
fig_cl_kpi       = make_cluster_kpi()
fig_cl_roi       = make_cluster_roi()
fig_cl_box       = make_cluster_box()
fig_cl_box_intl  = make_cluster_box_intl()
fig_cl_scat      = make_cluster_scatter()
fig_cl_scat_intl = make_cluster_scatter_intl()
fig_cl_tab       = make_cluster_tabela()
fig_tck_obra_g   = make_ticket_obra_grupo()
fig_tck_obra_c   = make_ticket_obra_categoria()
fig_tck_prod     = make_ticket_produtora()
fig_tck_hist     = make_ticket_histograma()
fig_cp_kpi       = make_capital_parado_kpi()
fig_cp_brk       = make_capital_parado_breakdown()

# Tab 1: renderiza inline (sempre visível na carga)
def to_div(fig, div_id):
    return fig.to_html(full_html=False, include_plotlyjs=False, div_id=div_id)

# Tabs 2-5: lazy — placeholder div + JSON embutido, init no primeiro click
def to_lazy(fig, div_id, height=450, search=None):
    """
    search=None   → sem busca
    search="text" → campo de busca textual (filtra pontos por string em text/customdata)
    search="rows" → campo de busca que filtra linhas de uma go.Table pela 1ª coluna
    """
    fig_json = fig.to_json()
    search_html = ""
    data_attr = ""
    if search:
        data_attr = f' data-search-mode="{search}"'
        placeholder = ("Buscar por projeto/produtora…" if search == "text"
                       else "Filtrar linhas da tabela…")
        search_html = (
            f'<div class="search-wrap">'
            f'<input type="text" class="search-input" placeholder="{placeholder}" '
            f'oninput="filterFig(\'{div_id}\', this.value)">'
            f'<span class="search-count" id="count_{div_id}"></span>'
            f'</div>'
        )
    return (
        f'{search_html}'
        f'<div id="{div_id}"{data_attr} style="width:100%;height:{height}px;"></div>\n'
        f'<script>window.__fig_{div_id} = {fig_json};</script>'
    )


def to_toggle(fig, div_id, height, labels, n_per_mode):
    """Lazy figure com botões HTML para alternar entre modos de visualização."""
    import json as _json
    fig_json = fig.to_json()
    y_titles = [
        "ROI Doméstico (receita/invest.)",
        "ROI Internacional (0–100)",
    ]
    meta = _json.dumps({"n": n_per_mode, "yTitles": y_titles})
    btns = "".join(
        f'<button class="tog-btn{"  active" if i == 0 else ""}" '
        f'onclick="togSwitch(\'{div_id}\',{i},this)">{lbl}</button>'
        for i, lbl in enumerate(labels)
    )
    return (
        f'<div class="toggle-bar">{btns}</div>'
        f'<div id="{div_id}" style="width:100%;height:{height}px;"></div>\n'
        f'<script>window.__fig_{div_id}={fig_json};'
        f'window.__togMeta_{div_id}={meta};</script>'
    )


# KPIs estáticos (sem Plotly)
kpi_html = "".join([
    f'''<div class="kpi-card" style="--cor:{CORES[g]}">
      <div class="label">{g.replace(" — ", "<br>")}</div>
      <div class="val">{fmt_n(resumo[g]["n_obras"])}</div>
      <div class="sub">obras · {fmt_bi(resumo[g]["inv_total_bi"])} invest.</div>
    </div>'''
    for g in GRUPOS
])

HTML = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>FASE 3 — Comparativo de Mecanismos de Fomento</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', sans-serif; background: #f0f2f5; color: #222; }}

  .topbar {{
    background: #2c3e50; color: white; padding: 18px 32px;
    display: flex; align-items: baseline; gap: 16px;
  }}
  .topbar h1 {{ font-size: 1.4rem; font-weight: 700; }}
  .topbar span {{ font-size: 0.85rem; opacity: 0.7; }}

  .tabs {{
    background: #34495e; display: flex; gap: 0; padding: 0 24px;
    border-bottom: 3px solid #1abc9c;
  }}
  .tab-btn {{
    padding: 12px 22px; cursor: pointer; color: #ccc; font-size: 0.9rem;
    border: none; background: none; transition: all .2s;
  }}
  .tab-btn:hover {{ color: white; background: rgba(255,255,255,0.08); }}
  .tab-btn.active {{ color: white; border-bottom: 3px solid #1abc9c; margin-bottom: -3px; }}

  .tab-content {{ display: none; padding: 28px 32px; }}
  .tab-content.active {{ display: block; }}

  .section-title {{
    font-size: 1rem; font-weight: 600; color: #2c3e50;
    margin-bottom: 12px; padding-bottom: 6px;
    border-bottom: 2px solid #1abc9c;
  }}
  .card {{
    background: white; border-radius: 10px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    padding: 20px; margin-bottom: 20px;
  }}
  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
  .kpi-bar {{
    display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px;
    margin-bottom: 20px;
  }}
  .kpi-card {{
    background: white; border-radius: 10px; padding: 16px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    border-left: 4px solid var(--cor);
  }}
  .kpi-card .label {{ font-size: 0.72rem; color: #666; text-transform: uppercase; letter-spacing: .5px; }}
  .kpi-card .val {{ font-size: 1.5rem; font-weight: 700; color: #2c3e50; }}
  .kpi-card .sub {{ font-size: 0.78rem; color: #888; }}
  .legenda {{
    display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 20px;
    background: white; padding: 14px 20px; border-radius: 8px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  }}
  .legenda-item {{ display: flex; align-items: center; gap: 8px; font-size: 0.85rem; }}
  .legenda-dot {{ width: 14px; height: 14px; border-radius: 50%; flex-shrink: 0; }}
  .toggle-bar {{ display: flex; gap: 6px; margin-bottom: 10px; }}
  .tog-btn {{
    padding: 5px 14px; font-size: 0.8rem; cursor: pointer; border-radius: 4px;
    border: 1px solid #aaa; background: #f0f2f5; color: #555; transition: all .15s;
  }}
  .tog-btn.active {{ background: #2c3e50; color: #fff; border-color: #2c3e50; }}
  .tog-btn:hover:not(.active) {{ background: #e0e4ea; }}
  .search-wrap {{
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 10px;
  }}
  .search-input {{
    flex: 1; max-width: 420px;
    padding: 8px 12px; border: 1px solid #cfd6df; border-radius: 6px;
    font-size: 0.88rem; font-family: inherit;
    transition: border-color .15s;
  }}
  .search-input:focus {{ outline: none; border-color: #1abc9c; box-shadow: 0 0 0 2px rgba(26,188,156,0.18); }}
  .search-count {{ font-size: 0.78rem; color: #666; }}
  @media (max-width: 900px) {{
    .grid-2, .kpi-bar {{ grid-template-columns: 1fr; }}
  }}
</style>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<script src="https://d3js.org/d3.v7.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/topojson@3/dist/topojson.min.js"></script>
</head>
<body>

<div class="topbar">
  <h1>FASE 3 — Comparativo de Mecanismos de Fomento</h1>
  <span>Cinema Brasileiro · Recorte {ANO_INI}–{ANO_FIM} (10 anos) · Escopo PRODECINE (exclui SAV/MINC, Arranjos e TV/PRODAV) · Dados ANCINE · Valores em R$2024 (IPCA deflac.)</span>
</div>

<div class="tabs">
  <button class="tab-btn active" onclick="showTab(event,'visao-geral')">Visão Geral</button>
  <button class="tab-btn" onclick="showTab(event,'financeiro')">Retorno Doméstico</button>
  <button class="tab-btn" onclick="showTab(event,'ret-intl')">Retorno Internacional</button>
  <button class="tab-btn" onclick="showTab(event,'roi-dom')">ROI Doméstico</button>
  <button class="tab-btn" onclick="showTab(event,'roi-intl')">ROI Intl (dist.)</button>
  <button class="tab-btn" onclick="showTab(event,'dispersao')">Dispersão</button>
  <button class="tab-btn" onclick="showTab(event,'categorias')">Por Categoria</button>
  <button class="tab-btn" onclick="showTab(event,'clusters')">Por Cluster Produtora</button>
  <button class="tab-btn" onclick="showTab(event,'ticket')">Ticket &amp; Capital Parado</button>
</div>

<!-- ── TAB 1: VISÃO GERAL ───────────────────────────────────────────── -->
<div id="tab-visao-geral" class="tab-content active">
  <div class="legenda">
    <div class="legenda-item"><div class="legenda-dot" style="background:#E07B54"></div>
      <span><b>Renúncia Pura</b> — só Lei do Audiovisual (Art.3/3-A/39), sem FSA</span></div>
    <div class="legenda-item"><div class="legenda-dot" style="background:#4C72B0"></div>
      <span><b>FSA Puro</b> — só Fundo Setorial do Audiovisual, sem renúncia</span></div>
    <div class="legenda-item"><div class="legenda-dot" style="background:#55A868"></div>
      <span><b>FSA + Renúncia — FSA Maj.</b> — ambos; FSA ≥ 50% do total</span></div>
    <div class="legenda-item"><div class="legenda-dot" style="background:#8172B2"></div>
      <span><b>FSA + Renúncia — Ren. Maj.</b> — ambos; Renúncia &gt; 50% do total</span></div>
  </div>
  <div class="card">
    <div class="section-title">Tabela Resumo — Principais Indicadores</div>
    {to_div(fig_tabela, "tabela_resumo")}
  </div>
  <div class="card">{to_toggle(fig_scatter_inv, "scatter_inv", 460, ["ROI Doméstico", "ROI Internacional"], len(GRUPOS))}</div>
</div>

<!-- ── TAB 2: RETORNO DOMÉSTICO ───────────────────────────────────── -->
<div id="tab-financeiro" class="tab-content">
  <div class="grid-2">
    <div class="card">{to_lazy(fig_comp_inv, "comp_inv", 380)}</div>
    <div class="card">{to_lazy(fig_roi_agr,  "roi_agr",  380)}</div>
  </div>
  <div class="card">{to_lazy(fig_inv_bil, "inv_bil", 420)}</div>
  <div class="card">{to_lazy(fig_inv_med, "inv_med", 380)}</div>
</div>

<!-- ── TAB 3: RETORNO INTERNACIONAL ───────────────────────────────── -->
<div id="tab-ret-intl" class="tab-content">
  <div class="card">{html_intl_map}</div>
  <div class="card">{to_lazy(fig_intl_lumiere, "intl_lumiere", 520)}</div>
  <div class="card">{to_lazy(fig_roi_int,      "roi_int_vg",   420)}</div>
</div>

<!-- ── TAB 4: ROI DOMÉSTICO ────────────────────────────────────────── -->
<div id="tab-roi-dom" class="tab-content">
  <div class="card">{to_lazy(fig_roi_dom, "roi_dom", 420)}</div>
</div>

<!-- ── TAB 4: ROI INTERNACIONAL ───────────────────────────────────── -->
<div id="tab-roi-intl" class="tab-content">
  <div class="grid-2">
    <div class="card">{to_lazy(fig_roi_int, "roi_int", 420)}</div>
    <div class="card">{to_lazy(fig_paises,  "paises",  420)}</div>
  </div>
</div>

<!-- ── TAB 5: DISPERSÃO ────────────────────────────────────────────── -->
<div id="tab-dispersao" class="tab-content">
  <div class="card">{to_lazy(fig_scatter, "scatter", 560, search="text")}</div>
</div>

<!-- ── TAB 6: POR CATEGORIA ────────────────────────────────────────── -->
<div id="tab-categorias" class="tab-content">
  <div class="card" style="background:#eef4fb;border-left:4px solid #1abc9c;">
    <div style="font-size:0.88rem;color:#333;">
      <b>Análise por Categoria de Chamada</b> — agrupa {len(df_cat):,} obras
      (excluídas <code>Renúncia</code>, <code>SAV/MINC / Arranjos Regionais</code>,
      <code>_tv_excluir</code> e <code>sem_categoria</code>) segundo a categoria FSA
      informada pela ANCINE. Cobre os instrumentos de pontuação (Bilheteria/Roteiro,
      Festivais), Complementação, Automáticos e Coprodução Internacional.
    </div>
  </div>
  <div class="card">{to_lazy(fig_cat_vol,      "cat_vol",      560)}</div>
  <div class="card">{to_lazy(fig_cat_roi,      "cat_roi",      560)}</div>
  <div class="grid-2">
    <div class="card">{to_lazy(fig_cat_box,      "cat_box",      540)}</div>
    <div class="card">{to_lazy(fig_cat_box_intl, "cat_box_intl", 540)}</div>
  </div>
  <div class="card">{to_lazy(fig_cat_tab,      "cat_tab",      520, search="rows")}</div>
  <div class="card">{to_lazy(fig_cat_scatter,  "cat_scatter",  580)}</div>
</div>

<!-- ── TAB 7: POR CLUSTER DE PRODUTORA ─────────────────────────────── -->
<div id="tab-clusters" class="tab-content">
  <div class="card">{to_lazy(fig_cl_scat,      "cl_scat",      580, search="text")}</div>
  <div class="card" style="background:#eef4fb;border-left:4px solid #1abc9c;">
    <div style="font-size:0.88rem;color:#333;line-height:1.5;">
      <b>Análise por Cluster de Produtora</b> — {len(df_prod):,} produtoras
      classificadas em 5 clusters segundo padrão de retorno:
      <ul style="margin:8px 0 0 20px;">
        <li><b style="color:#2E8B57;">Duplo Retorno</b> — receita ≥ R$ 2,5 mi <i>e</i> ROI intl máx ≥ 13 (cauda superior)</li>
        <li><b style="color:#4C72B0;">Retorno Doméstico</b> — receita ≥ R$ 10 mi, ou ROI doméstico > 0,6 com receita ≥ R$ 2,5 mi, sem ROI intl máx ≥ 13</li>
        <li><b style="color:#8172B2;">Retorno Internacional</b> — ROI intl máx ≥ 13 e sem critério de Retorno Doméstico</li>
        <li><b style="color:#D95F3A;">Fomento Baixo Retorno</b> — investimento > R$ 5 mi e ROI Internacional máx < 13</li>
        <li><b style="color:#A0A0A0;">Pequeno Porte</b> — demais produtoras (investimento ≤ R$ 5 mi)</li>
      </ul>
    </div>
  </div>
  <div class="card">{to_lazy(fig_cl_kpi,       "cl_kpi",       460)}</div>
  <div class="card">{to_lazy(fig_cl_roi,       "cl_roi",       480)}</div>
  <div class="grid-2">
    <div class="card">{to_lazy(fig_cl_box,       "cl_box",       480)}</div>
    <div class="card">{to_lazy(fig_cl_box_intl,  "cl_box_intl",  480)}</div>
  </div>
  <div class="card">{to_lazy(fig_cl_scat_intl, "cl_scat_intl", 580, search="text")}</div>
  <div class="card">{to_lazy(fig_cl_tab,       "cl_tab",       340, search="rows")}</div>
</div>

<!-- ── TAB 8: TICKET & CAPITAL PARADO ──────────────────────────────── -->
<div id="tab-ticket" class="tab-content">
  <div class="card" style="background:#eef4fb;border-left:4px solid #1abc9c;">
    <div style="font-size:0.88rem;color:#333;line-height:1.5;">
      <b>Ticket médio & Capital FSA sem retorno</b> — duas análises combinadas:
      <ul style="margin:8px 0 0 20px;">
        <li><b>Ticket médio</b>: investimento típico por obra e por produtora, segmentado
          por grupo de mecanismo, categoria de chamada e cluster de produtora. Média (barra)
          + mediana (diamante preto) para revelar assimetria. Eixos em escala log.</li>
        <li><b>Capital sem retorno</b>: quanto do FSA (R$ {fsa_universo/1e9:.2f} bi em
          {int(_tem_fsa.sum()):,} obras) foi alocado em obras sem retorno mensurável. Três
          cortes de severidade:
          <ul style="margin:4px 0 0 16px;">
            <li><b>Tier 1</b> — receita total = 0 <i>e</i> alcance intl = 0 (severíssimo)</li>
            <li><b>Tier 2</b> — receita total &lt; R$ 100 mil <i>e</i> alcance intl = 0 (definição-base)</li>
            <li><b>Tier 3</b> — bilheteria &lt; R$ 100 mil <i>e</i> alcance intl = 0 (moderado; ignora outras janelas)</li>
          </ul>
        </li>
      </ul>
    </div>
  </div>
  <div class="card">{to_lazy(fig_tck_obra_g, "tck_obra_g", 460)}</div>
  <div class="card">{to_lazy(fig_tck_obra_c, "tck_obra_c", 560)}</div>
  <div class="card">{to_lazy(fig_tck_prod,   "tck_prod",   480)}</div>
  <div class="card">{to_lazy(fig_tck_hist,   "tck_hist",   440)}</div>
  <div class="card">{to_lazy(fig_cp_kpi,     "cp_kpi",     460)}</div>
  <div class="card">{to_lazy(fig_cp_brk,     "cp_brk",     560)}</div>
</div>

<script>
var _initialized = {{}};
var _originalData = {{}};  // snapshot dos dados originais por divId (para filtros)

function _cloneData(data) {{
  // Plotly snapshot raso — retém refs de arrays originais, mas permite
  // reatribuir campos filtrados sem mutar a fonte
  return JSON.parse(JSON.stringify(data));
}}

function showTab(evt, name) {{
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  evt.currentTarget.classList.add('active');

  if (!_initialized[name]) {{
    _initialized[name] = true;
    var tab = document.getElementById('tab-' + name);
    tab.querySelectorAll('div[id]').forEach(function(div) {{
      var key = '__fig_' + div.id;
      if (window[key]) {{
        var fig = window[key];
        Plotly.newPlot(div, fig.data, fig.layout, {{responsive: true, displaylogo: false}});
        _originalData[div.id] = _cloneData(fig.data);
      }}
    }});
  }}
}}

// Toggle scatter entre ROI Dom (0) e ROI Intl (1)
function togSwitch(divId, modeIdx, btn) {{
  var div = document.getElementById(divId);
  var meta = window['__togMeta_' + divId];
  if (!div || !meta) return;
  var n = meta.n;
  var doUpdate = function() {{
    var vis = div.data.map(function(_, i) {{ return modeIdx === 0 ? i < n : i >= n; }});
    Plotly.update(div, {{visible: vis}}, {{'yaxis.title.text': meta.yTitles[modeIdx]}});
  }};
  if (div._fullLayout) {{ doUpdate(); }}
  else {{ setTimeout(doUpdate, 300); }}
  btn.closest('.toggle-bar').querySelectorAll('.tog-btn').forEach(function(b) {{
    b.classList.toggle('active', b === btn);
  }});
}}

// Filtra o gráfico com id=divId aplicando `query` (case-insensitive)
function filterFig(divId, query) {{
  var div = document.getElementById(divId);
  if (!div || !_originalData[divId]) return;
  var mode = div.getAttribute('data-search-mode') || 'text';
  var q = (query || '').trim().toLowerCase();
  var original = _originalData[divId];
  var countEl = document.getElementById('count_' + divId);
  var total = 0, visible = 0;

  if (mode === 'rows') {{
    // Tabela: filtra linhas cujo 1º valor (categoria/cluster) casa com q
    var newData = _cloneData(original);
    newData.forEach(function(trace) {{
      if (trace.type !== 'table' || !trace.cells || !trace.cells.values) return;
      var cols = trace.cells.values;
      var firstCol = cols[0] || [];
      total = firstCol.length;
      var keep = [];
      for (var i = 0; i < firstCol.length; i++) {{
        var v = String(firstCol[i] || '').toLowerCase();
        if (!q || v.indexOf(q) !== -1) keep.push(i);
      }}
      visible = keep.length;
      // Reconstrói cada coluna apenas com as linhas selecionadas
      trace.cells.values = cols.map(function(col) {{
        return keep.map(function(i) {{ return col[i]; }});
      }});
      // Se houver fill_color array por linha, filtra também
      if (trace.cells.fill_color && Array.isArray(trace.cells.fill_color)) {{
        trace.cells.fill_color = trace.cells.fill_color.map(function(col) {{
          if (!Array.isArray(col)) return col;
          return keep.map(function(i) {{ return col[i]; }});
        }});
      }}
    }});
    Plotly.react(divId, newData, div.layout || {{}}, {{responsive: true, displaylogo: false}});
  }} else {{
    // Scatter: filtra pontos cujo `text` (ou customdata[0]) contém q
    var newData2 = _cloneData(original);
    newData2.forEach(function(trace) {{
      if (!trace.x || !trace.y) return;
      var labels = trace.text;
      if (!labels && trace.customdata) {{
        labels = trace.customdata.map(function(r) {{
          return Array.isArray(r) ? r[0] : r;
        }});
      }}
      var n = trace.x.length;
      total += n;
      if (!q) {{ visible += n; return; }}
      var keep = [];
      for (var i = 0; i < n; i++) {{
        var v = String((labels && labels[i]) || '').toLowerCase();
        if (v.indexOf(q) !== -1) keep.push(i);
      }}
      visible += keep.length;
      trace.x = keep.map(function(i) {{ return trace.x[i]; }});
      trace.y = keep.map(function(i) {{ return trace.y[i]; }});
      if (labels) {{
        trace.text = keep.map(function(i) {{ return labels[i]; }});
      }}
      if (trace.customdata) {{
        trace.customdata = keep.map(function(i) {{ return trace.customdata[i]; }});
      }}
      // Tamanhos do marker podem ser array
      if (trace.marker && Array.isArray(trace.marker.size)) {{
        var ms = trace.marker.size;
        trace.marker = Object.assign({{}}, trace.marker, {{
          size: keep.map(function(i) {{ return ms[i]; }}),
        }});
      }}
    }});
    Plotly.react(divId, newData2, div.layout || {{}}, {{responsive: true, displaylogo: false}});
  }}

  if (countEl) {{
    countEl.textContent = q
      ? (visible + ' / ' + total + ' visíveis')
      : '';
  }}
}}
</script>

<script>
{js_intl_map}
</script>

</body>
</html>
"""

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
OUTPUT.write_text(HTML, encoding="utf-8")
print(f"\nPainel gerado: {OUTPUT}")
print("Abra no navegador para visualizar.")
