#!/usr/bin/env python3
"""
gerar_graficos_analise.py
=========================
Regenera os gráficos do analise.html com estética de publicação científica
(Nature-style: fundo branco, tipografia limpa, sem bordas desnecessárias).

Substitui os PNGs base64 originais (vindos do DOCX) por versões matplotlib
geradas a partir dos datasets da pipeline.

Uso:
    python scripts/gerar_graficos_analise.py

Requisitos: pandas, numpy, matplotlib
"""

import os, re, base64, io
import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.ticker import FuncFormatter, PercentFormatter, MaxNLocator
from matplotlib.gridspec import GridSpec

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML_PATH = os.path.join(ROOT, "output_final", "analise.html")

# ── Nature-style rcParams ──────────────────────────────────────────────────────
mpl.rcParams.update({
    "font.family":         "sans-serif",
    "font.sans-serif":     ["Arial", "Helvetica", "DejaVu Sans", "sans-serif"],
    "svg.fonttype":        "none",
    "pdf.fonttype":        42,
    "font.size":           20,
    "axes.titlesize":      22,
    "axes.titleweight":    "bold",
    "axes.titlepad":       20,
    "axes.labelsize":      20,
    "axes.labelcolor":     "#333333",
    "axes.spines.right":   False,
    "axes.spines.top":     False,
    "axes.linewidth":      1.5,
    "axes.edgecolor":      "#aaaaaa",
    "axes.facecolor":      "#fafafa",
    "figure.facecolor":    "white",
    "xtick.labelsize":     18,
    "ytick.labelsize":     18,
    "xtick.color":         "#555555",
    "ytick.color":         "#555555",
    "xtick.major.size":    6,
    "ytick.major.size":    6,
    "legend.frameon":      False,
    "legend.fontsize":     18,
    "grid.alpha":          0.45,
    "grid.linewidth":      1.0,
    "grid.color":          "#cccccc",
    "axes.grid":           True,
    "axes.grid.axis":      "y",
    "figure.dpi":          110,
    "savefig.dpi":         160,
    "savefig.facecolor":   "white",
    "savefig.bbox":        "tight",
    "savefig.pad_inches":  0.3,
})

# ── Paleta científica (muted, distinguível) ────────────────────────────────────
C_BLUE   = "#2166ac"
C_ORANGE = "#e08026"
C_GREEN  = "#4dac26"
C_RED    = "#d7191c"
C_PURPLE = "#762a83"
C_TEAL   = "#1b9e77"
C_GREY   = "#888888"
C_AMBER  = "#b8860b"

CLUSTER_COLORS = {
    "Duplo Retorno":         C_TEAL,
    "Retorno Doméstico":     C_BLUE,
    "Retorno Internacional": C_ORANGE,
    "Fomento Baixo Retorno": C_RED,
    "Pequeno Porte":         C_GREY,
}

CAT_SHORT = {
    "FSA Pontuação Bilheteria e Roteiro — Distribuidora": "Seletivo\nDistribuidora",
    "FSA Pontuação Bilheteria e Roteiro — Produtora":     "Seletivo\nProdutora",
    "FSA Pontuação Festivais e Roteiro":                  "Seletivo\nFestivais",
    "FSA Automático Bilheteria":                          "Automático\nBilheteria",
    "FSA Automático Festivais":                           "Automático\nFestivais",
    "FSA Comercialização / Distribuição":                 "Comercialização",
    "FSA Complementação":                                 "Complementação",
    "FSA Coprodução Internacional":                       "Coprodução\nIntl",
    "FSA SAV/MINC / Arranjos Regionais":                  "SAV/MINC\nRegionais",
}

# ── Helpers ────────────────────────────────────────────────────────────────────
def load_csv(name):
    path = os.path.join(ROOT, "resultados", "datasets", name)
    df = pd.read_csv(path, encoding="utf-8-sig", sep=None, engine="python")
    return df


def _num(df, col):
    return pd.to_numeric(df[col], errors="coerce").fillna(0)


def fig_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()


def make_fig(w=20, h=9):
    return plt.figure(figsize=(w, h))


def fmt_b(v, p):
    if abs(v) >= 1e9:  return f"R${v/1e9:.1f}bi"
    if abs(v) >= 1e6:  return f"R${v/1e6:.0f}M"
    if abs(v) >= 1e3:  return f"R${v/1e3:.0f}K"
    return f"R${v:.0f}"


def add_bar_labels(ax, bars, fmt="{:.2f}", color="#333333", fontsize=18, offset=None):
    for b in bars:
        h = b.get_height()
        if h < 0.001: continue
        off = offset if offset else max(ax.get_ylim()[1] * 0.01, h * 0.02)
        ax.text(b.get_x() + b.get_width() / 2, h + off,
                fmt.format(h), ha="center", va="bottom",
                fontsize=fontsize, color=color)


def add_hbar_labels(ax, bars, fmt="{:.2f}", color="#333333", fontsize=18):
    for b in bars:
        w = b.get_width()
        if abs(w) < 0.001: continue
        ax.text(w + ax.get_xlim()[1] * 0.01, b.get_y() + b.get_height() / 2,
                fmt.format(w), va="center", fontsize=fontsize, color=color)


# ── Carrega datasets ───────────────────────────────────────────────────────────
df_cham  = load_csv("base_nivel_chamada.csv")
df_inv   = load_csv("base_nivel_investimento.csv")
df_obra  = load_csv("base_nivel_obra.csv")

# ── Filtro: só obras com bilheteria > 0 ────────────────────────────────────────
df_obra = df_obra[_num(df_obra, "bilheteria_deflac") > 0].copy().reset_index(drop=True)

# ── Produtoras: carrega do painel.html (1243 prod., todos os mecanismos/clusters)
_PAINEL_PATH = os.path.join(ROOT, "output_final", "painel.html")
def _load_prod_painel():
    import json as _json
    with open(_PAINEL_PATH, encoding="utf-8") as _f:
        _html = _f.read()
    _m = re.search(r'const PROD\s*=\s*(\[.+?\]);\s*(?:const|var|let|//)', _html, re.DOTALL)
    _data = _json.loads(_m.group(1))
    _df = pd.DataFrame(_data)
    _CL = {"duplo": "Duplo Retorno", "dom": "Retorno Doméstico", "intl": "Retorno Internacional",
           "sem_retorno": "Fomento Baixo Retorno", "pequeno": "Pequeno Porte"}
    _df["cluster"]                   = _df["cl"].map(_CL).fillna(_df["cl"])
    _df["CNPJ_produtora"]            = _df["nm"]
    _df["razao_social"]              = _df["nm"]
    _df["n_obras"]                   = pd.to_numeric(_df["n"],         errors="coerce").fillna(0)
    _df["investimento_fsa_deflac"]   = pd.to_numeric(_df["inv_fsa_d"], errors="coerce").fillna(0)
    _df["investimento_total_deflac"] = pd.to_numeric(_df["inv_def"],   errors="coerce").fillna(0)
    _df["receita_total_deflac"]      = pd.to_numeric(_df["rec_def"],   errors="coerce").fillna(0)
    _df["roi_intl_medio"]            = pd.to_numeric(_df["ria"],       errors="coerce").fillna(0)
    _df["roi_dom_fsa_deflac"]        = _df.apply(
        lambda r: r["receita_total_deflac"] / r["investimento_fsa_deflac"]
        if r["investimento_fsa_deflac"] > 0 else 0.0, axis=1)
    return _df

df_prod = _load_prod_painel()

# ── df_cat: agregação por categoria FSA calculada do df_obra filtrado ──────────
# Substitui df_cham nos charts de categoria para garantir recorte consistente
_obra_fsa = df_obra[df_obra["categoria"].astype(str).str.startswith("FSA")].copy()
df_cat = (_obra_fsa.groupby("categoria")
          .apply(lambda g: pd.Series({
              "roi_dom_total_agregado":    _num(g, "receita_total_deflac").sum() / max(_num(g, "investimento_total_deflac").sum(), 1),
              "roi_intl_medio":            _num(g, "roi_internacional_0_100").mean(),
              "n_obras":                   len(g),
              "investimento_total_deflac": _num(g, "investimento_total_deflac").sum(),
              "pontuacao_festivais_media": _num(g, "pontuacao_festivais").mean(),
              "pontuacao_festivais_max":   _num(g, "pontuacao_festivais").max(),
              "critica_media":             _num(g, "critica_indice_1_5").mean(),
          }))
          .reset_index())


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 1 — Investimento total e receita estimada por ano
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig1():
    df = df_obra.copy()
    df["ano"]      = pd.to_numeric(df["ano"], errors="coerce")
    df["inv_tot"]  = _num(df, "investimento_total_deflac")
    df["bil"]      = _num(df, "bilheteria_deflac")
    df["jan"]      = _num(df, "outras_janelas_deflac")
    df["receita"]  = df["bil"] + df["jan"]

    grp = (df.groupby("ano")
             .agg(inv=("inv_tot","sum"), receita=("receita","sum"), n=("CPB","count"))
             .reset_index()
             .sort_values("ano"))

    fig, ax = plt.subplots(figsize=(22, 9))
    x = np.arange(len(grp))
    w = 0.38
    b1 = ax.bar(x - w/2, grp["inv"]   / 1e6, w, color=C_BLUE,   alpha=0.85, label="Investimento total (FSA + Renúncia)", zorder=3)
    b2 = ax.bar(x + w/2, grp["receita"]/ 1e6, w, color=C_ORANGE, alpha=0.85, label="Receita estimada (bilheteria + janelas)", zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels(grp["ano"].astype(int), fontsize=18)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.set_ylabel("R$ milhões (deflac. 2024)")
    ax.set_title("Investimento total e receita estimada por ano de produção")
    ax.legend(loc="upper left", fontsize=18)

    ax2 = ax.twinx()
    ax2.plot(x, grp["n"], "o--", color=C_PURPLE, linewidth=1.5, markersize=10,
             label="Nº de obras", zorder=4, alpha=0.8)
    ax2.set_ylabel("Nº de obras", color=C_PURPLE, fontsize=18)
    ax2.tick_params(axis="y", colors=C_PURPLE, labelsize=16)
    ax2.spines["right"].set_visible(True)
    ax2.spines["right"].set_color("#cccccc")
    ax2.spines["top"].set_visible(False)
    ax2.grid(False)
    ax2.legend(loc="upper right", fontsize=18)

    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA A1 — Cobertura de dados por fonte de evidência
# ══════════════════════════════════════════════════════════════════════════════
def chart_figA1():
    df = df_obra.copy()
    n_total = len(df)

    def _pct(col):
        s = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return (s > 0).sum() / n_total * 100

    fontes = {
        "Bilheteria": "bilheteria_deflac",
        "Outras Janelas": "outras_janelas_deflac",
        "Festivais (pontuação)": "pontuacao_festivais",
        "Admissões Lumière/CNC": "adm_eu_lumiere",
        "VOD Internacional": "vod_n_plataformas",
        "Crítica": "critica_indice_1_5",
        "Citações acadêmicas": "cita_n_papers",
    }

    labels = list(fontes.keys())
    pcts   = [_pct(c) for c in fontes.values()]
    order  = np.argsort(pcts)
    labels = [labels[i] for i in order]
    pcts   = [pcts[i] for i in order]

    colors = [C_BLUE if p >= 50 else C_ORANGE if p >= 20 else C_GREY for p in pcts]

    fig, ax = plt.subplots(figsize=(20, 8))
    bars = ax.barh(labels, pcts, color=colors, alpha=0.85, edgecolor="white", linewidth=0.5)
    ax.set_xlabel("% das obras com dado disponível")
    ax.set_title("Cobertura de dados por fonte de evidência na base de obras")
    ax.xaxis.set_major_formatter(PercentFormatter())
    ax.set_xlim(0, 105)
    ax.grid(axis="x")
    ax.grid(axis="y", alpha=0)
    for b, p in zip(bars, pcts):
        ax.text(p + 1, b.get_y() + b.get_height() / 2,
                f"{p:.1f}%", va="center", fontsize=18, color="#333333")

    handles = [mpatches.Patch(color=C_BLUE, label="≥ 50% cobertura"),
               mpatches.Patch(color=C_ORANGE, label="20–50% cobertura"),
               mpatches.Patch(color=C_GREY, label="< 20% cobertura")]
    ax.legend(handles=handles, loc="lower right", fontsize=16)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 2 — Distribuição do investimento: FSA puro, renúncia pura, misto
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig2():
    df = df_obra.copy()
    df["fsa_d"] = _num(df, "investimento_fsa_deflac")
    df["ren_d"] = _num(df, "investimento_renuncia_total_deflac")
    df["grupo"] = "FSA Puro"
    df.loc[(df["fsa_d"] > 0) & (df["ren_d"] > 0), "grupo"] = "Misto (FSA + Renúncia)"
    df.loc[(df["fsa_d"] == 0) & (df["ren_d"] > 0), "grupo"] = "Renúncia Pura"

    agg = df.groupby("grupo").agg(
        n_obras=("CPB", "count"),
        inv_tot=("investimento_total_deflac", lambda s: pd.to_numeric(s, errors="coerce").fillna(0).sum()),
        bil=("bilheteria_deflac", lambda s: pd.to_numeric(s, errors="coerce").fillna(0).sum()),
    ).reset_index()
    order = ["FSA Puro", "Misto (FSA + Renúncia)", "Renúncia Pura"]
    agg = agg.set_index("grupo").reindex(order).reset_index()
    colors = [C_BLUE, C_PURPLE, C_ORANGE]

    fig, axes = plt.subplots(1, 3, figsize=(24, 8))
    titles = ["Nº de obras", "Investimento total (R$ bi deflac.)", "Bilheteria total (R$ bi deflac.)"]
    data = [agg["n_obras"], agg["inv_tot"] / 1e9, agg["bil"] / 1e9]
    fmts = ["{:.0f}", "R${:.2f}bi", "R${:.2f}bi"]

    for ax, title, vals, colors_, fmt in zip(axes, titles, data, [colors]*3, fmts):
        bars = ax.bar(agg["grupo"], vals, color=colors_, alpha=0.85, edgecolor="white", linewidth=0.5)
        ax.set_title(title, fontsize=20)
        ax.set_xticklabels(agg["grupo"], fontsize=16, rotation=0, wrap=True)
        ax.tick_params(axis="x", labelsize=16)
        for b, v in zip(bars, vals):
            ax.text(b.get_x() + b.get_width() / 2,
                    b.get_height() + max(vals) * 0.02,
                    fmt.format(v), ha="center", va="bottom", fontsize=18)
    fig.suptitle("Distribuição do investimento entre FSA puro, renúncia pura e arranjos mistos",
                 fontsize=22, fontweight="bold", y=1.02)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 3 — ROI doméstico por grupo de mecanismo
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig3():
    df = df_obra.copy()
    df["fsa_d"] = _num(df, "investimento_fsa_deflac")
    df["ren_d"] = _num(df, "investimento_renuncia_total_deflac")
    df["bil"]   = _num(df, "bilheteria_deflac")
    df["jan"]   = _num(df, "outras_janelas_deflac")
    df["rec"]   = df["bil"] + df["jan"]
    df["inv"]   = _num(df, "investimento_total_deflac")
    df["grupo"] = "FSA Puro"
    df.loc[(df["fsa_d"] > 0) & (df["ren_d"] > 0), "grupo"] = "Misto"
    df.loc[(df["fsa_d"] == 0) & (df["ren_d"] > 0), "grupo"] = "Renúncia Pura"

    # Aggregate: total ROI = sum(rec) / sum(inv)
    agg = df.groupby("grupo").apply(
        lambda g: pd.Series({
            "roi_bil": g["bil"].sum() / max(g["inv"].sum(), 1),
            "roi_rec": g["rec"].sum() / max(g["inv"].sum(), 1),
            "n": len(g),
        })
    ).reset_index()
    order = ["FSA Puro", "Misto", "Renúncia Pura"]
    agg = agg.set_index("grupo").reindex(order).reset_index()

    fig, ax = plt.subplots(figsize=(18, 8))
    x = np.arange(len(agg))
    w = 0.35
    b1 = ax.bar(x - w/2, agg["roi_bil"], w, color=C_BLUE, alpha=0.85, label="ROI Bilheteria / Investimento")
    b2 = ax.bar(x + w/2, agg["roi_rec"], w, color=C_TEAL,  alpha=0.85, label="ROI Receita total / Investimento")
    ax.axhline(1.0, color="#666666", linewidth=1, linestyle="--", alpha=0.6, label="Paridade (ROI = 1)")
    ax.set_xticks(x)
    ax.set_xticklabels(agg["grupo"], fontsize=18)
    ax.set_ylabel("ROI agregado")
    ax.set_title("ROI doméstico agregado por grupo de mecanismo de financiamento")
    ax.legend(fontsize=18)
    add_bar_labels(ax, b1, fmt="{:.2f}x", color=C_BLUE)
    add_bar_labels(ax, b2, fmt="{:.2f}x", color=C_TEAL)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 21 — Bubble chart: categorias no espaço ROI dom × ROI intl
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig21():
    df = (df_cat
          .assign(roi_dom=lambda x: x["roi_dom_total_agregado"],
                  roi_intl=lambda x: x["roi_intl_medio"],
                  inv=lambda x: x["investimento_total_deflac"],
                  label=lambda x: x["categoria"].map(CAT_SHORT).fillna(x["categoria"])))

    fig, ax = plt.subplots(figsize=(22, 12))
    scatter = ax.scatter(df["roi_dom"], df["roi_intl"],
                         s=df["n_obras"] * 2.5, alpha=0.75,
                         c=df["inv"], cmap="Blues_r",
                         vmin=0, vmax=df["inv"].max(),
                         edgecolors=C_BLUE, linewidths=0.6, zorder=3)

    med_dom  = df["roi_dom"].median()
    med_intl = df["roi_intl"].median()
    ax.axvline(med_dom,  color="#aaaaaa", linewidth=1, linestyle="--", alpha=0.7)
    ax.axhline(med_intl, color="#aaaaaa", linewidth=1, linestyle="--", alpha=0.7)
    ax.text(ax.get_xlim()[1]*0.98, med_intl + 0.1, "mediana intl.", ha="right", fontsize=16, color="#888888")
    ax.text(med_dom + 0.01, ax.get_ylim()[1]*0.97, "mediana dom.", ha="left", fontsize=16, color="#888888")

    for _, row in df.iterrows():
        ax.annotate(row["label"], (row["roi_dom"], row["roi_intl"]),
                    textcoords="offset points", xytext=(6, 4),
                    fontsize=16, color="#333333")

    cbar = fig.colorbar(scatter, ax=ax, shrink=0.6, pad=0.01)
    cbar.set_label("Investimento total (R$ deflac.)", fontsize=16)
    cbar.ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v/1e6:.0f}M"))
    ax.set_xlabel("ROI Doméstico (receita / invest. total)")
    ax.set_ylabel("ROI Internacional médio")
    ax.set_title("Posicionamento de cada categoria FSA: retorno doméstico vs. alcance internacional")
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 5 — Volume de investimento por categoria FSA
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig5():
    df = (df_cat
          .assign(inv=lambda x: x["investimento_total_deflac"],
                  n=lambda x: x["n_obras"],
                  label=lambda x: x["categoria"].map(CAT_SHORT).fillna(x["categoria"].str.replace("FSA ", "")))
          .sort_values("inv"))

    fig, ax = plt.subplots(figsize=(20, 10))
    colors = [C_BLUE if v > df["inv"].median() else C_GREY for v in df["inv"]]
    bars = ax.barh(df["label"], df["inv"] / 1e6, color=colors, alpha=0.85,
                   edgecolor="white", linewidth=0.5)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.set_xlabel("Investimento total deflacionado (R$ milhões, R$ 2024)")
    ax.set_title("Categorias FSA por volume de investimento total deflacionado")
    ax.grid(axis="x"); ax.grid(axis="y", alpha=0)
    for b, v in zip(bars, df["inv"] / 1e6):
        ax.text(v + df["inv"].max() / 1e6 * 0.01, b.get_y() + b.get_height() / 2,
                f"R${v:.0f}M", va="center", fontsize=18, color="#333333")
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 6 — ROI Total por categoria FSA
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig6():
    df = (df_cat
          .assign(roi=lambda x: x["roi_dom_total_agregado"],
                  n=lambda x: x["n_obras"],
                  label=lambda x: x["categoria"].map(CAT_SHORT).fillna(x["categoria"].str.replace("FSA ", "")))
          .sort_values("roi"))

    fig, ax = plt.subplots(figsize=(20, 10))
    colors = [C_GREEN if v >= 1 else C_ORANGE if v >= 0.5 else C_RED for v in df["roi"]]
    bars = ax.barh(df["label"], df["roi"], color=colors, alpha=0.85, edgecolor="white", linewidth=0.5)
    ax.axvline(1.0, color="#666666", linewidth=1, linestyle="--", alpha=0.7, label="Paridade (ROI = 1)")
    ax.set_xlabel("ROI Total (receita / investimento total)")
    ax.set_title("ROI Total Deflacionado por categoria FSA")
    ax.grid(axis="x"); ax.grid(axis="y", alpha=0)
    for b, v in zip(bars, df["roi"]):
        ax.text(v + 0.01, b.get_y() + b.get_height() / 2,
                f"{v:.2f}x", va="center", fontsize=18, color="#333333")
    handles = [mpatches.Patch(color=C_GREEN, label="ROI ≥ 1"),
               mpatches.Patch(color=C_ORANGE, label="ROI 0,5–1"),
               mpatches.Patch(color=C_RED, label="ROI < 0,5")]
    ax.legend(handles=handles, fontsize=16, loc="lower right")
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 7 — Score internacional total e médio por categoria
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig7():
    df = df_cat.copy()
    df["score_med"] = _num(df, "pontuacao_festivais_media")
    df["total"]     = df["score_med"] * df["n_obras"]
    df["roi_intl"]  = _num(df, "roi_intl_medio")
    df["label"]     = df["categoria"].map(CAT_SHORT).fillna(df["categoria"].str.replace("FSA ", ""))

    def _cor(lbl):
        s = str(lbl)
        if "Distrib" in s: return C_GREEN
        if "Produt"  in s: return C_BLUE
        if "Festiv"  in s: return C_ORANGE
        if "Coprod"  in s: return C_PURPLE
        if "Autom"   in s: return C_TEAL
        return C_GREY

    colors = [_cor(l) for l in df["label"]]

    fig, axes = plt.subplots(1, 2, figsize=(24, 10))

    # Left: scatter score médio vs total acumulado
    ax = axes[0]
    ax.scatter(df["score_med"], df["total"],
               s=df["n_obras"] * 20, c=colors, alpha=0.80,
               edgecolors="#555555", linewidths=1.0, zorder=3)
    for _, row in df.iterrows():
        ax.annotate(row["label"], (row["score_med"], row["total"]),
                    textcoords="offset points", xytext=(10, 5),
                    fontsize=16, color="#333333")
    ax.set_xlabel("Score médio por obra (festivais)")
    ax.set_ylabel("Score total acumulado (festivais)")
    ax.set_title("Média por obra vs. total acumulado\npor categoria FSA")
    ax.grid(True, alpha=0.3)

    # Right: ROI Internacional médio por categoria
    ax2 = axes[1]
    df2 = df.sort_values("roi_intl")
    bars = ax2.barh(df2["label"], df2["roi_intl"], color=C_ORANGE, alpha=0.85, edgecolor="white")
    ax2.set_xlabel("ROI Internacional médio")
    ax2.set_title("ROI Internacional médio\npor categoria FSA")
    ax2.grid(axis="x"); ax2.grid(axis="y", alpha=0)
    for b, v in zip(bars, df2["roi_intl"]):
        ax2.text(v + df2["roi_intl"].max() * 0.01, b.get_y() + b.get_height() / 2,
                 f"{v:.2f}", va="center", fontsize=18, color="#333333")

    fig.suptitle("Score de festivais e retorno internacional por categoria FSA", fontsize=22, fontweight="bold")
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 25 — Índice crítico médio por categoria
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig25():
    # % obras com presença em festivais (pontuação > 0) por categoria
    obra_fsa = df_obra[df_obra["categoria"].astype(str).str.startswith("FSA")].copy()
    pct = (obra_fsa.groupby("categoria")
           .apply(lambda g: pd.Series({
               "pct_festival": (_num(g, "pontuacao_festivais") > 0).mean() * 100,
               "score_med":    _num(g, "pontuacao_festivais").mean(),
               "n":            len(g),
           }))
           .reset_index())
    pct["label"] = pct["categoria"].map(CAT_SHORT).fillna(pct["categoria"].str.replace("FSA ", ""))
    pct = pct[pct["n"] >= 5].sort_values("pct_festival")
    mean_pct = pct["pct_festival"].mean()

    fig, ax = plt.subplots(figsize=(20, 9))
    colors = [C_GREEN if v >= 50 else C_ORANGE if v >= 25 else C_RED for v in pct["pct_festival"]]
    bars = ax.barh(pct["label"], pct["pct_festival"], color=colors, alpha=0.85, edgecolor="white")
    ax.axvline(mean_pct, color="#666666", linewidth=1.5, linestyle="--",
               alpha=0.7, label=f"Média (categ. ≥ 5 obras) = {mean_pct:.1f}%")
    ax.set_xlabel("% de obras com presença em festivais")
    ax.set_title("Presença em festivais por categoria FSA\n(obras com bilheteria > 0, pontuação > 0)")
    ax.xaxis.set_major_formatter(PercentFormatter())
    ax.set_xlim(0, 115)
    ax.grid(axis="x"); ax.grid(axis="y", alpha=0)
    for b, (v, s) in zip(bars, zip(pct["pct_festival"], pct["score_med"])):
        ax.text(v + 1, b.get_y() + b.get_height() / 2,
                f"{v:.0f}%  (méd {s:.1f}pts)", va="center", fontsize=16, color="#333333")
    ax.legend(fontsize=18)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 4 — % obras com ROI internacional ≥ 13 por grupo
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig4():
    df = df_obra.copy()
    df["fsa_d"] = _num(df, "investimento_fsa_deflac")
    df["ren_d"] = _num(df, "investimento_renuncia_total_deflac")
    df["roi_i"] = _num(df, "roi_internacional_0_100")
    df["grupo"] = "FSA Puro"
    df.loc[(df["fsa_d"] > 0) & (df["ren_d"] > 0), "grupo"] = "Misto"
    df.loc[(df["fsa_d"] == 0) & (df["ren_d"] > 0), "grupo"] = "Renúncia Pura"
    df["intl_qualif"] = (df["roi_i"] >= 13).astype(int)

    agg = df.groupby("grupo").agg(
        total=("CPB", "count"),
        qualif=("intl_qualif", "sum")
    ).reset_index()
    agg["pct"] = agg["qualif"] / agg["total"] * 100
    order = ["FSA Puro", "Misto", "Renúncia Pura"]
    agg = agg.set_index("grupo").reindex(order).reset_index()

    fig, ax = plt.subplots(figsize=(18, 8))
    bars = ax.bar(agg["grupo"], agg["pct"], color=[C_BLUE, C_PURPLE, C_ORANGE], alpha=0.85, edgecolor="white")
    ax.yaxis.set_major_formatter(PercentFormatter())
    ax.set_ylabel("% de obras com ROI Internacional ≥ 13")
    ax.set_title("Percentual de obras com ROI Internacional qualificado (≥ 13) por grupo de mecanismo")
    for b, (v, n, q) in zip(bars, zip(agg["pct"], agg["total"], agg["qualif"])):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.3,
                f"{v:.1f}%\n({q:.0f}/{n:.0f})", ha="center", va="bottom", fontsize=18)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 8 — Distribuição de produtoras por cluster
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig8():
    df = df_prod.copy()
    agg = df.groupby("cluster").agg(
        n=("CNPJ_produtora", "count"),
        fsa=("investimento_fsa_deflac", lambda s: _num(pd.DataFrame({"v": s}), "v").sum()),
        rec=("receita_total_deflac",    lambda s: _num(pd.DataFrame({"v": s}), "v").sum()),
    ).reset_index()
    agg["fsa_sum"] = agg.apply(lambda r: pd.to_numeric(r["fsa"], errors="coerce") if not isinstance(r["fsa"], float) else r["fsa"], axis=1)

    order = [c for c in CLUSTER_COLORS if c in agg["cluster"].values]
    agg = agg.set_index("cluster").reindex(order).reset_index().dropna(subset=["n"])

    fig, axes = plt.subplots(1, 2, figsize=(24, 9))

    # Nº de produtoras
    ax = axes[0]
    colors = [CLUSTER_COLORS.get(c, C_GREY) for c in agg["cluster"]]
    bars = ax.bar(range(len(agg)), agg["n"], color=colors, alpha=0.85, edgecolor="white")
    ax.set_xticks(range(len(agg)))
    ax.set_xticklabels([c.replace(" ", "\n") for c in agg["cluster"]], fontsize=16)
    ax.set_title("Nº de produtoras por cluster")
    ax.set_ylabel("Nº de produtoras")
    for b, v in zip(bars, agg["n"]):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + agg["n"].max() * 0.02,
                f"{int(v)}", ha="center", va="bottom", fontsize=18)

    # Investimento FSA
    ax2 = axes[1]
    df2 = df.copy()
    df2["fsa"] = _num(df2, "investimento_fsa_deflac")
    agg2 = df2.groupby("cluster")["fsa"].sum().reindex(order).reset_index()
    colors2 = [CLUSTER_COLORS.get(c, C_GREY) for c in agg2["cluster"]]
    bars2 = ax2.bar(range(len(agg2)), agg2["fsa"] / 1e6, color=colors2, alpha=0.85, edgecolor="white")
    ax2.set_xticks(range(len(agg2)))
    ax2.set_xticklabels([c.replace(" ", "\n") for c in agg2["cluster"]], fontsize=16)
    ax2.set_title("Investimento FSA por cluster (R$ M deflac.)")
    ax2.set_ylabel("R$ milhões (deflac. 2024)")
    ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    for b, v in zip(bars2, agg2["fsa"] / 1e6):
        ax2.text(b.get_x() + b.get_width() / 2, b.get_height() + agg2["fsa"].max() / 1e6 * 0.02,
                 f"R${v:.0f}M", ha="center", va="bottom", fontsize=18)

    fig.suptitle("Distribuição de produtoras por cluster (retorno internacional ≥ 13)", fontsize=22, fontweight="bold")
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 9 — Investimento e receita por cluster
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig9():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    df["tot"] = _num(df, "investimento_total_deflac")
    df["rec"] = _num(df, "receita_total_deflac")
    order = [c for c in CLUSTER_COLORS if c in df["cluster"].values]
    agg = df.groupby("cluster").agg(
        fsa=("fsa", "sum"), tot=("tot", "sum"), rec=("rec", "sum"), n=("CNPJ_produtora", "count")
    ).reindex(order).reset_index()

    fig, ax = plt.subplots(figsize=(22, 9))
    x = np.arange(len(agg))
    w = 0.27
    b1 = ax.bar(x - w, agg["tot"] / 1e6, w, color=C_BLUE, alpha=0.85, label="Investimento total")
    b2 = ax.bar(x,     agg["fsa"] / 1e6, w, color=C_PURPLE, alpha=0.85, label="Investimento FSA")
    b3 = ax.bar(x + w, agg["rec"] / 1e6, w, color=C_ORANGE, alpha=0.85, label="Receita estimada")
    ax.set_xticks(x)
    ax.set_xticklabels([c.replace(" ", "\n") for c in agg["cluster"]], fontsize=18)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.set_ylabel("R$ milhões (deflac. 2024)")
    ax.set_title("Investimento e receita estimada por cluster de produtora")
    ax.legend(fontsize=18)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 10 — ROI dom. e intensidade internacional por cluster
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig10():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    df["tot"] = _num(df, "investimento_total_deflac")
    df["rec"] = _num(df, "receita_total_deflac")
    df["roi_i"] = _num(df, "roi_intl_medio")
    # ROI dom: sum(rec) / sum(invest. total) per cluster — mesmo denominador do painel
    order = [c for c in CLUSTER_COLORS if c in df["cluster"].values]
    agg = df.groupby("cluster").apply(lambda g: pd.Series({
        "roi_dom": g["rec"].sum() / max(g["tot"].sum(), 1),
        "roi_intl": g["roi_i"].mean(),
        "n": len(g),
    })).reindex(order).reset_index()

    fig, ax = plt.subplots(figsize=(22, 9))
    x = np.arange(len(agg))
    w = 0.35
    b1 = ax.bar(x - w/2, agg["roi_dom"],  w, color=C_BLUE,   alpha=0.85, label="ROI Doméstico agregado")
    b2 = ax.bar(x + w/2, agg["roi_intl"], w, color=C_ORANGE, alpha=0.85, label="ROI Internacional médio")
    ax.axhline(1.0, color="#666666", linewidth=1, linestyle="--", alpha=0.6, label="Paridade")
    ax.set_xticks(x)
    ax.set_xticklabels([c.replace(" ", "\n") for c in agg["cluster"]], fontsize=18)
    ax.set_ylabel("ROI")
    ax.set_title("ROI doméstico agregado e intensidade internacional média por cluster")
    ax.legend(fontsize=18)
    add_bar_labels(ax, b1, fmt="{:.2f}x", color=C_BLUE)
    add_bar_labels(ax, b2, fmt="{:.2f}x", color=C_ORANGE)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 22 — Scatter: produtoras no espaço ROI dom × ROI intl (cor = cluster)
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig22():
    df = df_prod.copy()
    df["roi_d"] = _num(df, "roi_dom_fsa_deflac")
    df["roi_i"] = _num(df, "roi_intl_medio")
    df["fsa"]   = _num(df, "investimento_fsa_deflac")

    # Cap extremes for readability
    df["roi_d"] = df["roi_d"].clip(-2, 15)
    df["roi_i"] = df["roi_i"].clip(0, 80)

    fig, ax = plt.subplots(figsize=(22, 12))
    for cluster, color in CLUSTER_COLORS.items():
        sub = df[df["cluster"] == cluster]
        if len(sub) == 0: continue
        sz = np.clip(np.sqrt(sub["fsa"] / 500), 100, 3000)
        ax.scatter(sub["roi_d"], sub["roi_i"], c=color, s=sz, alpha=0.65,
                   label=f"{cluster} (n={len(sub)})", zorder=3, linewidths=0)

    ax.axvline(1.0,  color="#888888", linewidth=2, linestyle="--", alpha=0.7)
    ax.axhline(13.0, color="#555555", linewidth=2, linestyle="--", alpha=0.7)
    ax.text(1.05, ax.get_ylim()[1] * 0.97 if ax.get_ylim()[1] > 0 else 75,
            "ROI Dom = 1,0x", fontsize=16, color="#888888")
    ax.text(ax.get_xlim()[0] + 0.1 if ax.get_xlim()[0] > -2 else -1.8,
            13.5, "ROI Intl = 13", fontsize=16, color="#555555")
    ax.set_xlabel("ROI Doméstico (receita / FSA)")
    ax.set_ylabel("ROI Internacional médio")
    ax.set_title("Produtoras: retorno doméstico vs. alcance internacional por cluster\n(tamanho = investimento FSA)")
    ax.legend(fontsize=18, markerscale=4, loc="upper right")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 11 — Scatter: investimento vs. receita por produtora
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig11():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    df["rec"] = _num(df, "receita_total_deflac")
    df["n"]   = _num(df, "n_obras")
    df = df[(df["fsa"] > 0) | (df["rec"] > 0)]

    fig, ax = plt.subplots(figsize=(22, 12))
    for cluster, color in CLUSTER_COLORS.items():
        sub = df[df["cluster"] == cluster]
        if len(sub) == 0: continue
        ax.scatter(sub["fsa"] / 1e6, sub["rec"] / 1e6,
                   c=color, s=np.clip(sub["n"] * 80, 150, 3000),
                   alpha=0.65, label=cluster, zorder=3, linewidths=0)

    mx = max(df["fsa"].max(), df["rec"].max()) / 1e6 * 1.05
    ax.plot([0, mx], [0, mx], color="#bbbbbb", linewidth=1, linestyle="--",
            label="Paridade (receita = investimento)", zorder=2)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.set_xlabel("Investimento FSA total (R$ M deflac. 2024)")
    ax.set_ylabel("Receita total estimada (R$ M deflac. 2024)")
    ax.set_title("Relação entre investimento e receita por produtora\n(tamanho = nº de obras, cor = cluster)")
    ax.legend(fontsize=18, markerscale=3)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 12 — Curva de Lorenz: concentração do FSA entre produtoras
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig12():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    df_pos = df[df["fsa"] > 0].sort_values("fsa")

    cumfsa  = df_pos["fsa"].cumsum() / df_pos["fsa"].sum()
    cumprod = np.arange(1, len(df_pos) + 1) / len(df_pos)

    # Gini
    n = len(df_pos)
    gini = 1 - 2 * np.trapz(cumfsa, cumprod)

    fig, ax = plt.subplots(figsize=(18, 12))
    ax.plot(cumprod * 100, cumfsa * 100, color=C_BLUE, linewidth=2,
            label=f"Curva de Lorenz (Gini = {gini:.3f})")
    ax.plot([0, 100], [0, 100], color="#cccccc", linewidth=1, linestyle="--",
            label="Igualdade perfeita")
    ax.fill_between(cumprod * 100, cumfsa * 100, cumprod * 100,
                    alpha=0.10, color=C_BLUE)
    ax.set_xlabel("% acumulado de produtoras (ordenadas por investimento)")
    ax.set_ylabel("% acumulado do investimento FSA")
    ax.set_title("Concentração do FSA entre produtoras do sistema (FSA + renúncia)\n(Curva de Lorenz)")
    ax.legend(fontsize=18)
    ax.set_xlim(0, 100); ax.set_ylim(0, 100)

    # Anotações de percentis
    for pct in [0.50, 0.80, 0.90]:
        idx = int(pct * len(df_pos))
        x_val = pct * 100
        y_val = cumfsa.iloc[idx] * 100
        ax.annotate(f"Top {100-pct*100:.0f}% = {100-y_val:.0f}% do FSA",
                    xy=(x_val, y_val), xytext=(x_val - 30, y_val + 8),
                    arrowprops=dict(arrowstyle="->", color="#888888", lw=0.8),
                    fontsize=16, color="#555555")

    ax.xaxis.set_major_formatter(PercentFormatter())
    ax.yaxis.set_major_formatter(PercentFormatter())
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 23 — Curva acumulada do FSA por produtora
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig23():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    df_pos = df[df["fsa"] > 0].sort_values("fsa", ascending=False).reset_index(drop=True)
    cumfsa = df_pos["fsa"].cumsum() / df_pos["fsa"].sum() * 100
    x = np.arange(1, len(df_pos) + 1)

    fig, ax = plt.subplots(figsize=(20, 9))
    ax.plot(x, cumfsa, color=C_BLUE, linewidth=2)
    ax.fill_between(x, cumfsa, alpha=0.08, color=C_BLUE)

    # Vertical markers
    for n_top, color in [(10, C_RED), (25, C_ORANGE), (50, C_GREEN)]:
        y_val = cumfsa.iloc[n_top - 1]
        ax.axvline(n_top, color=color, linewidth=1, linestyle="--", alpha=0.7)
        ax.text(n_top + 1, y_val - 3, f"Top {n_top}\n= {y_val:.1f}%",
                fontsize=16, color=color)

    ax.set_xlabel("Nº de produtoras (da maior para a menor investidora)")
    ax.set_ylabel("% acumulado do FSA total")
    ax.set_title("Curva acumulada do FSA total por produtora com investimento positivo")
    ax.yaxis.set_major_formatter(PercentFormatter())
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# FIGURA 13 — Participação no FSA e ticket mediano por tier de produtora
# ══════════════════════════════════════════════════════════════════════════════
def chart_fig13():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    df = df[df["fsa"] > 0]
    total_fsa = df["fsa"].sum()
    df_sorted = df.sort_values("fsa", ascending=False).reset_index(drop=True)
    n = len(df_sorted)

    tiers = [
        ("Top 10", df_sorted.head(10)),
        ("11–25",  df_sorted.iloc[10:25]),
        ("26–50",  df_sorted.iloc[25:50]),
        ("51–100", df_sorted.iloc[50:100]),
        ("101+",   df_sorted.iloc[100:]),
    ]

    labels  = [t[0] for t in tiers]
    shares  = [t[1]["fsa"].sum() / total_fsa * 100 for t in tiers]
    tickets = [t[1]["fsa"].median() / 1e6 for t in tiers]  # R$ milhões
    ns      = [len(t[1]) for t in tiers]

    fig, ax1 = plt.subplots(figsize=(20, 9))
    ax2 = ax1.twinx()
    x = np.arange(len(labels))
    w = 0.4
    bars = ax1.bar(x - w/2, shares, w, color=C_BLUE, alpha=0.85, label="Participação no FSA (%)")
    ax2.bar(x + w/2, tickets, w, color=C_ORANGE, alpha=0.85, label="Ticket mediano (R$ M)")

    ax1.set_xticks(x)
    ax1.set_xticklabels([f"{l}\n(n={ns[i]})" for i, l in enumerate(labels)], fontsize=18)
    ax1.set_ylabel("Participação no FSA total (%)")
    ax1.yaxis.set_major_formatter(PercentFormatter())
    ax2.set_ylabel("Ticket mediano por produtora (R$ M deflac.)", color=C_ORANGE)
    ax2.tick_params(axis="y", colors=C_ORANGE)
    ax2.spines["right"].set_visible(True)
    ax2.spines["right"].set_color("#cccccc")
    ax2.grid(False)

    for b, v in zip(bars, shares):
        ax1.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.5,
                 f"{v:.1f}%", ha="center", va="bottom", fontsize=18)

    ax1.set_title("Participação no FSA e ticket mediano anual por tier de produtora")
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=18)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# Mapa: caption parcial → função geradora
# ══════════════════════════════════════════════════════════════════════════════
CHART_MAP = [
    ("Figura 1. ",  chart_fig1),
    ("Figura A1.", chart_figA1),
    ("Figura 2. ",  chart_fig2),
    ("Figura 3. ",  chart_fig3),
    ("Figura 21.", chart_fig21),
    ("Figura 5. ",  chart_fig5),
    ("Figura 6. ",  chart_fig6),
    ("Figura 7. ",  chart_fig7),
    ("Figura 25.", chart_fig25),
    ("Figura 4. ",  chart_fig4),
    ("Figura 8. ",  chart_fig8),
    ("Figura 9. ",  chart_fig9),
    ("Figura 10.", chart_fig10),
    ("Figura 22.", chart_fig22),
    ("Figura 11.", chart_fig11),
    ("Figura 12.", chart_fig12),
    ("Figura 23.", chart_fig23),
    ("Figura 13.", chart_fig13),
]


# ══════════════════════════════════════════════════════════════════════════════
# Substituição no HTML
# ══════════════════════════════════════════════════════════════════════════════
def replace_charts(html: str) -> tuple[str, int]:
    """
    Abordagem baseada em posições:
    1. Localiza todos os <img src="data:image/png;base64,..."> e seus spans.
    2. Para cada img, encontra o próximo <p class="chart-caption"> imediatamente seguinte.
    3. Extrai o texto Figura N da legenda e chama a função correspondente.
    4. Substitui o base64 in-place usando offsets.
    """
    # Encontra todos os imgs com base64 png: (start, end, b64_start, b64_end)
    img_pat = re.compile(r'<img[^>]+src="data:image/png;base64,([A-Za-z0-9+/=]+)"')
    caption_pat = re.compile(
        r'<p[^>]*class="chart-caption"[^>]*>.*?<em>(.*?)</em>',
        re.DOTALL,
    )

    imgs     = list(img_pat.finditer(html))
    captions = list(caption_pat.finditer(html))

    # Para cada img, achar a legenda imediatamente após (a mais próxima depois do img)
    replacements = []  # lista de (b64_start, b64_end, new_b64)

    cap_idx = 0
    for img_m in imgs:
        img_end = img_m.end()
        # Avança captions até encontrar uma que comece após o img
        while cap_idx < len(captions) and captions[cap_idx].start() < img_end:
            cap_idx += 1
        if cap_idx >= len(captions):
            continue
        next_cap = captions[cap_idx]

        # Verifica que não há outro img entre este e a legenda
        cap_start = next_cap.start()
        text_between = html[img_end:cap_start]
        if '<img' in text_between:
            continue  # outra img no meio → não é a legenda desta

        caption_text = next_cap.group(1).strip()
        # Resolve qual função usar
        func = None
        matched_key = None
        for key, fn in CHART_MAP:
            if caption_text.startswith(key):
                func = fn
                matched_key = key
                break
        if func is None:
            continue  # não tem mapeamento

        # Posições do base64 dentro do html
        b64_start = img_m.start(1)
        b64_end   = img_m.end(1)
        replacements.append((b64_start, b64_end, matched_key, func))

    # Executa as substituições de trás para frente (preserva offsets)
    replaced = 0
    # Pré-gera cada gráfico (evita gerar duplicatas)
    cache = {}
    for b64_start, b64_end, key, func in replacements:
        if key not in cache:
            try:
                print(f"  Gerando {key}...", flush=True)
                cache[key] = func()
                print(f"  {key} OK")
            except Exception as e:
                print(f"  ERRO {key}: {e}")
                cache[key] = None

    result = list(html)
    for b64_start, b64_end, key, func in sorted(replacements, reverse=True):
        new_b64 = cache.get(key)
        if new_b64 is None:
            continue
        result[b64_start:b64_end] = list(new_b64)
        replaced += 1

    return "".join(result), replaced


def main():
    print(f"Carregando {HTML_PATH}...")
    with open(HTML_PATH, encoding="utf-8") as f:
        html = f.read()

    print(f"Substituindo gráficos...")
    new_html, n = replace_charts(html)

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(new_html)
    print(f"\n{n} grafico(s) substituido(s) em {HTML_PATH}")

    # Copia para docs/ e output_final/ se existirem
    for copy_path in [
        os.path.join(ROOT, "docs", "analise.html"),
    ]:
        if os.path.exists(copy_path):
            with open(copy_path, "w", encoding="utf-8") as f:
                f.write(new_html)
            print(f"Copia salva em {copy_path}")


if __name__ == "__main__":
    main()
