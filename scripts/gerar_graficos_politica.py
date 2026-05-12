#!/usr/bin/env python3
"""
gerar_graficos_politica.py
==========================
Regenera os gráficos de 'Uma política de fomento baseada em evidências_v6.html'
com estética de publicação científica (Nature-style: fundo branco, tipografia limpa).

Uso:
    python scripts/gerar_graficos_politica.py

Tamanho de fonte:
    savefig.dpi=160, figsize width≈11" → imagem 1760px → exibida em 880px (escala 0.5)
    Todas as fontes no rcParams são 2× o tamanho desejado no HTML renderizado.
"""

import os, re, base64, io, json
import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.ticker import FuncFormatter, PercentFormatter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML_PATH = os.path.join(
    ROOT, "output_final",
    "Uma política de fomento baseada em evidências_v6.html"
)

# ── Nature-style rcParams ──────────────────────────────────────────────────────
# Imagem salva em 1760px (11" × 160dpi), exibida em 880px → escala 0.5
# Fontes: 2× o tamanho desejado para compensar o downscaling
mpl.rcParams.update({
    "font.family":         "sans-serif",
    "font.sans-serif":     ["Arial", "Helvetica", "DejaVu Sans", "sans-serif"],
    "svg.fonttype":        "none",
    "pdf.fonttype":        42,
    "font.size":           26,
    "axes.titlesize":      28,
    "axes.titleweight":    "bold",
    "axes.titlepad":       18,
    "axes.labelsize":      24,
    "axes.labelcolor":     "#333333",
    "axes.spines.right":   False,
    "axes.spines.top":     False,
    "axes.linewidth":      1.5,
    "axes.edgecolor":      "#aaaaaa",
    "axes.facecolor":      "#fafafa",
    "figure.facecolor":    "white",
    "xtick.labelsize":     22,
    "ytick.labelsize":     22,
    "xtick.color":         "#555555",
    "ytick.color":         "#555555",
    "xtick.major.size":    6,
    "ytick.major.size":    6,
    "legend.frameon":      False,
    "legend.fontsize":     22,
    "grid.alpha":          0.45,
    "grid.linewidth":      0.9,
    "grid.color":          "#cccccc",
    "axes.grid":           True,
    "axes.grid.axis":      "y",
    "figure.dpi":          110,
    "savefig.dpi":         160,
    "savefig.facecolor":   "white",
    "savefig.bbox":        "tight",
    "savefig.pad_inches":  0.3,
})

# ── Paleta científica ──────────────────────────────────────────────────────────
C_BLUE   = "#2166ac"
C_ORANGE = "#e08026"
C_GREEN  = "#4dac26"
C_RED    = "#d7191c"
C_PURPLE = "#762a83"
C_TEAL   = "#1b9e77"
C_GREY   = "#888888"
C_AMBER  = "#b8860b"
C_MUTED  = "#999999"

CAT_SHORT = {
    "FSA Pontuação Bilheteria e Roteiro — Distribuidora": "Seletivo\nDistribuidora",
    "FSA Pontuação Bilheteria e Roteiro — Produtora":     "Seletivo\nProdutora",
    "FSA Pontuação Festivais e Roteiro":                  "Seletivo\nFestivais",
    "FSA Automático Bilheteria":                          "Automático\nBilheteria",
    "FSA Automático Festivais":                           "Automático\nFestivais",
    "FSA Comercialização / Distribuição":                 "Comercialização",
    "FSA Complementação":                                 "Complementação",
    "FSA Coprodução Internacional":                       "Coprodução\nIntl",
}

# ── Helpers ────────────────────────────────────────────────────────────────────
def load_csv(name):
    path = os.path.join(ROOT, "resultados", "datasets", name)
    return pd.read_csv(path, encoding="utf-8-sig", sep=None, engine="python")


def _num(df, col):
    return pd.to_numeric(df[col], errors="coerce").fillna(0)


def fig_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()


def fmt_m(v, p):
    if abs(v) >= 1e9: return f"R${v/1e9:.1f}bi"
    if abs(v) >= 1e6: return f"R${v/1e6:.0f}M"
    if abs(v) >= 1e3: return f"R${v/1e3:.0f}K"
    return f"R${v:.0f}"


# ── Datasets ───────────────────────────────────────────────────────────────────
df_prod = load_csv("base_nivel_produtora.csv")
df_inv  = load_csv("base_nivel_investimento.csv")
df_obra = load_csv("base_nivel_obra.csv")

# ── Dados por categoria — hardcoded do objeto CATS do painel (script 03) ──────
# roi_tot_def = sum(receita_def) / sum(FSA_def + renúncia_def) por categoria
_CAT_DATA = {
    "FSA Pontuação Festivais e Roteiro":               {"roi_tot_def": 0.0230, "intl_avg": 3.69,  "intl_total": 282.10, "n_obras": 78,  "inv": 104324502},
    "FSA Pontuação Bilheteria e Roteiro — Produtora":  {"roi_tot_def": 0.5349, "intl_avg": 2.28,  "intl_total": 213.26, "n_obras": 163, "inv": 173848657},
    "FSA Pontuação Bilheteria e Roteiro — Distribuidora": {"roi_tot_def": 0.9419, "intl_avg": 3.41, "intl_total": 394.63, "n_obras": 131, "inv": 222132407},
    "FSA Automático Bilheteria":                       {"roi_tot_def": 0.3206, "intl_avg": 2.37,  "intl_total":  58.65, "n_obras": 52,  "inv": 72508712},
    "FSA Automático Festivais":                        {"roi_tot_def": 0.0607, "intl_avg": 8.17,  "intl_total":  39.71, "n_obras": 10,  "inv": 8325780},
    "FSA Coprodução Internacional":                    {"roi_tot_def": 0.0657, "intl_avg": 2.53,  "intl_total":  75.80, "n_obras": 30,  "inv": 11967067},
    "FSA Comercialização / Distribuição":              {"roi_tot_def": 0.3820, "intl_avg": 3.78,  "intl_total":  62.35, "n_obras": 147, "inv": 37519926},
    "FSA Complementação":                              {"roi_tot_def": 0.4748, "intl_avg": 1.98,  "intl_total":  73.06, "n_obras": 90,  "inv": 65480424},
}
df_cat_painel = pd.DataFrame([
    {"categoria": k, "roi_dom_med": v["roi_tot_def"], "roi_intl": v["intl_avg"],
     "intl_total": v["intl_total"], "n_obras": v["n_obras"], "inv": v["inv"]}
    for k, v in _CAT_DATA.items()
])


# ══════════════════════════════════════════════════════════════════════════════
# CHART 1 — Vocação Comercial vs. Alcance Internacional por Categoria FSA
# Canvas máximo: 22×12" — S=22/11=2 para manter fontes idênticas em 880px
# Colorbar/legenda embaixo (horizontal)
# ══════════════════════════════════════════════════════════════════════════════
def chart1():
    S = 22 / 11  # figsize 22" @ 160dpi = 3520px → 880px display → escala 0.25
    df = df_cat_painel.copy()
    df["label"] = df["categoria"].map(CAT_SHORT).fillna(df["categoria"])

    with mpl.rc_context({
        "font.size":        round(26 * S),
        "axes.titlesize":   round(28 * S),
        "axes.labelsize":   round(24 * S),
        "xtick.labelsize":  round(22 * S),
        "ytick.labelsize":  round(22 * S),
        "legend.fontsize":  round(22 * S),
    }):
        fig, ax = plt.subplots(figsize=(33, 18))
        scatter = ax.scatter(df["roi_dom_med"], df["roi_intl"],
                             s=df["n_obras"] * 12, alpha=0.75,
                             c=df["inv"], cmap="Blues_r",
                             vmin=0, vmax=df["inv"].max(),
                             edgecolors=C_BLUE, linewidths=1.0, zorder=3)

        med_dom  = df["roi_dom_med"].median()
        med_intl = df["roi_intl"].median()
        ax.axvline(med_dom,  color="#aaaaaa", linewidth=1.5, linestyle="--", alpha=0.7)
        ax.axhline(med_intl, color="#aaaaaa", linewidth=1.5, linestyle="--", alpha=0.7)

        for _, row in df.iterrows():
            ax.annotate(row["label"],
                        (row["roi_dom_med"], row["roi_intl"]),
                        textcoords="offset points", xytext=(10, 6),
                        fontsize=round(20 * S), color="#333333")

        # Colorbar vertical à direita — evita colisão com o xlabel
        cbar = fig.colorbar(scatter, ax=ax, orientation="vertical",
                            shrink=0.6, pad=0.03, aspect=30)
        cbar.set_label("Investimento total (R$ deflac.)", fontsize=round(20 * S))
        cbar.ax.tick_params(labelsize=round(18 * S))
        cbar.ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v/1e6:.0f}M"))
        ax.set_xlabel("ROI Doméstico — média pond. (bilheteria / invest. total)")
        ax.set_ylabel("ROI Internacional médio")
        ax.set_title("Vocação Comercial vs. Alcance Internacional por Categoria FSA\n(tamanho = nº de obras  |  cor = volume de investimento)")
        ax.grid(True, alpha=0.3)
        fig.tight_layout()
        return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 2 — ROI Doméstico e Internacional por Categoria FSA (barras horizontais)
# Canvas máximo: 22×10" — S=2 — legenda centralizada embaixo
# ══════════════════════════════════════════════════════════════════════════════
def chart2():
    S = 22 / 11  # = 2
    df = df_cat_painel.copy()
    df["label"] = df["categoria"].map(CAT_SHORT).fillna(df["categoria"])

    def _cor(label):
        if "Distrib" in str(label):   return C_GREEN
        if "Produt" in str(label):    return C_BLUE
        if "Festivais" in str(label): return C_ORANGE
        if "Coprod" in str(label):    return C_PURPLE
        if "Autom" in str(label):     return C_TEAL
        return C_GREY

    df_dom  = df.sort_values("roi_dom_med", ascending=True)
    df_intl = df.sort_values("intl_total",  ascending=True)

    with mpl.rc_context({
        "font.size":        round(26 * S),
        "axes.titlesize":   round(28 * S),
        "axes.labelsize":   round(24 * S),
        "xtick.labelsize":  round(22 * S),
        "ytick.labelsize":  round(22 * S),
        "legend.fontsize":  round(22 * S),
    }):
        fig, axes = plt.subplots(1, 2, figsize=(33, 14))

        # Left: ROI dom
        ax = axes[0]
        colors_dom = [_cor(l) for l in df_dom["label"]]
        bars = ax.barh(df_dom["label"], df_dom["roi_dom_med"],
                       color=colors_dom, alpha=0.88, edgecolor="white")
        ax.set_title("ROI Doméstico\n(bilh. / invest. total)", fontsize=round(22 * S))
        ax.set_xlabel("ROI doméstico", fontsize=round(20 * S))
        ax.tick_params(axis="y", labelsize=round(18 * S))
        ax.axvline(1.0, color="#666", linewidth=1.2, linestyle="--", alpha=0.6)
        ax.grid(axis="x"); ax.grid(axis="y", alpha=0)
        for b in bars:
            v = b.get_width()
            if v > 0.01:
                ax.text(v + 0.01, b.get_y() + b.get_height() / 2,
                        f"{v:.2f}x", va="center", fontsize=round(18 * S), color="#333")

        # Right: ROI intl TOTAL
        ax2 = axes[1]
        colors_intl = [_cor(l) for l in df_intl["label"]]
        bars2 = ax2.barh(df_intl["label"], df_intl["intl_total"],
                         color=colors_intl, alpha=0.88, edgecolor="white")
        ax2.set_title("ROI Internacional\n(total acumulado)", fontsize=round(22 * S))
        ax2.set_xlabel("ROI Internacional total", fontsize=round(20 * S))
        ax2.tick_params(axis="y", labelsize=round(18 * S))
        ax2.grid(axis="x"); ax2.grid(axis="y", alpha=0)
        for b in bars2:
            v = b.get_width()
            if v > 1:
                ax2.text(v + 2, b.get_y() + b.get_height() / 2,
                         f"{v:.0f}", va="center", fontsize=round(18 * S), color="#333")

        handles = [
            mpatches.Patch(color=C_GREEN,  label="Seletivo Distribuidora"),
            mpatches.Patch(color=C_BLUE,   label="Seletivo Produtora"),
            mpatches.Patch(color=C_ORANGE, label="Seletivo Festivais"),
            mpatches.Patch(color=C_PURPLE, label="Coprodução Internacional"),
            mpatches.Patch(color=C_TEAL,   label="Automático"),
            mpatches.Patch(color=C_GREY,   label="Outros"),
        ]
        # Legenda centralizada na base da figura
        fig.legend(handles=handles, fontsize=round(18 * S),
                   loc="upper center", bbox_to_anchor=(0.5, 0.02),
                   ncol=3, frameon=True, facecolor="white", edgecolor="#cccccc")
        fig.suptitle("Desempenho por categoria FSA — Vocação Comercial vs. Alcance Internacional",
                     fontsize=round(22 * S), fontweight="bold")
        fig.tight_layout()
        fig.subplots_adjust(bottom=0.18)  # espaço para a legenda embaixo
        return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 3 — Scatter interativo (Plotly) das produtoras por cluster
# Idêntico ao painel produtoras > aba por cluster > primeira dispersão
# Retorna HTML string (não base64 PNG)
# ══════════════════════════════════════════════════════════════════════════════
def chart3():
    df = df_prod.copy()
    df["inv_mi"]   = _num(df, "investimento_total_deflac") / 1e6
    df["fsa_mi"]   = _num(df, "investimento_fsa_deflac")   / 1e6
    df["rec_mi"]   = _num(df, "receita_total_deflac")      / 1e6
    df["roi_d"]    = _num(df, "roi_dom_total_deflac")
    df["roi_intl"] = _num(df, "roi_intl_medio")
    df["n_obras"]  = _num(df, "n_obras")
    df["nome"]     = df["razao_social"].fillna("").astype(str).str[:40]

    # Filtrar pontos válidos para log scale
    df = df[(df["inv_mi"] > 0) & (df["roi_d"] > 0)].copy()

    CLUSTER_PAL = {
        "Duplo Retorno":         "#f5c842",
        "Retorno Doméstico":     "#5b8cff",
        "Retorno Internacional": "#ff80b0",
        "Fomento Baixo Retorno": "#f87171",
        "Pequeno Porte":         "#7b849a",
    }
    PLOT_ORDER = ["Pequeno Porte", "Fomento Baixo Retorno",
                  "Retorno Internacional", "Retorno Doméstico", "Duplo Retorno"]

    traces = []
    for cluster in PLOT_ORDER:
        sub = df[df["cluster"] == cluster].copy()
        if len(sub) == 0:
            continue
        color = CLUSTER_PAL.get(cluster, "#888888")
        # Tamanho dos marcadores proporcional ao nº de obras (igual ao painel)
        sz = np.clip(sub["n_obras"] * 3, 6, 30).tolist()

        # Customdata para hover: [invest_total, invest_fsa, receita, roi_intl, n_obras]
        customdata = sub[["inv_mi", "fsa_mi", "rec_mi", "roi_intl", "n_obras"]].values.tolist()

        traces.append({
            "type": "scatter",
            "mode": "markers",
            "name": f"{cluster} (n={len(sub)})",
            "x": sub["inv_mi"].tolist(),
            "y": sub["roi_d"].tolist(),
            "text": sub["nome"].tolist(),
            "customdata": customdata,
            "hovertemplate": (
                "<b>%{text}</b><br>"
                "Invest. total: R$ %{customdata[0]:.1f}M<br>"
                "Invest. FSA: R$ %{customdata[1]:.1f}M<br>"
                "Renda doméstica: R$ %{customdata[2]:.1f}M<br>"
                "ROI Internacional: %{customdata[3]:.2f}x<br>"
                "Nº obras: %{customdata[4]:.0f}"
                "<extra></extra>"
            ),
            "marker": {
                "color": color,
                "size": sz,
                "opacity": 0.55 if cluster == "Pequeno Porte" else 0.78,
                "line": {"color": "white", "width": 0.5},
            },
        })

    # Linha ROI = 1
    x_range = [df["inv_mi"].min() * 0.8, df["inv_mi"].max() * 1.2]
    traces.append({
        "type": "scatter",
        "mode": "lines+text",
        "name": "ROI Dom. = 1×",
        "x": x_range,
        "y": [1.0, 1.0],
        "line": {"color": "#888888", "width": 1.5, "dash": "dash"},
        "text": ["", "ROI Dom. = 1×"],
        "textposition": "top right",
        "textfont": {"color": "#666666", "size": 12},
        "hoverinfo": "skip",
        "showlegend": True,
    })

    layout = {
        "title": {
            "text": "Dispersão das produtoras por cluster<br><sup>Investimento acumulado vs. retorno doméstico (mesmo recorte do painel produtoras)</sup>",
            "font": {"size": 16},
        },
        "xaxis": {
            "title": "Investimento total acumulado — R$ mi (escala log)",
            "type": "log",
            "showgrid": True,
            "gridcolor": "#e5e5e5",
            "zeroline": False,
        },
        "yaxis": {
            "title": "ROI doméstico total deflacionado (escala log)",
            "type": "log",
            "showgrid": True,
            "gridcolor": "#e5e5e5",
            "zeroline": False,
        },
        "legend": {
            "orientation": "h",
            "x": 0.5,
            "xanchor": "center",
            "y": -0.14,
            "yanchor": "top",
            "bgcolor": "rgba(255,255,255,0.85)",
            "bordercolor": "#dddddd",
            "borderwidth": 1,
        },
        "paper_bgcolor": "white",
        "plot_bgcolor":  "#f8f9fa",
        "margin": {"l": 70, "r": 30, "t": 80, "b": 130},
        "font": {"family": "Arial, Helvetica, sans-serif", "size": 13},
        "hoverlabel": {
            "bgcolor": "white",
            "bordercolor": "#cccccc",
            "font": {"size": 13},
        },
    }

    traces_json = json.dumps(traces)
    layout_json = json.dumps(layout)

    # Wrapper div com mesmo estilo que article img (bleed + border-radius)
    # Marcadores <!-- chart3-plotly --> para facilitar substituição em runs subsequentes
    div_id = "plotly-chart3-prod"
    html = (
        '<!-- chart3-plotly -->'
        '<div style="display:block;width:calc(100% + 160px);max-width:none;'
        'margin:40px -80px;border-radius:10px;overflow:hidden;background:white">'
        f'<div id="{div_id}" style="width:100%;height:750px"></div>'
        '<script src="https://cdn.plot.ly/plotly-2.35.2.min.js" charset="utf-8"></script>'
        f'<script>Plotly.newPlot("{div_id}",'
        f'{traces_json},'
        f'{layout_json},'
        '{"responsive":true,"displayModeBar":true,"displaylogo":false})'
        '</script>'
        '</div>'
        '<!-- /chart3-plotly -->'
    )
    return html


# ══════════════════════════════════════════════════════════════════════════════
# CHART 4 — Produtoras ativas por ano e universo acumulado
# ══════════════════════════════════════════════════════════════════════════════
def chart4():
    years         = [2014,2015,2016,2017,2018,2019,2020,2021,2022,2023]
    active_counts = [144, 149, 231, 238, 304, 327, 257, 188, 230, 338]
    cum_counts    = [144, 254, 402, 529, 643, 766, 840, 895, 981,1084]

    x = np.arange(len(years))
    fig, ax1 = plt.subplots(figsize=(11, 7))
    ax2 = ax1.twinx()

    bars = ax1.bar(x, active_counts, color=C_BLUE, alpha=0.80, label="Ativas no ano",
                   edgecolor="white", linewidth=0.5, zorder=3)
    line, = ax2.plot(x, cum_counts, color=C_RED, linewidth=2.5,
                     marker="o", markersize=7, label="Universo acumulado", zorder=4)
    ax2.fill_between(x, cum_counts, alpha=0.07, color=C_RED)

    for b, v in zip(bars, active_counts):
        ax1.text(b.get_x() + b.get_width() / 2, b.get_height() + max(active_counts) * 0.015,
                 str(v), ha="center", va="bottom", fontsize=20, color=C_BLUE)

    ax1.set_xticks(x)
    ax1.set_xticklabels(years, fontsize=20)
    ax1.set_ylabel("Produtoras ativas no ano", color=C_BLUE)
    ax1.tick_params(axis="y", labelcolor=C_BLUE)
    ax1.set_ylim(0, max(active_counts) * 1.22)
    ax2.set_ylabel("Universo acumulado (total distinto)", color=C_RED)
    ax2.tick_params(axis="y", labelcolor=C_RED)
    ax2.spines["right"].set_visible(True)
    ax2.spines["right"].set_color("#dddddd")
    ax2.grid(False)

    handles = [mpatches.Patch(color=C_BLUE, label="Produtoras ativas no ano"),
               mpatches.Patch(color=C_RED,  label="Universo acumulado")]
    ax1.legend(handles=handles, fontsize=20, loc="upper left")
    ax1.set_title("Produtoras ativas por ano e universo acumulado\nMais produtoras, mais obras, menos dinheiro por projeto")
    ax1.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 5 — Curva de Lorenz: concentração do FSA
# ══════════════════════════════════════════════════════════════════════════════
def chart5():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    vals = df["fsa"].sort_values().values
    cumvals  = np.cumsum(vals) / vals.sum()
    cumshare = np.arange(1, len(vals) + 1) / len(vals)
    gini = 1 - 2 * np.trapz(cumvals, cumshare)

    fig, ax = plt.subplots(figsize=(9, 9))
    ax.plot([0, 1], [0, 1], color="#cccccc", linewidth=1.5, linestyle="--",
            label="Igualdade perfeita", zorder=2)
    ax.fill_between(cumshare, cumvals, cumshare, alpha=0.10, color=C_BLUE, zorder=2)
    ax.plot(cumshare, cumvals, color=C_BLUE, linewidth=2.5,
            label="Lorenz — FSA por produtora", zorder=3)

    ax.annotate(f"Gini = {gini:.2f}",
                xy=(0.65, 0.25), fontsize=24, color=C_BLUE, fontweight="bold",
                bbox=dict(facecolor="white", edgecolor=C_BLUE, alpha=0.9,
                          boxstyle="round,pad=0.5"))

    for pct, label_offset in [(0.50, (-0.15, 0.07)), (0.80, (-0.20, 0.06))]:
        idx = int(pct * len(vals))
        x_v, y_v = pct, cumvals[idx]
        ax.annotate(f"Top {100-pct*100:.0f}% = {100*(1-y_v):.0f}% do FSA",
                    xy=(x_v, y_v), xytext=(x_v + label_offset[0], y_v + label_offset[1]),
                    arrowprops=dict(arrowstyle="->", color="#888888", lw=1.0),
                    fontsize=20, color="#555555")

    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ax.xaxis.set_major_formatter(PercentFormatter(1))
    ax.yaxis.set_major_formatter(PercentFormatter(1))
    ax.set_xlabel("Fração de produtoras (% acumulado)")
    ax.set_ylabel("Fração do investimento FSA (% acumulado)")
    ax.set_title("Curva de Lorenz — distribuição do investimento FSA por produtora")
    ax.legend(fontsize=20)
    ax.grid(True, alpha=0.4)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 6 — Distribuição do ticket anual combinado
# ══════════════════════════════════════════════════════════════════════════════
def chart6():
    labels = ["0–100k","100–200k","200–300k","300–500k",
              "500–750k","750k–1M","1–1,5M","1,5–2M",
              "2–3M","3–5M","5–10M","10M+"]
    counts = [371, 139, 79, 96, 54, 41, 43, 25, 17, 19, 7, 3]

    colors = ([C_RED]    * 4 +
              [C_AMBER]  * 2 +
              [C_TEAL]   * 3 +
              [C_BLUE]   * 3)

    fig, ax = plt.subplots(figsize=(11, 7))
    bars = ax.bar(labels, counts, color=colors, alpha=0.88, edgecolor="white", zorder=3)

    ax.axvline(3.5, color=C_RED,   linewidth=1.5, linestyle="--", alpha=0.7)
    ax.axvline(5.5, color=C_AMBER, linewidth=1.5, linestyle="--", alpha=0.7)
    max_count = max(counts)
    ax.text(3.6, max_count * 0.97, "R$500k", fontsize=20, color=C_RED,   va="top")
    ax.text(5.6, max_count * 0.97, "R$1M",   fontsize=20, color=C_AMBER, va="top")

    for b, v in zip(bars, counts):
        if v > 0:
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + max_count * 0.01,
                    str(v), ha="center", va="bottom", fontsize=20)

    ax.set_xlabel("Ticket anual combinado por produtora (FSA anualizado + Proxy RLP)")
    ax.set_ylabel("Nº de produtoras")
    ax.set_title("Distribuição do ticket anual combinado (fomento + Proxy RLP)\nProdutoras com FSA positivo no período 2014–2023")
    ax.tick_params(axis="x", labelsize=16, rotation=22)
    handles = [mpatches.Patch(color=C_RED,   label="Crítico (< R$500k)"),
               mpatches.Patch(color=C_AMBER, label="Atenção (R$500k–1M)"),
               mpatches.Patch(color=C_TEAL,  label="Viável (R$1–3M)"),
               mpatches.Patch(color=C_BLUE,  label="Bom (> R$3M)")]
    ax.legend(handles=handles, fontsize=20)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 7 — Raça: % entre Inscritos vs. Selecionados
# ══════════════════════════════════════════════════════════════════════════════
def chart7():
    stages       = ["Inscritos\nsem PA", "Inscritos\ncom PA", "Selecionados\ncom PA"]
    pct_neg_vals = [15.2, 22.8, 32.4]
    pct_bra_vals = [56.0, 48.0, 39.0]
    pct_out_vals = [100 - n - b for n, b in zip(pct_neg_vals, pct_bra_vals)]

    x = np.arange(len(stages))
    w = 0.28
    fig, ax = plt.subplots(figsize=(11, 7))

    b1 = ax.bar(x - w, pct_neg_vals, w, color=C_AMBER,  alpha=0.88,
                label="Negro/Pardo/Indígena", edgecolor="white")
    b2 = ax.bar(x,     pct_bra_vals, w, color=C_BLUE,   alpha=0.70,
                label="Branco", edgecolor="white")
    b3 = ax.bar(x + w, pct_out_vals, w, color=C_GREY,   alpha=0.60,
                label="Outros / Não informado", edgecolor="white")

    for bars, vals in [(b1, pct_neg_vals), (b2, pct_bra_vals), (b3, pct_out_vals)]:
        for b, v in zip(bars, vals):
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.5,
                    f"{v:.1f}%", ha="center", va="bottom", fontsize=22, fontweight="bold",
                    color="#222222")

    ax.annotate("", xy=(x[2] - w + w/2, pct_neg_vals[2] + 3),
                xytext=(x[0] - w + w/2, pct_neg_vals[0] + 3),
                arrowprops=dict(arrowstyle="->", color=C_AMBER, lw=2.5))
    ax.text(x[1] - w/2, max(pct_neg_vals) + 5.5, "+17,2pp",
            ha="center", fontsize=24, color=C_AMBER, fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(stages, fontsize=22)
    ax.set_ylabel("% do total de diretores declarantes")
    ax.yaxis.set_major_formatter(PercentFormatter())
    ax.set_ylim(0, 75)
    ax.set_title("Raça — % entre Inscritos vs. Selecionados (Política Afirmativa)\n"
                 "99 editais FSA/BRDE 2015–2025 · dados declaratórios · direção")
    ax.legend(fontsize=20)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHARTS list — chart3 retorna HTML (não b64)
# ══════════════════════════════════════════════════════════════════════════════
CHARTS = [
    ("chart1", chart1, "png"),
    ("chart2", chart2, "png"),
    ("chart3", chart3, "html"),  # Plotly interativo
    ("chart4", chart4, "png"),
    ("chart5", chart5, "png"),
    ("chart6", chart6, "png"),
    ("chart7", chart7, "png"),
]


def main():
    print(f"Carregando {HTML_PATH}...")
    with open(HTML_PATH, encoding="utf-8") as f:
        html = f.read()

    # Gera todos os charts
    print("Gerando graficos...")
    results = []
    for name, fn, kind in CHARTS:
        print(f"  {name}...", flush=True)
        try:
            r = fn()
            results.append((kind, r))
            sz = len(r) // 1000
            print(f"  {name} OK ({sz}KB, {kind})")
        except Exception as e:
            print(f"  {name} ERRO: {e}")
            import traceback; traceback.print_exc()
            results.append((kind, None))

    # Detecta se chart3 já está como Plotly (marcadores <!-- chart3-plotly -->)
    PLOTLY_MARKER_PAT = re.compile(
        r'<!-- chart3-plotly -->.*?<!-- /chart3-plotly -->', re.DOTALL
    )
    plotly3_match = PLOTLY_MARKER_PAT.search(html)
    has_plotly3   = plotly3_match is not None
    print(f"chart3 já é Plotly no HTML: {has_plotly3}")

    # Encontra todas as imgs base64 no HTML (apenas PNG charts)
    B64_PATTERN = re.compile(r'<img([^>]*?)src="(data:image/png;base64,[A-Za-z0-9+/=]+)"([^>]*?)>')
    b64_matches = list(B64_PATTERN.finditer(html))
    print(f"Imagens base64 encontradas: {len(b64_matches)}")

    # Monta lista de substituições (start, end, new_content)
    # chart3 (html): substitui o bloco Plotly existente ou o <img> original
    # demais charts (png): substituem apenas o src base64, em ordem
    subs  = []
    b64_i = 0  # ponteiro para b64_matches (pula chart3 se já for Plotly)
    replaced_png = replaced_html = 0

    for chart_i, (kind, result) in enumerate(results):
        if result is None:
            if not (chart_i == 2 and has_plotly3):
                b64_i += 1
            continue

        if chart_i == 2 and kind == "html":
            if has_plotly3:
                # Substitui o bloco Plotly existente
                subs.append((plotly3_match.start(), plotly3_match.end(), result))
            else:
                # Substitui o <img> original (primeira vez)
                if b64_i < len(b64_matches):
                    m = b64_matches[b64_i]
                    subs.append((m.start(), m.end(), result))
                    b64_i += 1
            replaced_html += 1
        else:
            if b64_i < len(b64_matches):
                m = b64_matches[b64_i]
                subs.append((m.start(2), m.end(2), "data:image/png;base64," + result))
                b64_i += 1
                replaced_png += 1

    # Aplica do fim para o início para preservar offsets
    html_chars = list(html)
    for start, end, new in sorted(subs, key=lambda x: x[0], reverse=True):
        html_chars[start:end] = list(new)

    html_out = "".join(html_chars)
    print(f"{replaced_png} gráfico(s) PNG substituído(s)")
    print(f"{replaced_html} gráfico(s) HTML (Plotly) substituído(s)")

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html_out)
    print(f"Salvo: {HTML_PATH}")

    # Copia para docs/politica.html
    docs_path = os.path.join(ROOT, "docs", "politica.html")
    if os.path.exists(docs_path):
        import shutil
        shutil.copy2(HTML_PATH, docs_path)
        print(f"Copia salva em {docs_path}")


if __name__ == "__main__":
    main()
