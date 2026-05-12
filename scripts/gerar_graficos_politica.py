#!/usr/bin/env python3
"""
gerar_graficos_politica.py
==========================
Regenera os gráficos de 'Uma política de fomento baseada em evidências_v6.html'
com estética de publicação científica (Nature-style: fundo branco, tipografia limpa).

Mesma lógica de gerar_graficos_analise.py, mas para o artigo de opinião.

Uso:
    python scripts/gerar_graficos_politica.py
"""

import os, re, base64, io
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

# ── Nature-style rcParams (idêntico ao gerar_graficos_analise.py) ──────────────
mpl.rcParams.update({
    "font.family":         "sans-serif",
    "font.sans-serif":     ["Arial", "Helvetica", "DejaVu Sans", "sans-serif"],
    "svg.fonttype":        "none",
    "pdf.fonttype":        42,
    "font.size":           10,
    "axes.titlesize":      11,
    "axes.titleweight":    "bold",
    "axes.titlepad":       10,
    "axes.labelsize":      10,
    "axes.labelcolor":     "#333333",
    "axes.spines.right":   False,
    "axes.spines.top":     False,
    "axes.linewidth":      0.8,
    "axes.edgecolor":      "#aaaaaa",
    "axes.facecolor":      "#fafafa",
    "figure.facecolor":    "white",
    "xtick.labelsize":     9,
    "ytick.labelsize":     9,
    "xtick.color":         "#555555",
    "ytick.color":         "#555555",
    "xtick.major.size":    3,
    "ytick.major.size":    3,
    "legend.frameon":      False,
    "legend.fontsize":     9,
    "grid.alpha":          0.45,
    "grid.linewidth":      0.5,
    "grid.color":          "#cccccc",
    "axes.grid":           True,
    "axes.grid.axis":      "y",
    "figure.dpi":          130,
    "savefig.dpi":         150,
    "savefig.facecolor":   "white",
    "savefig.bbox":        "tight",
    "savefig.pad_inches":  0.15,
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
df_cham = load_csv("base_nivel_chamada.csv")
df_prod = load_csv("base_nivel_produtora.csv")
df_inv  = load_csv("base_nivel_investimento.csv")


# ══════════════════════════════════════════════════════════════════════════════
# CHART 1 — Desempenho por tipo de chamada: ROI dom vs. intl
# ══════════════════════════════════════════════════════════════════════════════
def chart1():
    agg = (df_cham[df_cham["categoria"].isin(CAT_SHORT)]
           .groupby("categoria", as_index=False)
           .agg(roi_dom=("roi_dom_fsa_agregado", "mean"),
                roi_intl=("roi_intl_medio", "mean"),
                n=("n_obras", "sum"))
           .assign(label=lambda x: x["categoria"].map(CAT_SHORT))
           .sort_values("roi_dom", ascending=False))

    fig, ax = plt.subplots(figsize=(13, 5))
    x = np.arange(len(agg))
    w = 0.36
    b1 = ax.bar(x - w/2, agg["roi_dom"],  w, color=C_BLUE,   alpha=0.85,
                label="ROI Doméstico / FSA", zorder=3)
    b2 = ax.bar(x + w/2, agg["roi_intl"], w, color=C_ORANGE, alpha=0.85,
                label="ROI Internacional médio", zorder=3)
    ax.axhline(1.0, color="#666666", linewidth=1, linestyle="--", alpha=0.5,
               label="Paridade (ROI = 1)")

    for bar, col in [(b1, C_BLUE), (b2, C_ORANGE)]:
        for b in bar:
            h = b.get_height()
            if h > 0.05:
                ax.text(b.get_x() + b.get_width() / 2, h + 0.03,
                        f"{h:.2f}", ha="center", va="bottom", fontsize=8.5, color=col)

    ax.set_xticks(x)
    ax.set_xticklabels(agg["label"], fontsize=9)
    ax.set_ylabel("ROI médio")
    ax.set_title("Desempenho por tipo de chamada: retorno doméstico vs. internacional")
    ax.legend(fontsize=9, loc="upper right")
    ax.set_ylim(0, max(agg["roi_intl"].max(), agg["roi_dom"].max()) * 1.22)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 2 — Perfil de retorno: ROI dom + ROI intl por categoria
# ══════════════════════════════════════════════════════════════════════════════
def chart2():
    cats_cmp = {
        "FSA Pontuação Bilheteria e Roteiro — Distribuidora": "Seletivo\nDistribuidora",
        "FSA Pontuação Bilheteria e Roteiro — Produtora":     "Seletivo\nProdutora",
        "FSA Automático Bilheteria":                          "Automático\nBilheteria",
        "FSA Pontuação Festivais e Roteiro":                  "Seletivo\nFestivais",
        "FSA Automático Festivais":                           "Automático\nFestivais",
        "FSA Coprodução Internacional":                       "Coprodução\nIntl",
    }
    sub = (df_cham[df_cham["categoria"].isin(cats_cmp)]
           .groupby("categoria", as_index=False)
           .agg(dom=("roi_dom_fsa_agregado", "mean"),
                intl=("roi_intl_medio", "mean"),
                n=("n_obras", "sum"))
           .assign(label=lambda x: x["categoria"].map(cats_cmp))
           .sort_values("dom", ascending=True))

    fig, axes = plt.subplots(1, 2, figsize=(13, 5.5))

    ax = axes[0]
    colors = [C_GREEN if "Distrib" in l else C_BLUE if "Produt" in l else C_MUTED
              for l in sub["label"]]
    bars = ax.barh(sub["label"], sub["dom"], color=colors, alpha=0.88, edgecolor="white")
    ax.set_title("ROI Doméstico médio (receita / FSA)")
    ax.set_xlabel("ROI")
    ax.axvline(1.0, color="#666", linewidth=0.8, linestyle="--", alpha=0.6)
    ax.grid(axis="x"); ax.grid(axis="y", alpha=0)
    for b in bars:
        v = b.get_width()
        if v > 0.02:
            ax.text(v + 0.01, b.get_y() + b.get_height() / 2,
                    f"{v:.2f}x", va="center", fontsize=9, color="#333")

    ax2 = axes[1]
    bars2 = ax2.barh(sub["label"], sub["intl"], color=C_ORANGE, alpha=0.88, edgecolor="white")
    ax2.set_title("ROI Internacional médio")
    ax2.set_xlabel("ROI Internacional")
    ax2.grid(axis="x"); ax2.grid(axis="y", alpha=0)
    for b in bars2:
        v = b.get_width()
        if v > 0.1:
            ax2.text(v + 0.05, b.get_y() + b.get_height() / 2,
                     f"{v:.2f}", va="center", fontsize=9, color="#333")

    fig.suptitle("Perfil de retorno por tipo de chamada: doméstico vs. internacional",
                 fontsize=11, fontweight="bold")
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 3 — Scatter: produtoras por cluster (investimento vs. receita)
# ══════════════════════════════════════════════════════════════════════════════
def chart3():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    df["rec"] = _num(df, "receita_total_deflac")
    df["n"]   = _num(df, "n_obras")

    fig, ax = plt.subplots(figsize=(13, 6.5))
    for cluster, color in CLUSTER_COLORS.items():
        sub = df[df["cluster"] == cluster]
        if len(sub) == 0: continue
        sz = np.where(sub["n"] > 10, 80, np.where(sub["n"] > 5, 50, 28))
        ax.scatter(sub["fsa"] / 1e6, sub["rec"] / 1e6, c=color, s=sz,
                   alpha=0.65, label=f"{cluster} (n={len(sub)})",
                   zorder=4, linewidths=0)

    mx = max(df["fsa"].max(), df["rec"].max()) / 1e6 * 1.05
    ax.plot([0, mx], [0, mx], color="#bbbbbb", linewidth=1, linestyle="--",
            label="Paridade (receita = investimento)", zorder=2)

    # Destacar notáveis
    notaveis = {"Filmes de Plástico": C_AMBER, "Globo": C_RED, "Downtown": C_PURPLE}
    for _, row in df.iterrows():
        for nome, cor in notaveis.items():
            if nome.lower() in str(row.get("razao_social", "")).lower():
                ax.scatter(row["fsa"] / 1e6, row["rec"] / 1e6, c=cor, s=100,
                           zorder=6, marker="*", linewidths=0)
                ax.annotate(str(row["razao_social"])[:22],
                            (row["fsa"] / 1e6, row["rec"] / 1e6),
                            textcoords="offset points", xytext=(8, 4),
                            fontsize=8.5, color=cor, zorder=7)

    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.set_xlabel("Investimento FSA total (R$ M deflac. 2024)")
    ax.set_ylabel("Receita total estimada (R$ M deflac. 2024)")
    ax.set_title("Distribuição das produtoras por retorno e investimento")
    ax.legend(fontsize=8.5, markerscale=1.5, loc="upper left")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 4 — Produtoras por faixa de investimento FSA
# ══════════════════════════════════════════════════════════════════════════════
def chart4():
    df = df_prod.copy()
    df["fsa"] = _num(df, "investimento_fsa_deflac")
    df = df[df["fsa"] > 0]

    bins   = [0, 500_000, 1_000_000, 2_500_000, 5_000_000, 10_000_000, 25_000_000]
    labels = ["< 500K", "500K–1M", "1–2,5M", "2,5–5M", "5–10M", "> 10M"]
    df["tier"] = pd.cut(df["fsa"], bins=bins, labels=labels, right=True)

    counts    = df["tier"].value_counts().reindex(labels).fillna(0)
    total_inv = df.groupby("tier")["fsa"].sum().reindex(labels).fillna(0)

    fig, axes = plt.subplots(1, 2, figsize=(13, 5.5))

    ax = axes[0]
    bar_colors = [C_BLUE if l in ("< 500K", "500K–1M", "1–2,5M") else C_ORANGE for l in labels]
    bars = ax.bar(labels, counts, color=bar_colors, alpha=0.88, edgecolor="white")
    ax.set_title("Nº de produtoras por faixa de investimento FSA")
    ax.set_ylabel("Nº de produtoras")
    ax.tick_params(axis="x", labelsize=9)
    for b in bars:
        h = b.get_height()
        ax.text(b.get_x() + b.get_width() / 2, h + max(counts) * 0.02,
                f"{int(h)}", ha="center", va="bottom", fontsize=9)

    ax2 = axes[1]
    bars2 = ax2.bar(labels, total_inv / 1e6, color=bar_colors, alpha=0.88, edgecolor="white")
    ax2.set_title("Volume FSA por faixa (R$ M deflac. 2024)")
    ax2.set_ylabel("R$ milhões (deflac. 2024)")
    ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax2.tick_params(axis="x", labelsize=9)
    for b, v in zip(bars2, total_inv / 1e6):
        ax2.text(b.get_x() + b.get_width() / 2, b.get_height() + total_inv.max() / 1e6 * 0.02,
                 f"R${v:.0f}M", ha="center", va="bottom", fontsize=9)

    handles = [mpatches.Patch(color=C_BLUE, label="Pequeno/médio porte"),
               mpatches.Patch(color=C_ORANGE, label="Grande porte")]
    axes[0].legend(handles=handles, fontsize=9)
    fig.suptitle("Distribuição do investimento FSA por porte de produtora",
                 fontsize=11, fontweight="bold")
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

    fig, ax = plt.subplots(figsize=(9, 7))
    ax.plot([0, 1], [0, 1], color="#cccccc", linewidth=1.2, linestyle="--",
            label="Igualdade perfeita", zorder=2)
    ax.fill_between(cumshare, cumvals, cumshare, alpha=0.10, color=C_BLUE, zorder=2)
    ax.plot(cumshare, cumvals, color=C_BLUE, linewidth=2.5,
            label=f"Lorenz — FSA por produtora", zorder=3)

    ax.annotate(f"Gini = {gini:.2f}",
                xy=(0.65, 0.25), fontsize=15, color=C_BLUE, fontweight="bold",
                bbox=dict(facecolor="white", edgecolor=C_BLUE, alpha=0.9,
                          boxstyle="round,pad=0.5"))

    # Percentil annotations
    for pct, label_offset in [(0.50, (-0.15, 0.07)), (0.80, (-0.20, 0.06))]:
        idx = int(pct * len(vals))
        x_v, y_v = pct, cumvals[idx]
        ax.annotate(f"Top {100-pct*100:.0f}% = {100*(1-y_v):.0f}% do FSA",
                    xy=(x_v, y_v), xytext=(x_v + label_offset[0], y_v + label_offset[1]),
                    arrowprops=dict(arrowstyle="->", color="#888888", lw=0.8),
                    fontsize=9, color="#555555")

    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ax.xaxis.set_major_formatter(PercentFormatter(1))
    ax.yaxis.set_major_formatter(PercentFormatter(1))
    ax.set_xlabel("Fração de produtoras (% acumulado)")
    ax.set_ylabel("Fração do investimento FSA (% acumulado)")
    ax.set_title("Curva de Lorenz — distribuição do investimento FSA por produtora")
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.4)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 6 — Relação investimento FSA vs bilheteria por obra
# ══════════════════════════════════════════════════════════════════════════════
def chart6():
    df = df_inv.copy()
    df = df[df["categoria"].str.startswith("FSA", na=False)].copy()
    df["bil"] = _num(df, "bilheteria_deflac_r2024")

    # Compute ticket (FSA per obra) from aggregates
    agg = (df_cham.groupby("categoria")
           .apply(lambda g: pd.to_numeric(g["investimento_fsa_deflac"], errors="coerce").fillna(0).sum()
                  / max(g["n_obras"].sum(), 1),
                  include_groups=False)
           .rename("ticket"))
    df = df.merge(agg, left_on="categoria", right_index=True, how="left")
    df["ticket"] = _num(df, "ticket")

    sub = df[(df["ticket"] > 0) & (df["bil"] > 0)].copy()

    cat_color = {}
    for cat in sub["categoria"].unique():
        if "Distribuidora" in cat:   cat_color[cat] = C_GREEN
        elif "Festivais" in cat:     cat_color[cat] = C_ORANGE
        elif "Automático" in cat:    cat_color[cat] = C_MUTED
        elif "Coprodução" in cat:    cat_color[cat] = C_PURPLE
        else:                        cat_color[cat] = C_BLUE

    fig, ax = plt.subplots(figsize=(13, 6.5))
    for cat in sub["categoria"].unique():
        s = sub[sub["categoria"] == cat]
        ax.scatter(s["ticket"] / 1e6, s["bil"] / 1e6, c=cat_color.get(cat, C_BLUE),
                   s=30, alpha=0.55, zorder=3, linewidths=0)

    mx = max(sub["ticket"].max(), sub["bil"].max()) / 1e6 * 1.1
    ax.plot([0, mx], [0, mx], color="#bbbbbb", linewidth=1, linestyle="--",
            label="Paridade (ROI = 1)", zorder=2)

    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"R${v:.0f}M"))
    ax.set_xlabel("Investimento FSA médio por obra, por categoria (R$ M deflac. 2024)")
    ax.set_ylabel("Bilheteria (R$ M deflac. 2024)")
    ax.set_title("Relação entre investimento FSA e bilheteria por obra")

    handles = [
        mpatches.Patch(color=C_GREEN,  label="Seletivo Distribuidora"),
        mpatches.Patch(color=C_ORANGE, label="Seletivo Festivais"),
        mpatches.Patch(color=C_BLUE,   label="Seletivo Produtora"),
        mpatches.Patch(color=C_MUTED,  label="Automático"),
        mpatches.Patch(color=C_PURPLE, label="Coprodução Intl"),
        mpatches.Patch(color="none",   label=""),
    ]
    ax.legend(handles=handles[:-1], fontsize=9, loc="upper left")
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    return fig_b64(fig)


# ══════════════════════════════════════════════════════════════════════════════
# CHART 7 — Diversidade: % diretoras e % regional por categoria
# ══════════════════════════════════════════════════════════════════════════════
def chart7():
    df = df_cham[df_cham["categoria"].isin(CAT_SHORT)].copy()
    df["pct_genero_feminino"]    = _num(df, "pct_genero_feminino")
    df["pct_diversidade_regional"] = _num(df, "pct_diversidade_regional")

    agg = (df.groupby("categoria", as_index=False)
             .agg(fem=("pct_genero_feminino", "mean"),
                  div=("pct_diversidade_regional", "mean"),
                  n=("n_obras", "sum"))
             .assign(label=lambda x: x["categoria"].map(CAT_SHORT))
             .sort_values("fem", ascending=True))

    fig, axes = plt.subplots(1, 2, figsize=(13, 5.5))

    for ax, col, title, color in [
        (axes[0], "fem", "% Obras com diretora mulher", C_PURPLE),
        (axes[1], "div", "% Obras de regiões fora do eixo SP-RJ", C_AMBER),
    ]:
        bars = ax.barh(agg["label"], agg[col] * 100, color=color, alpha=0.88, edgecolor="white")
        ax.set_title(title)
        ax.set_xlabel("%")
        ax.xaxis.set_major_formatter(PercentFormatter())
        ax.grid(axis="x"); ax.grid(axis="y", alpha=0)
        for b in bars:
            v = b.get_width()
            if v > 0.5:
                ax.text(v + 0.5, b.get_y() + b.get_height() / 2,
                        f"{v:.1f}%", va="center", fontsize=9, color="#333")

    fig.suptitle("Diversidade de gênero e regional por tipo de chamada FSA",
                 fontsize=11, fontweight="bold")
    fig.tight_layout()
    return fig_b64(fig)


CHARTS = [
    ("chart1", chart1),
    ("chart2", chart2),
    ("chart3", chart3),
    ("chart4", chart4),
    ("chart5", chart5),
    ("chart6", chart6),
    ("chart7", chart7),
]


def main():
    print(f"Carregando {HTML_PATH}...")
    with open(HTML_PATH, encoding="utf-8") as f:
        html = f.read()

    # Gera todos os charts
    print("Gerando graficos...")
    b64_list = []
    for name, fn in CHARTS:
        print(f"  {name}...", flush=True)
        try:
            b64 = fn()
            b64_list.append(b64)
            print(f"  {name} OK ({len(b64)//1000}KB)")
        except Exception as e:
            print(f"  {name} ERRO: {e}")
            import traceback; traceback.print_exc()
            b64_list.append(None)

    # Encontra todas as imgs base64 no HTML (qualquer formato de <img>)
    # Abordagem: substitui apenas o base64 in-place para preservar atributos
    B64_PATTERN = re.compile(r'(data:image/png;base64,)([A-Za-z0-9+/=]+)')
    b64_matches = list(B64_PATTERN.finditer(html))
    print(f"\nEncontradas {len(b64_matches)} imagens no HTML")

    n_replace = min(len(b64_matches), len(b64_list))
    # Aplica do fim para o início (preserva offsets)
    result = list(html)
    replaced = 0
    for i in reversed(range(n_replace)):
        new_b64 = b64_list[i]
        if new_b64 is None:
            continue
        m = b64_matches[i]
        result[m.start(2):m.end(2)] = list(new_b64)
        replaced += 1

    html = "".join(result)
    print(f"{replaced} grafico(s) substituido(s)")

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Salvo: {HTML_PATH}")

    # Copia para docs/politica.html
    docs_path = os.path.join(ROOT, "docs", "politica.html")
    if os.path.exists(docs_path):
        import shutil
        shutil.copy2(HTML_PATH, docs_path)
        print(f"Copia salva em {docs_path}")


if __name__ == "__main__":
    main()
