#!/usr/bin/env python3
"""
Regenera os 7 gráficos do artigo de opinião com tema escuro, tamanhos maiores
e fontes legíveis. Também corrige fonte, article-meta e link G1.
"""

import os, re, base64, io
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.ticker import FuncFormatter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML_PATH = os.path.join(ROOT, "output_final",
                         "Uma política de fomento baseada em evidências_v6.html")

# ── Paleta (idêntica ao CSS do artigo) ───────────────────────────────────────
BG      = "#141416"
SURF    = "#1e1e22"
SURF2   = "#252529"
TEXT    = "#d8d4ce"
TEXT2   = "#b8b4ae"
MUTED   = "#908a80"
DIM     = "#6a655c"
ACCENT  = "#6c7bf7"
GOLD    = "#c89a6a"
CORAL   = "#e05c5c"
GREEN   = "#3fc97f"
PURPLE  = "#a97cf7"
RULE    = "#ffffff13"

CLUSTER_COLORS = {
    "Duplo Retorno":        GREEN,
    "Retorno Doméstico":    ACCENT,
    "Retorno Internacional":GOLD,
    "Fomento Baixo Retorno":CORAL,
    "Pequeno Porte":        DIM,
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

# ── Helpers ───────────────────────────────────────────────────────────────────
def style_ax(ax, xlabel=None, ylabel=None, grid_axis='y'):
    ax.set_facecolor(SURF)
    ax.tick_params(colors=TEXT, labelsize=13, length=4)
    for side in ('top', 'right'):
        ax.spines[side].set_visible(False)
    for side in ('left', 'bottom'):
        ax.spines[side].set_color(RULE)
    ax.grid(axis=grid_axis, color=RULE, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    if xlabel:
        ax.set_xlabel(xlabel, color=TEXT2, fontsize=13, labelpad=10)
    if ylabel:
        ax.set_ylabel(ylabel, color=TEXT2, fontsize=13, labelpad=10)

def make_fig(w=14, h=6.5):
    fig = plt.figure(figsize=(w, h), facecolor=BG)
    return fig

def fig_b64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=130, bbox_inches='tight',
                facecolor=BG, edgecolor='none')
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()

def fmt_m(v, p):
    if v >= 1_000_000:
        return f"R${v/1_000_000:.1f}M"
    if v >= 1_000:
        return f"R${v/1_000:.0f}K"
    return f"R${v:.0f}"

# ── Load datasets ─────────────────────────────────────────────────────────────
def load_csv(name):
    path = os.path.join(ROOT, "resultados", "datasets", name)
    df = pd.read_csv(path, encoding='utf-8-sig', sep=None, engine='python')
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='ignore')
    return df

df_cham = load_csv("base_nivel_chamada.csv")
df_prod = load_csv("base_nivel_produtora.csv")
df_inv  = load_csv("base_nivel_investimento.csv")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 1 — Resultado por tipo de chamada (dom vs intl)
# ─────────────────────────────────────────────────────────────────────────────
def chart1():
    agg = (df_cham[df_cham['categoria'].isin(CAT_SHORT)]
           .groupby('categoria', as_index=False)
           .agg(roi_dom=('roi_dom_fsa_agregado', 'mean'),
                roi_intl=('roi_intl_medio', 'mean'),
                n=('n_obras', 'sum'))
           .assign(label=lambda x: x['categoria'].map(CAT_SHORT))
           .sort_values('roi_dom', ascending=False))

    fig = make_fig(14, 6.5)
    ax = fig.add_subplot(111)
    style_ax(ax, ylabel='ROI médio')

    x = np.arange(len(agg))
    w = 0.36
    b1 = ax.bar(x - w/2, agg['roi_dom'], w, color=ACCENT, alpha=0.92, zorder=3,
                label='ROI Doméstico / FSA')
    b2 = ax.bar(x + w/2, agg['roi_intl'], w, color=GOLD,   alpha=0.92, zorder=3,
                label='ROI Internacional (médio)')

    for bar, col in [(b1, ACCENT), (b2, GOLD)]:
        for b in bar:
            h = b.get_height()
            if h > 0.05:
                ax.text(b.get_x() + b.get_width()/2, h + 0.04,
                        f'{h:.2f}', ha='center', va='bottom',
                        fontsize=10, color=col, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(agg['label'], fontsize=12, color=TEXT, linespacing=1.3)
    ax.axhline(1.0, color=TEXT2, linewidth=1, linestyle='--', alpha=0.35, zorder=2)
    ax.set_title('Desempenho por tipo de chamada: retorno doméstico vs. internacional',
                 color=TEXT, fontsize=15, fontweight='bold', pad=14)
    ax.legend(fontsize=13, facecolor=SURF2, edgecolor=RULE, labelcolor=TEXT,
              framealpha=0.95, loc='upper right')
    ax.set_ylim(0, max(agg['roi_intl'].max(), agg['roi_dom'].max()) * 1.25)
    fig.tight_layout(pad=1.8)
    return fig_b64(fig)

# ─────────────────────────────────────────────────────────────────────────────
# CHART 2 — Distribuidora vs Produtora (4 métricas)
# ─────────────────────────────────────────────────────────────────────────────
def chart2():
    cats_cmp = {
        "FSA Pontuação Bilheteria e Roteiro — Distribuidora": "Seletivo\nDistribuidora",
        "FSA Pontuação Bilheteria e Roteiro — Produtora":     "Seletivo\nProdutora",
        "FSA Automático Bilheteria":                          "Automático\nBilheteria",
        "FSA Pontuação Festivais e Roteiro":                  "Seletivo\nFestivais",
        "FSA Automático Festivais":                           "Automático\nFestivais",
        "FSA Coprodução Internacional":                       "Coprodução\nIntl",
    }
    sub = (df_cham[df_cham['categoria'].isin(cats_cmp)]
           .groupby('categoria', as_index=False)
           .agg(dom=('roi_dom_fsa_agregado', 'mean'),
                intl=('roi_intl_medio', 'mean'),
                fest=('pct_com_festival', 'mean'),
                n=('n_obras', 'sum'))
           .assign(label=lambda x: x['categoria'].map(cats_cmp))
           .sort_values('dom', ascending=False))

    fig = make_fig(16, 7)
    axes = fig.subplots(1, 2)

    # Left: ROI dom
    ax = axes[0]
    style_ax(ax, ylabel='ROI Doméstico / FSA')
    colors = [GREEN if 'Distrib' in l else ACCENT if 'Produt' in l else MUTED
              for l in sub['label']]
    bars = ax.barh(sub['label'], sub['dom'], color=colors, alpha=0.9, zorder=3)
    ax.set_title('ROI Doméstico médio', color=TEXT, fontsize=14, fontweight='bold', pad=10)
    ax.tick_params(axis='y', labelsize=12)
    for b in bars:
        w = b.get_width()
        if w > 0.02:
            ax.text(w + 0.01, b.get_y() + b.get_height()/2,
                    f'{w:.2f}', va='center', fontsize=11, color=TEXT2)
    ax.set_facecolor(SURF)

    # Right: ROI intl
    ax = axes[1]
    style_ax(ax, ylabel='ROI Internacional médio')
    bars = ax.barh(sub['label'], sub['intl'], color=GOLD, alpha=0.9, zorder=3)
    ax.set_title('ROI Internacional médio', color=TEXT, fontsize=14, fontweight='bold', pad=10)
    ax.tick_params(axis='y', labelsize=12)
    for b in bars:
        w = b.get_width()
        if w > 0.1:
            ax.text(w + 0.05, b.get_y() + b.get_height()/2,
                    f'{w:.2f}', va='center', fontsize=11, color=TEXT2)
    ax.set_facecolor(SURF)

    fig.patch.set_facecolor(BG)
    fig.suptitle('Perfil de retorno por tipo de chamada: doméstico vs. internacional',
                 color=TEXT, fontsize=15, fontweight='bold', y=1.01)
    fig.tight_layout(pad=2)
    return fig_b64(fig)

# ─────────────────────────────────────────────────────────────────────────────
# CHART 3 — Scatter de produtoras por cluster
# ─────────────────────────────────────────────────────────────────────────────
def chart3():
    df = df_prod.copy()
    df['investimento_fsa_deflac'] = pd.to_numeric(df['investimento_fsa_deflac'], errors='coerce').fillna(0)
    df['receita_total_deflac']    = pd.to_numeric(df['receita_total_deflac'], errors='coerce').fillna(0)

    fig = make_fig(14, 7.5)
    ax = fig.add_subplot(111)
    style_ax(ax, xlabel='Investimento FSA total (R$ deflac. 2024)',
             ylabel='Receita total estimada (R$ deflac. 2024)', grid_axis='both')

    for cluster, color in CLUSTER_COLORS.items():
        sub = df[df['cluster'] == cluster]
        sz = np.where(sub['n_obras'] > 10, 80, np.where(sub['n_obras'] > 5, 50, 28))
        ax.scatter(sub['investimento_fsa_deflac'], sub['receita_total_deflac'],
                   c=color, s=sz, alpha=0.75, label=cluster, zorder=4, linewidths=0)

    # Line of parity (ROI=1)
    mx = max(df['investimento_fsa_deflac'].max(), df['receita_total_deflac'].max()) * 1.1
    ax.plot([0, mx], [0, mx], color=TEXT2, linewidth=1, linestyle='--', alpha=0.3,
            label='ROI = 1 (paridade)', zorder=2)

    # Highlight notable companies
    notaveis = {
        'Filmes de Plástico': GOLD,
        'Globo': CORAL,
        'Downtown': PURPLE,
    }
    for _, row in df.iterrows():
        for nome, cor in notaveis.items():
            if nome.lower() in str(row.get('razao_social', '')).lower():
                ax.scatter(row['investimento_fsa_deflac'], row['receita_total_deflac'],
                           c=cor, s=120, zorder=6, marker='*', linewidths=0)
                ax.annotate(str(row['razao_social'])[:22],
                            (row['investimento_fsa_deflac'], row['receita_total_deflac']),
                            textcoords='offset points', xytext=(8, 4),
                            fontsize=9, color=cor, zorder=7)

    ax.xaxis.set_major_formatter(FuncFormatter(fmt_m))
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_m))
    ax.set_title('Distribuição das produtoras por retorno e investimento (clusters)',
                 color=TEXT, fontsize=15, fontweight='bold', pad=14)
    ax.legend(fontsize=12, facecolor=SURF2, edgecolor=RULE, labelcolor=TEXT,
              framealpha=0.95, markerscale=1.3)
    fig.tight_layout(pad=1.8)
    return fig_b64(fig)

# ─────────────────────────────────────────────────────────────────────────────
# CHART 4 — Proliferação: histograma de investimento por produtora
# ─────────────────────────────────────────────────────────────────────────────
def chart4():
    df = df_prod.copy()
    df['fsa'] = pd.to_numeric(df['investimento_fsa_deflac'], errors='coerce').fillna(0)
    df = df[df['fsa'] > 0]

    # Investment tiers
    bins   = [0, 500_000, 1_000_000, 2_500_000, 5_000_000, 10_000_000, 25_000_000]
    labels = ['< 500K', '500K–1M', '1–2,5M', '2,5–5M', '5–10M', '> 10M']
    df['tier'] = pd.cut(df['fsa'], bins=bins, labels=labels, right=True)

    counts = df['tier'].value_counts().reindex(labels).fillna(0)
    total_inv = df.groupby('tier')['fsa'].sum().reindex(labels).fillna(0)

    fig = make_fig(14, 6.5)
    axes = fig.subplots(1, 2)

    # Left: count of producers
    ax = axes[0]
    style_ax(ax, ylabel='Nº de produtoras')
    bar_colors = [ACCENT if l in ('< 500K', '500K–1M', '1–2,5M') else GOLD
                  for l in labels]
    bars = ax.bar(labels, counts, color=bar_colors, alpha=0.9, zorder=3)
    ax.set_title('Nº de produtoras por faixa de investimento FSA',
                 color=TEXT, fontsize=13, fontweight='bold', pad=10)
    ax.tick_params(axis='x', labelsize=11)
    for b in bars:
        h = b.get_height()
        ax.text(b.get_x() + b.get_width()/2, h + 5, f'{int(h)}',
                ha='center', va='bottom', fontsize=11, color=TEXT, fontweight='bold')
    ax.set_facecolor(SURF)

    # Right: total investment per tier
    ax = axes[1]
    style_ax(ax, ylabel='Investimento total (R$ deflac. 2024)')
    bars2 = ax.bar(labels, total_inv / 1e6, color=GOLD, alpha=0.9, zorder=3)
    ax.set_title('Volume de investimento FSA por faixa (R$ M)',
                 color=TEXT, fontsize=13, fontweight='bold', pad=10)
    ax.tick_params(axis='x', labelsize=11)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f'R${v:.0f}M'))
    for b in bars2:
        h = b.get_height()
        if h > 1:
            ax.text(b.get_x() + b.get_width()/2, h + 2, f'R${h:.0f}M',
                    ha='center', va='bottom', fontsize=10, color=TEXT2)
    ax.set_facecolor(SURF)

    fig.patch.set_facecolor(BG)
    fig.suptitle('Fragmentação do investimento público: muitas produtoras, pouco ticket médio',
                 color=TEXT, fontsize=15, fontweight='bold', y=1.02)
    fig.tight_layout(pad=2)
    return fig_b64(fig)

# ─────────────────────────────────────────────────────────────────────────────
# CHART 5 — Curva de Lorenz
# ─────────────────────────────────────────────────────────────────────────────
def chart5():
    df = df_prod.copy()
    df['fsa'] = pd.to_numeric(df['investimento_fsa_deflac'], errors='coerce').fillna(0)
    vals = df['fsa'].sort_values().values
    cumvals = np.cumsum(vals)
    cumvals = cumvals / cumvals[-1]
    cumshare = np.arange(1, len(vals)+1) / len(vals)

    gini = 1 - 2 * np.trapz(cumvals, cumshare)

    fig = make_fig(9, 7)
    ax = fig.add_subplot(111)
    style_ax(ax, xlabel='Fração de produtoras (% acumulado)',
             ylabel='Fração do investimento FSA (% acumulado)', grid_axis='both')

    ax.plot([0, 1], [0, 1], color=TEXT2, linewidth=1.2, linestyle='--',
            alpha=0.45, label='Igualdade perfeita', zorder=2)
    ax.fill_between(cumshare, cumvals, cumshare, alpha=0.15, color=ACCENT, zorder=2)
    ax.plot(cumshare, cumvals, color=ACCENT, linewidth=2.5,
            label=f'Lorenz — investimento FSA', zorder=3)

    ax.annotate(f'Gini = {gini:.2f}', xy=(0.65, 0.25),
                fontsize=16, color=ACCENT, fontweight='bold',
                bbox=dict(facecolor=SURF2, edgecolor=ACCENT, alpha=0.85,
                          boxstyle='round,pad=0.5'))

    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f'{v:.0%}'))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f'{v:.0%}'))
    ax.set_title('Curva de Lorenz — distribuição do investimento FSA por produtora',
                 color=TEXT, fontsize=14, fontweight='bold', pad=14)
    ax.legend(fontsize=13, facecolor=SURF2, edgecolor=RULE, labelcolor=TEXT)
    fig.tight_layout(pad=1.8)
    return fig_b64(fig)

# ─────────────────────────────────────────────────────────────────────────────
# CHART 6 — Scatter obras: investimento vs bilheteria (capacidade de carga)
# ─────────────────────────────────────────────────────────────────────────────
def chart6():
    # Per-obra, apenas FSA (excluindo TV e renúncia)
    df = df_inv.copy()
    df = df[df['categoria'].str.startswith('FSA', na=False)].copy()
    df['bil'] = pd.to_numeric(df['bilheteria_deflac_r2024'], errors='coerce').fillna(0)

    # Look for investment column
    inv_col = None
    for c in ['investimento_fsa', 'valor_fsa', 'fsa_deflac']:
        if c in df.columns:
            inv_col = c
            break

    if inv_col is None:
        # Use a fixed average based on category aggregates from df_cham
        agg = df_cham.groupby('categoria').apply(
            lambda g: g['investimento_fsa_deflac'].sum() / g['n_obras'].sum()
        ).rename('ticket')
        df = df.merge(agg, left_on='categoria', right_index=True, how='left')
        inv_col = 'ticket'

    df[inv_col] = pd.to_numeric(df[inv_col], errors='coerce').fillna(0)
    sub = df[(df[inv_col] > 0) & (df['bil'] > 0)].copy()

    fig = make_fig(14, 7)
    ax = fig.add_subplot(111)
    style_ax(ax, xlabel='Investimento FSA estimado por obra (R$ deflac. 2024)',
             ylabel='Bilheteria (R$ deflac. 2024)', grid_axis='both')

    # Color by category group
    cat_color = {}
    for cat in sub['categoria'].unique():
        if 'Distribuidora' in cat:
            cat_color[cat] = GREEN
        elif 'Festivais' in cat:
            cat_color[cat] = GOLD
        elif 'Automático' in cat:
            cat_color[cat] = MUTED
        elif 'Coprodução' in cat:
            cat_color[cat] = PURPLE
        else:
            cat_color[cat] = ACCENT

    for cat in sub['categoria'].unique():
        s = sub[sub['categoria'] == cat]
        ax.scatter(s[inv_col], s['bil'], c=cat_color.get(cat, ACCENT),
                   s=30, alpha=0.55, zorder=3, linewidths=0)

    # Parity line
    mx = max(sub[inv_col].max(), sub['bil'].max()) * 1.1
    ax.plot([0, mx], [0, mx], color=TEXT2, linewidth=1, linestyle='--',
            alpha=0.35, zorder=2, label='ROI = 1')

    ax.xaxis.set_major_formatter(FuncFormatter(fmt_m))
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_m))
    ax.set_title('Relação entre investimento FSA e bilheteria por obra (obras com dados disponíveis)',
                 color=TEXT, fontsize=14, fontweight='bold', pad=14)

    legend_elements = [
        mpatches.Patch(color=GREEN,  label='Seletivo Distribuidora'),
        mpatches.Patch(color=GOLD,   label='Seletivo Festivais'),
        mpatches.Patch(color=ACCENT, label='Seletivo Produtora'),
        mpatches.Patch(color=MUTED,  label='Automático'),
        mpatches.Patch(color=PURPLE, label='Coprodução Intl'),
    ]
    ax.legend(handles=legend_elements, fontsize=12, facecolor=SURF2,
              edgecolor=RULE, labelcolor=TEXT, framealpha=0.95)
    fig.tight_layout(pad=1.8)
    return fig_b64(fig)

# ─────────────────────────────────────────────────────────────────────────────
# CHART 7 — Diversidade: % diretoras por categoria de chamada
# ─────────────────────────────────────────────────────────────────────────────
def chart7():
    # Use pct_genero_feminino from chamada dataset
    df = df_cham.copy()
    df = df[df['categoria'].isin(CAT_SHORT)].copy()
    df['pct_genero_feminino'] = pd.to_numeric(df['pct_genero_feminino'], errors='coerce').fillna(0)
    df['pct_diversidade_regional'] = pd.to_numeric(df['pct_diversidade_regional'], errors='coerce').fillna(0)

    agg = (df.groupby('categoria', as_index=False)
             .agg(fem=('pct_genero_feminino', 'mean'),
                  div=('pct_diversidade_regional', 'mean'),
                  n=('n_obras', 'sum'))
             .assign(label=lambda x: x['categoria'].map(CAT_SHORT))
             .sort_values('fem', ascending=True))

    fig = make_fig(14, 6.5)
    axes = fig.subplots(1, 2)

    for ax, col, title, color in [
        (axes[0], 'fem', '% Obras com diretora mulher', PURPLE),
        (axes[1], 'div', '% Obras de regiões fora do eixo', GOLD),
    ]:
        style_ax(ax)
        bars = ax.barh(agg['label'], agg[col] * 100, color=color, alpha=0.88, zorder=3)
        ax.set_title(title, color=TEXT, fontsize=13, fontweight='bold', pad=10)
        ax.tick_params(axis='y', labelsize=11)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f'{v:.0f}%'))
        ax.set_facecolor(SURF)
        for b in bars:
            w = b.get_width()
            if w > 1:
                ax.text(w + 0.3, b.get_y() + b.get_height()/2,
                        f'{w:.1f}%', va='center', fontsize=10, color=TEXT2)

    fig.patch.set_facecolor(BG)
    fig.suptitle('Diversidade de gênero e regional por tipo de chamada',
                 color=TEXT, fontsize=15, fontweight='bold', y=1.02)
    fig.tight_layout(pad=2)
    return fig_b64(fig)

# ─────────────────────────────────────────────────────────────────────────────
# Gera todos os charts
# ─────────────────────────────────────────────────────────────────────────────
print("Gerando gráficos...")
charts = [
    ("chart1", chart1),
    ("chart2", chart2),
    ("chart3", chart3),
    ("chart4", chart4),
    ("chart5", chart5),
    ("chart6", chart6),
    ("chart7", chart7),
]

b64_list = []
for name, fn in charts:
    print(f"  {name}...", end=' ', flush=True)
    try:
        b64 = fn()
        b64_list.append(b64)
        print(f"OK ({len(b64)//1000}KB)")
    except Exception as e:
        print(f"ERRO: {e}")
        import traceback; traceback.print_exc()
        b64_list.append(None)

# ─────────────────────────────────────────────────────────────────────────────
# Atualiza o HTML
# ─────────────────────────────────────────────────────────────────────────────
print("\nAtualizando HTML...")
with open(HTML_PATH, encoding='utf-8') as f:
    html = f.read()

# 1. Substitui fontes: Literata → Lora
html = html.replace(
    'family=Inter:wght@400;500;600;700;800&family=Literata:ital,wght@0,400;0,500;0,600;0,700;1,400;1,500&display=swap',
    'family=Inter:wght@400;500;600;700;800&family=Lora:ital,wght@0,400;0,500;0,600;0,700;1,400;1,700;1,400;1,700&display=swap'
)
html = html.replace("'Literata',Georgia,'Times New Roman',serif", "'Lora',Georgia,serif")
html = html.replace("'Literata', Georgia, 'Times New Roman', serif", "'Lora', Georgia, serif")

# 2. Remove "Política audiovisual · " do article-meta
html = html.replace('Política audiovisual · Maio 2026', 'Maio 2026')

# 3. Adiciona link G1 na citação de Gabriel Martins
# A cite já tem: — Gabriel Martins (G1, 2026)
# Substituímos por link para a matéria de referência (Rede Pampa/Agência Globo)
G1_URL = "https://www.radiopampa.com.br/prestigio-recente-no-oscar-abriu-portas-para-o-cinema-brasileiro-cineastas-avaliam/"
old_cite = '— Gabriel Martins (G1, 2026)'
new_cite = f'— Gabriel Martins (<a href="{G1_URL}" target="_blank" rel="noopener">G1, 2026</a>)'
html = html.replace(old_cite, new_cite)

# 4. Aumenta o container das imagens (width breakout mais amplo)
# Troca artigo img regra de 100px de margem para 150px
html = html.replace(
    'article img{display:block;width:calc(100% + 100px);max-width:none;margin:32px -50px;border-radius:10px}',
    'article img{display:block;width:calc(100% + 160px);max-width:none;margin:40px -80px;border-radius:10px}'
)
html = html.replace(
    '@media(max-width:820px){article img{width:100%;max-width:100%;margin:24px 0}}',
    '@media(max-width:860px){article img{width:100%;max-width:100%;margin:24px 0}}'
)

# 5. Substitui os 7 gráficos base64
# Padrão: <p><img role="img" src="data:image/png;base64,XXXX" ...>
IMG_PATTERN = re.compile(r'<p><img role="img" src="data:image/png;base64,([^"]+)"[^>]*></p>')
matches = list(IMG_PATTERN.finditer(html))
print(f"  Encontradas {len(matches)} imagens no HTML")

if len(matches) != len(b64_list):
    print(f"  AVISO: esperava {len(b64_list)} imagens, encontrou {len(matches)}")

# Substitui do fim para o início (preserva offsets)
replacements = []
for i, (m, b64) in enumerate(zip(matches, b64_list)):
    if b64:
        new_tag = f'<p><img role="img" src="data:image/png;base64,{b64}" alt="Gráfico {i+1}" style="width:100%"></p>'
        replacements.append((m.start(), m.end(), new_tag))

for start, end, new_tag in reversed(replacements):
    html = html[:start] + new_tag + html[end:]

print(f"  {len(replacements)} gráficos substituídos")

# Salva
with open(HTML_PATH, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"\nDone! Arquivo salvo: {HTML_PATH}")

# Copia para docs/politica.html se existir
docs_path = os.path.join(ROOT, "docs", "politica.html")
if os.path.exists(docs_path):
    import shutil
    # docs/politica.html pode ser diferente (versão pré-edição)
    # Vamos criar uma cópia do output_final no docs
    shutil.copy2(HTML_PATH, docs_path)
    print(f"Copiado para: {docs_path}")
