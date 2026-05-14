"""Gera gráfico: Investimento Total + Produtoras Ativas + Ticket Médio por Produtora (2014-2023).

Estilo: fundo claro, dois painéis empilhados (como v6).
Recorte: apenas obras com bilheteria > 0.
Ticket médio = (investimento_total + receita_total) / nº produtoras ativas no ano.
"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

obra = pd.read_csv(
    r"C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\datasets\base_nivel_obra.csv",
    sep=';', low_memory=False
)

# Recorte: obras onde cruzamento investimento × bilheteria foi possível
obra = obra[obra['bilheteria_deflac'] > 0].copy()

obra['renda_bruta'] = obra['bilheteria_deflac'] + obra['outras_janelas_deflac']

by_year = obra.groupby('ano').agg(
    inv_total=('investimento_total_deflac', 'sum'),
    renda_bruta=('renda_bruta', 'sum'),
    n_obras=('CPB', 'count'),
    n_produtoras=('CNPJ_produtora', 'nunique'),
).reset_index()

# RLP = 15% da renda bruta (bilheteria + outras janelas)
by_year['rlp'] = by_year['renda_bruta'] * 0.15
by_year['ticket_prod'] = (by_year['inv_total'] + by_year['rlp']) / by_year['n_produtoras']
by_year['inv_total_mi'] = by_year['inv_total'] / 1e6
by_year['ticket_prod_mi'] = by_year['ticket_prod'] / 1e6

# ── Style (light, matching v6) ──
BG = '#ffffff'
PANEL_BG = '#f8f8f8'
TEXT = '#2a2a2a'
MUTED = '#666666'
GRID = '#e0e0e0'
GREEN = '#6aaa5a'
PURPLE = '#8b5ab0'
GOLD = '#d4940a'

plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.sans-serif': ['Inter', 'Segoe UI', 'Helvetica', 'Arial'],
    'font.size': 10,
    'axes.facecolor': BG,
    'figure.facecolor': BG,
    'text.color': TEXT,
    'axes.labelcolor': TEXT,
    'xtick.color': MUTED,
    'ytick.color': MUTED,
    'axes.edgecolor': GRID,
    'grid.color': GRID,
    'grid.alpha': 0.6,
})

fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 9), height_ratios=[1.3, 1],
                                gridspec_kw={'hspace': 0.35})

years = by_year['ano'].values
x = np.arange(len(years))

# ═══════════════════════════════════════════════════════════════
# PAINEL 1: Barras = Investimento Total, Linha = Produtoras
# ═══════════════════════════════════════════════════════════════
bars = ax1.bar(x, by_year['inv_total_mi'], 0.55,
               color=GREEN, alpha=0.7, label='Total investido (FSA + Renúncia)',
               edgecolor='none', zorder=3)

# Value labels on bars
for bar, val in zip(bars, by_year['inv_total_mi']):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 8,
             f'R$ {val:.0f}M', ha='center', va='bottom', fontsize=8,
             color=GREEN, fontweight=700)

ax1.set_ylabel('Total investido\nR$ milhões (deflac. 2024)', color=GREEN, fontsize=10, fontweight=600)
ax1.set_xticks(x)
ax1.set_xticklabels(years, fontsize=10, fontweight=600)
ax1.set_ylim(0, max(by_year['inv_total_mi']) * 1.25)
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'R$ {v:,.0f}M'))
ax1.grid(axis='y', linestyle='--', alpha=0.4, zorder=0)
ax1.tick_params(axis='y', labelcolor=GREEN)

# Right axis: produtoras
ax1r = ax1.twinx()
ax1r.plot(x, by_year['n_produtoras'], color=PURPLE, linewidth=2.5,
          marker='o', markersize=7, markerfacecolor=PURPLE,
          markeredgecolor='white', markeredgewidth=1.5,
          label='Nº de produtoras brasileiras', zorder=5)

for xi, val in zip(x, by_year['n_produtoras']):
    ax1r.annotate(str(val), (xi, val), textcoords="offset points",
                  xytext=(0, 12), ha='center', fontsize=9,
                  color=PURPLE, fontweight=700)

ax1r.set_ylabel('Nº de produtoras', color=PURPLE, fontsize=10, fontweight=600)
ax1r.tick_params(axis='y', labelcolor=PURPLE)
ax1r.set_ylim(0, max(by_year['n_produtoras']) * 1.35)
ax1r.spines['right'].set_color(PURPLE)
ax1r.spines['right'].set_alpha(0.4)

# Title panel 1
ax1.set_title('Série Histórica — FSA + Renúncia Fiscal  (2014–2023)\n'
              'Obras com fomento público (FSA e/ou renúncia) e bilheteria > 0',
              fontsize=13, fontweight=700, color=TEXT, loc='left', pad=12)

# Legend panel 1
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax1r.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2,
           loc='upper right', frameon=True, framealpha=0.9,
           edgecolor=GRID, fontsize=9)

# ═══════════════════════════════════════════════════════════════
# PAINEL 2: Distribuição do Ticket por Produtora (box plot)
# ═══════════════════════════════════════════════════════════════
obra['renda_bruta_obra'] = obra['bilheteria_deflac'] + obra['outras_janelas_deflac']
by_prod_year = obra.groupby(['ano', 'CNPJ_produtora']).agg(
    inv=('investimento_total_deflac', 'sum'),
    renda=('renda_bruta_obra', 'sum'),
).reset_index()
by_prod_year['ticket'] = (by_prod_year['inv'] + by_prod_year['renda'] * 0.15) / 1e6

box_data = [by_prod_year[by_prod_year['ano'] == y]['ticket'].values for y in years]

bp = ax2.boxplot(box_data, positions=x, widths=0.45, patch_artist=True,
                 showfliers=False, zorder=3,
                 medianprops=dict(color='white', linewidth=2),
                 whiskerprops=dict(color=MUTED, linewidth=1),
                 capprops=dict(color=MUTED, linewidth=1))
for patch in bp['boxes']:
    patch.set_facecolor(GOLD)
    patch.set_alpha(0.7)
    patch.set_edgecolor(GOLD)

# Median labels
for xi, year in enumerate(years):
    med = np.median(by_prod_year[by_prod_year['ano'] == year]['ticket'].values)
    ax2.text(xi, med + 0.25, f'R$ {med:.1f}M', ha='center', va='bottom',
             fontsize=8, color=GOLD, fontweight=700)

# R$ 1M viability line
ax2.axhline(1.0, color='#e67e22', linestyle='--', linewidth=1.5, alpha=0.7, zorder=4)
ax2.text(len(x) - 0.5, 1.07, 'Limiar de viabilidade (R$ 1M/ano)',
         fontsize=8, color='#e67e22', fontweight=600, ha='right', va='bottom')

# R$ 500K subsistence line
ax2.axhline(0.5, color='#c0392b', linestyle='--', linewidth=1.8, alpha=0.8, zorder=4)
ax2.text(len(x) - 0.5, 0.33, 'Custo fixo mínimo (R$ 500K/ano)',
         fontsize=8, color='#c0392b', fontweight=600, ha='right', va='bottom')

# % below thresholds per year
for xi, year in enumerate(years):
    sub = by_prod_year[by_prod_year['ano'] == year]['ticket']
    pct_500 = (sub < 0.5).sum() / len(sub) * 100
    pct_1m = (sub < 1.0).sum() / len(sub) * 100
    ax2.text(xi, -0.55, f'{pct_1m:.0f}%', ha='center', fontsize=7.5,
             color='#e67e22', fontweight=700)
    ax2.text(xi, -0.85, f'{pct_500:.0f}%', ha='center', fontsize=7.5,
             color='#c0392b', fontweight=700)
ax2.text(-0.8, -0.55, '< 1M:', ha='right', fontsize=7.5, color='#e67e22', fontweight=600)
ax2.text(-0.8, -0.85, '< 500K:', ha='right', fontsize=7.5, color='#c0392b', fontweight=600)

ax2.set_title('Distribuição do Ticket por Produtora/Ano (Investimento + 15% da Renda Bruta)',
              fontsize=13, fontweight=700, color=TEXT, loc='left', pad=12)
ax2.set_ylabel('R$ milhões (deflac. 2024)', color=GOLD, fontsize=10, fontweight=600)
ax2.set_xticks(x)
ax2.set_xticklabels(years, fontsize=10, fontweight=600)
ax2.set_ylim(-1.1, 12)
ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'R$ {v:.0f}M' if v >= 1 else f'R$ {v:.1f}M'))
ax2.grid(axis='y', linestyle='--', alpha=0.4, zorder=0)
ax2.tick_params(axis='y', labelcolor=GOLD)

# Footer
n_obras = len(obra)
fig.text(0.04, 0.01,
         f'Fonte: BNDES/FSA + ANCINE (Renúncia Fiscal) | Valores deflacionados pelo IPCA (base: R$ 2024) | '
         f'Universo: {n_obras:,} obras, {obra["CNPJ_produtora"].nunique():,} produtoras | '
         f'Recorte: obras com bilheteria > 0 | 2024: dados não disponíveis',
         fontsize=7.5, color=MUTED, ha='left')

plt.tight_layout(rect=[0, 0.03, 1, 1])

OUT = r"C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\grafico_ticket_produtoras.png"
fig.savefig(OUT, dpi=200, bbox_inches='tight', facecolor=BG)
plt.close()
print(f"Saved: {OUT}")
