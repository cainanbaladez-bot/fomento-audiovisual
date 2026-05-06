"""
07_gerar_painel_final.py
Gera o painel unificado com todas as análises.
Output: output_final/Análise do Retorno do Fomento Público ao Audiovisual Brasileiro (FSA - Renúncia Fiscal)_v2.html

Estrutura do sidebar:
  - Visão Geral (sub-tabs: Geral | Financeiro | Retorno Internacional)
  - Categorias das Chamadas (critério de seleção + comparativo por categoria)
  - Produtoras (painel produtoras + concentração + ticket médio por cluster)
  - Concentração (concentração e distribuição de produtoras)
  - Curtas + Longas (trajetória de diretores)
  - Diversidade (políticas afirmativas)
  - Soft Power (crítica cinematográfica e citação acadêmica)
"""

import re
import os
import json
import html
import unicodedata
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Derive obra count from tabela consolidada (avoid hardcoding)
_tabela_path = os.path.join(BASE, 'resultados', 'tabela_consolidada_obras.xlsx')
try:
    _df_count = pd.read_excel(_tabela_path, sheet_name=0)
    N_OBRAS = len(_df_count)
except Exception:
    N_OBRAS = 610  # fallback
del _df_count

SRC_CMP   = os.path.join(BASE, 'resultados', 'painel_comparativo.html')
SRC_CONC  = os.path.join(BASE, 'resultados', 'painel_concentracao_produtoras.html')
SRC_CS    = os.path.join(BASE, 'resultados', 'painel_criterio_selecao.html')
SRC_PR    = os.path.join(BASE, 'resultados', 'painel_produtoras.html')
SRC_DIV   = os.path.join(BASE, 'resultados', 'painel_diversidade.html')
SRC_SP    = os.path.join(BASE, 'resultados', 'painel_softpower.html')
OUT       = os.path.join(BASE, 'output_final',
                         'Análise do Retorno do Fomento Público ao Audiovisual Brasileiro'
                         ' (FSA - Renúncia Fiscal)_v2.html')

os.makedirs(os.path.dirname(OUT), exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# Helper: prefix IDs in HTML + JS
# ─────────────────────────────────────────────────────────────────────────────
def prefix_ids(content, ids, prefix):
    """Prefix all element IDs and JS references with `prefix`."""
    for eid in sorted(ids, key=len, reverse=True):
        e = re.escape(eid)
        new_id = prefix + eid
        patterns = [
            (re.compile(r'(id=")' + e + r'"'),              r'\g<1>' + new_id + '"'),
            (re.compile(r"(id=')" + e + r"'"),              r"\g<1>" + new_id + "'"),
            (re.compile(r'getElementById\("' + e + r'"\)'), 'getElementById("' + new_id + '")'),
            (re.compile(r"getElementById\('" + e + r"'\)"), "getElementById('" + new_id + "')"),
            (re.compile(r'querySelector\("#' + e + r'"\)'), 'querySelector("#' + new_id + '")'),
            (re.compile(r"querySelector\('#" + e + r"'\)"), "querySelector('#" + new_id + "')"),
            (re.compile(r"Plotly\.newPlot\(\s*'" + e + r"'"), "Plotly.newPlot('" + new_id + "'"),
            (re.compile(r'Plotly\.newPlot\(\s*"' + e + r'"'), 'Plotly.newPlot("' + new_id + '"'),
            (re.compile(r"Plotly\.react\(\s*'" + e + r"'"),   "Plotly.react('" + new_id + "'"),
            (re.compile(r'Plotly\.react\(\s*"' + e + r'"'),   'Plotly.react("' + new_id + '"'),
        ]
        for pat, repl in patterns:
            content = pat.sub(repl, content)
    return content


def extract_css(html):
    """Extract first <style> block content from HTML."""
    m = re.search(r'<style[^>]*>([\s\S]*?)</style>', html)
    return m.group(1) if m else ''


# ─────────────────────────────────────────────────────────────────────────────
# Read sources
# ─────────────────────────────────────────────────────────────────────────────
def read(path):
    with open(path, encoding='utf-8') as f:
        return f.read()

def get_body(html):
    body_start = html.find('<body>') + len('<body>')
    body_end = html.rfind('</body>')
    if body_start > 0 and body_end > body_start:
        return html[body_start:body_end]
    return html

def extract_tab_panel(html, tid):
    start = html.find(f'id="{tid}"')
    if start == -1:
        return f'<div id="{tid}"><p>not found</p></div>'
    div_start = html.rfind('<div', 0, start)
    pos = div_start
    depth = 0
    while pos < len(html):
        open_m = html.find('<div', pos)
        close_m = html.find('</div>', pos)
        if open_m == -1 and close_m == -1:
            break
        if open_m != -1 and (close_m == -1 or open_m < close_m):
            depth += 1
            pos = open_m + 4
        else:
            depth -= 1
            pos = close_m + 6
            if depth == 0:
                return html[div_start:pos]
    return html[div_start:]

def _fmt_num(value, decimals=0):
    try:
        if pd.isna(value):
            return ''
        value = float(value)
    except Exception:
        return ''
    if decimals:
        return f'{value:,.{decimals}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return f'{value:,.0f}'.replace(',', '.')

def _safe_text(value):
    if pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text.lower() == 'nan' else text

def _table_rows(df, columns, max_rows=None):
    rows = []
    data = df.head(max_rows) if max_rows else df
    for _, row in data.iterrows():
        cells = []
        for label, col, align in columns:
            value = row.get(col, '')
            cells.append(
                f'<td style="padding:6px 8px;text-align:{align};vertical-align:top">'
                f'{html.escape(_safe_text(value))}</td>'
            )
        rows.append('<tr style="border-bottom:1px solid rgba(255,255,255,.05)">' + ''.join(cells) + '</tr>')
    return ''.join(rows)

def _build_international_listing_html():
    cons_path = os.path.join(BASE, 'resultados', 'tabela_consolidada_obras.xlsx')
    if not os.path.exists(cons_path):
        return ''
    df = pd.read_excel(cons_path, sheet_name='Obras')
    for col in ['Pontuação Festivais', 'Adm. EU — Lumière', 'VOD Intl — N Países', 'VOD Intl — N Plataformas']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['tem_festival'] = df.get('Pontuação Festivais', 0) > 0
    df['tem_bilheteria_eu'] = df.get('Adm. EU — Lumière', 0) > 0
    df['tem_vod'] = df.get('VOD Intl — N Países', 0) > 0
    intl = df[df['tem_festival'] | df['tem_bilheteria_eu'] | df['tem_vod']].copy()
    if not len(intl):
        return ''
    intl['Festival'] = intl['tem_festival'].map(lambda v: 'sim' if v else '')
    intl['Bilheteria EU'] = intl['tem_bilheteria_eu'].map(lambda v: 'sim' if v else '')
    intl['VOD'] = intl['tem_vod'].map(lambda v: 'sim' if v else '')
    intl['Adm. EU'] = intl.get('Adm. EU — Lumière', 0).map(lambda v: _fmt_num(v))
    intl['N paises VOD'] = intl.get('VOD Intl — N Países', 0).map(lambda v: _fmt_num(v))
    intl['N plataformas VOD'] = intl.get('VOD Intl — N Plataformas', 0).map(lambda v: _fmt_num(v))
    intl['Pontuacao festivais'] = intl.get('Pontuação Festivais', 0).map(lambda v: _fmt_num(v, 1))
    intl['Paises'] = intl.get('Países Lista', '').fillna('')
    intl['Paises festivais'] = intl.get('Países Festivais', '').fillna('')
    intl['Paises Lumiere'] = intl.get('Países Lumière', '').fillna('')
    intl['Paises VOD'] = intl.get('Países VOD Europa', '').fillna('')
    intl = intl.sort_values(
        ['tem_festival', 'tem_bilheteria_eu', 'tem_vod', 'Pontuação Festivais', 'Adm. EU — Lumière', 'VOD Intl — N Países', 'Projeto'],
        ascending=[False, False, False, False, False, False, True]
    )
    cols = [
        ('Obra', 'Projeto', 'left'),
        ('Ano', 'Ano', 'center'),
        ('Festival', 'Festival', 'center'),
        ('Bilheteria EU', 'Bilheteria EU', 'center'),
        ('VOD', 'VOD', 'center'),
        ('Pontos festivais', 'Pontuacao festivais', 'right'),
        ('Adm. EU', 'Adm. EU', 'right'),
        ('Paises VOD', 'N paises VOD', 'right'),
        ('Plataformas', 'N plataformas VOD', 'right'),
        ('Paises alcancados', 'Paises', 'left'),
        ('Paises festivais', 'Paises festivais', 'left'),
        ('Paises Lumiere', 'Paises Lumiere', 'left'),
        ('Paises VOD', 'Paises VOD', 'left'),
    ]
    head = ''.join(
        f'<th style="padding:6px 8px;text-align:{align};color:var(--muted);white-space:nowrap">{html.escape(label)}</th>'
        for label, _, align in cols
    )
    rows = _table_rows(intl, cols)
    n_fest = int(intl['tem_festival'].sum())
    n_bil = int(intl['tem_bilheteria_eu'].sum())
    n_vod = int(intl['tem_vod'].sum())
    return f'''
<div class="card" style="margin:20px 0 8px;background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:8px;padding:14px 16px">
  <div style="font-family:var(--font-head);font-size:17px;font-style:italic;color:var(--text);margin-bottom:5px">Listagem completa do alcance internacional</div>
  <div style="font-size:11px;color:var(--muted);line-height:1.55;margin-bottom:12px">
    {len(intl)} longas com presença internacional identificada: {n_fest} em festivais, {n_bil} com bilheteria EU/Lumière e {n_vod} em VOD Europa.
  </div>
  <div style="overflow-x:auto;max-height:520px">
    <table style="width:100%;border-collapse:collapse;font-size:10.5px">
      <thead><tr style="border-bottom:1px solid var(--border)">{head}</tr></thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</div>'''

def _trend_trace(x, y, x_new, color):
    if len(x) < 2:
        return None
    try:
        coef = np.polyfit(x, y, 1)
        y_new = np.polyval(coef, x_new)
    except Exception:
        return None
    return go.Scatter(
        x=x_new,
        y=y_new,
        mode='lines',
        line=dict(color=color, width=2, dash='dash'),
        hoverinfo='skip',
        showlegend=False,
    )

def _corr_label(df, x_col, y_col):
    sub = df[[x_col, y_col]].dropna()
    sub = sub[(sub[x_col] > 0) & (sub[y_col] > 0)]
    if len(sub) < 3:
        return 'n insuficiente'
    corr = sub[x_col].corr(sub[y_col])
    if pd.isna(corr):
        return 'n insuficiente'
    return f'n={len(sub)} | r={corr:.2f}'

def _build_festival_vod_lumiere_scatter_html():
    cons_path = os.path.join(BASE, 'resultados', 'tabela_consolidada_obras.xlsx')
    if not os.path.exists(cons_path):
        return ''
    df = pd.read_excel(cons_path, sheet_name='Obras')
    required = ['Projeto', 'Ano', 'Pontuação Festivais', 'Adm. EU — Lumière', 'VOD Intl — N Países', 'VOD Intl — N Plataformas']
    if any(col not in df.columns for col in required):
        return ''

    for col in ['Pontuação Festivais', 'Adm. EU — Lumière', 'VOD Intl — N Países', 'VOD Intl — N Plataformas']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['Projeto'] = df['Projeto'].fillna('').astype(str)
    df['Ano'] = df['Ano'].fillna('').astype(str)

    left = df[(df['Pontuação Festivais'] > 0) & (df['VOD Intl — N Países'] > 0)].copy()
    right = df[(df['Pontuação Festivais'] > 0) & (df['Adm. EU — Lumière'] > 0)].copy()
    if left.empty and right.empty:
        return ''

    from plotly.subplots import make_subplots

    fig = make_subplots(
        rows=1,
        cols=2,
        horizontal_spacing=0.12,
        subplot_titles=[
            'Festivais internacionais x VOD Europa',
            'Festivais internacionais x Lumière',
        ],
    )

    hover = (
        '<b>%{customdata[0]}</b><br>Ano: %{customdata[1]}<br>'
        'Pontuação festivais: %{x:.1f}<br>%{customdata[2]}<extra></extra>'
    )

    if not left.empty:
        fig.add_trace(go.Scatter(
            x=left['Pontuação Festivais'],
            y=left['VOD Intl — N Países'],
            customdata=np.column_stack([
                left['Projeto'],
                left['Ano'],
                ['Países VOD: ' + str(int(v)) + '<br>Plataformas: ' + str(int(p))
                 for v, p in zip(left['VOD Intl — N Países'], left['VOD Intl — N Plataformas'])]
            ]),
            mode='markers',
            name='VOD Europa',
            marker=dict(size=9, color='#5B6BB5', opacity=0.78, line=dict(color='#e8eaf2', width=0.4)),
            hovertemplate=hover,
        ), row=1, col=1)
        xs = np.linspace(left['Pontuação Festivais'].min(), left['Pontuação Festivais'].max(), 40)
        trend = _trend_trace(left['Pontuação Festivais'].to_numpy(), left['VOD Intl — N Países'].to_numpy(), xs, '#9aa8ff')
        if trend:
            fig.add_trace(trend, row=1, col=1)

    if not right.empty:
        fig.add_trace(go.Scatter(
            x=right['Pontuação Festivais'],
            y=right['Adm. EU — Lumière'],
            customdata=np.column_stack([
                right['Projeto'],
                right['Ano'],
                ['Admissões EU: ' + _fmt_num(v) for v in right['Adm. EU — Lumière']]
            ]),
            mode='markers',
            name='Lumière',
            marker=dict(size=9, color='#00e5c8', opacity=0.78, line=dict(color='#e8eaf2', width=0.4)),
            hovertemplate=hover,
        ), row=1, col=2)
        xs = np.linspace(right['Pontuação Festivais'].min(), right['Pontuação Festivais'].max(), 40)
        trend = _trend_trace(right['Pontuação Festivais'].to_numpy(), right['Adm. EU — Lumière'].to_numpy(), xs, '#65ffe6')
        if trend:
            fig.add_trace(trend, row=1, col=2)

    fig.update_layout(
        height=500,
        margin=dict(l=60, r=35, t=70, b=55),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=11),
        showlegend=False,
    )
    fig.update_xaxes(title_text='Pontuação em festivais internacionais', gridcolor='rgba(255,255,255,0.08)', zeroline=False)
    fig.update_yaxes(title_text='Nº de países VOD', gridcolor='rgba(255,255,255,0.08)', zeroline=False, row=1, col=1)
    fig.update_yaxes(title_text='Admissões EU (Lumière)', gridcolor='rgba(255,255,255,0.08)', zeroline=False, row=1, col=2)
    fig.add_annotation(text=_corr_label(df, 'Pontuação Festivais', 'VOD Intl — N Países'), xref='paper', yref='paper', x=0.22, y=1.08, showarrow=False, font=dict(size=10, color='#b7bdcc'))
    fig.add_annotation(text=_corr_label(df, 'Pontuação Festivais', 'Adm. EU — Lumière'), xref='paper', yref='paper', x=0.78, y=1.08, showarrow=False, font=dict(size=10, color='#b7bdcc'))

    fig_html = pio.to_html(fig, full_html=False, include_plotlyjs=False, div_id='festival-vod-lumiere-scatter')
    return f'''
<div class="card" style="margin:20px 0;background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:8px;padding:14px 16px">
  <div style="font-family:var(--font-head);font-size:17px;font-style:italic;color:var(--text);margin-bottom:5px">Dispersão: festivais internacionais x VOD e Lumière</div>
  <div style="font-size:11px;color:var(--muted);line-height:1.55;margin-bottom:12px">
    Cada ponto é uma obra com pontuação em festivais internacionais e presença em VOD Europa ou bilheteria europeia Lumière. As linhas tracejadas indicam a tendência linear entre os indicadores.
  </div>
  {fig_html}
</div>'''

cmp_raw  = read(SRC_CMP)
conc_raw = read(SRC_CONC)
cs_raw   = read(SRC_CS)
pr_raw   = read(SRC_PR)
_div_section = read(SRC_DIV)
_sp_section  = read(SRC_SP)

# CSS for CL/DIV/SP sections (dark theme overrides + utility classes)
_cl_css_path = os.path.join(BASE, 'cl_div_sp_css.txt')
_cl_div_sp_css = open(_cl_css_path, encoding='utf-8').read() if os.path.exists(_cl_css_path) else ''

print("Files loaded OK")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — COMPARATIVO
# Extract only visao-geral, categorias, clusters, ticket
# (financeiro, roi-dom, roi-intl, dispersao are dropped)
# ─────────────────────────────────────────────────────────────────────────────
cmp_ids = [
    'tab-visao-geral', 'tab-financeiro', 'tab-ret-intl', 'tab-roi-dom', 'tab-roi-intl',
    'tab-dispersao', 'tab-categorias', 'tab-clusters', 'tab-ticket',
]

cmp_css = extract_css(cmp_raw)

# Extract all tab divs (we keep them all but only surface some)
tab_names_all = ['visao-geral', 'financeiro', 'ret-intl', 'roi-dom', 'roi-intl', 'dispersao',
                 'categorias', 'clusters', 'ticket']
# Sub-tabs shown in Visão Geral: geral + retorno-dom + retorno-intl
# categorias → injected into CS section; clusters → injected into PR section
tab_names_vg = ['visao-geral', 'financeiro', 'ret-intl']

tab_divs = {}
for tname in tab_names_all:
    tid = f'tab-{tname}'
    extracted = extract_tab_panel(cmp_raw, tid)
    if extracted and 'not found' not in extracted:
        tab_divs[tname] = extracted
    else:
        print(f"WARNING: could not extract tab div for {tid}")
        tab_divs[tname] = f'<div id="{tid}" class="tab-content"><p>Content not found</p></div>'

# Normalize inline font-family in ret-intl HTML content
if 'ret-intl' in tab_divs:
    tab_divs['ret-intl'] = tab_divs['ret-intl'].replace(
        "font-family:'DM Sans',sans-serif", "font-family:var(--font-mono)")
    _intl_scatter_html = _build_festival_vod_lumiere_scatter_html()
    _intl_listing_html = _build_international_listing_html()
    _intl_lumiere_marker = '<div class="card"><div id="intl_lumiere"'
    if _intl_scatter_html and _intl_lumiere_marker in tab_divs['ret-intl']:
        tab_divs['ret-intl'] = tab_divs['ret-intl'].replace(
            _intl_lumiere_marker,
            _intl_scatter_html + '\n' + _intl_lumiere_marker,
            1,
        )
    _intl_insert_at = tab_divs['ret-intl'].rfind('</div>')
    if _intl_insert_at != -1:
        tab_divs['ret-intl'] = tab_divs['ret-intl'][:_intl_insert_at] + _intl_listing_html + tab_divs['ret-intl'][_intl_insert_at:]

cmp_scripts_raw = re.findall(r'<script>([\s\S]*?)</script>', cmp_raw)
cmp_scripts = '\n'.join(f'<script>\n{s}\n</script>' for s in cmp_scripts_raw)

tab_divs_prefixed = {}
for tname, div in tab_divs.items():
    tab_divs_prefixed[tname] = prefix_ids(div, cmp_ids, 'cmp-')

cmp_scripts_prefixed = prefix_ids(cmp_scripts, cmp_ids, 'cmp-')

cmp_scripts_prefixed = re.sub(
    r'function showTab\s*\([\s\S]*?(?=\nfunction |\nvar |\Z)',
    '',
    cmp_scripts_prefixed
)

_intl_map_data_match = re.search(r'var MAP_DATA\s*=\s*([\s\S]*?);\s*var CNAMES', cmp_scripts_prefixed)
_intl_cnames_match = re.search(r'var CNAMES\s*=\s*([\s\S]*?);\s*var gG', cmp_scripts_prefixed)
_intl_map_data_js = _intl_map_data_match.group(1) if _intl_map_data_match else '{}'
_intl_cnames_js = _intl_cnames_match.group(1) if _intl_cnames_match else '{}'
_iso2_to_iso3 = {
    "AL": "ALB", "AT": "AUT", "BA": "BIH", "BE": "BEL", "BG": "BGR",
    "BY": "BLR", "CH": "CHE", "CZ": "CZE", "DE": "DEU", "DK": "DNK",
    "EE": "EST", "ES": "ESP", "FI": "FIN", "FR": "FRA", "GB": "GBR",
    "GR": "GRC", "HR": "HRV", "HU": "HUN", "IE": "IRL", "IS": "ISL",
    "IT": "ITA", "LT": "LTU", "LU": "LUX", "LV": "LVA", "MD": "MDA",
    "ME": "MNE", "MK": "MKD", "MT": "MLT", "NL": "NLD", "NO": "NOR",
    "PL": "POL", "PT": "PRT", "RO": "ROU", "RS": "SRB", "SE": "SWE",
    "SI": "SVN", "SK": "SVK", "TR": "TUR", "UA": "UKR",
    "AR": "ARG", "AU": "AUS", "BF": "BFA", "BR": "BRA", "CA": "CAN",
    "CL": "CHL", "CN": "CHN", "CO": "COL", "CU": "CUB", "EG": "EGY",
    "HK": "HKG", "IN": "IND", "JP": "JPN", "KR": "KOR", "MA": "MAR",
    "MX": "MEX", "NZ": "NZL", "PY": "PRY", "RU": "RUS", "TW": "TWN",
    "US": "USA", "UY": "URY", "ZA": "ZAF",
}
intl_map_fallback_script = """
<script>
(function(){
  var MAP_DATA = __MAP_DATA__;
  var CNAMES = __CNAMES__;
  var ISO3 = __ISO3__;

  function _target() {
    var svg = document.getElementById('intl-map-svg');
    if(!svg) return null;
    var host = document.getElementById('intl-map-plotly');
    if(!host) {
      host = document.createElement('div');
      host.id = 'intl-map-plotly';
      host.style.width = '100%';
      host.style.height = '500px';
      svg.insertAdjacentElement('afterend', host);
    }
    svg.style.display = 'none';
    return host;
  }

  function _activeRows() {
    var showFest = !document.getElementById('chk-fest') || document.getElementById('chk-fest').checked;
    var showVod = !document.getElementById('chk-vod') || document.getElementById('chk-vod').checked;
    var rows = [];
    Object.keys(MAP_DATA || {}).forEach(function(k) {
      var info = MAP_DATA[k] || {};
      var iso2 = info.iso2;
      var iso3 = ISO3[iso2];
      if(!iso3) return;
      var fest = showFest ? (info.fest || []) : [];
      var vod = showVod ? (info.vod || []) : [];
      var value = fest.length + vod.length;
      if(value <= 0) return;
      var samples = [];
      if(fest.length) samples.push('Festivais:<br>' + fest.join('<br>'));
      if(vod.length) samples.push('VOD:<br>' + vod.join('<br>'));
      rows.push({
        iso2: iso2,
        iso3: iso3,
        name: CNAMES[iso2] || iso2,
        fest: fest.length,
        vod: vod.length,
        festList: fest,
        vodList: vod,
        value: value,
        sample: samples.join('<br>')
      });
    });
    return rows;
  }

  function _esc(s) {
    return String(s || '').replace(/[&<>"']/g, function(ch) {
      return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[ch];
    });
  }

  function _renderCountryDetail(row) {
    var box = document.getElementById('intl-country-detail');
    if(!box || !row) return;
    var fest = row.festList || [];
    var vod = row.vodList || [];
    var html = '<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:8px">';
    html += '<div><div style="font-size:12px;font-weight:800;color:#222">' + _esc(row.name) + ' (' + _esc(row.iso2) + ')</div>';
    html += '<div style="font-size:9px;color:#666;margin-top:2px">Fontes: festivais na base ATA BRDE/FSA 2024; VOD na base Lumière VOD.</div></div>';
    html += '<div style="font-size:10px;color:#555;white-space:nowrap"><b style="color:#E8702A">' + fest.length + '</b> obras em festival · <b style="color:#5B6BB5">' + vod.length + '</b> títulos VOD</div></div>';
    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;align-items:start">';
    html += '<div><div style="font-size:9px;font-weight:700;color:#E8702A;margin-bottom:5px;text-transform:uppercase;letter-spacing:.06em">Festivais</div>';
    html += fest.length ? fest.map(function(f){ return '<div style="padding:4px 0;border-bottom:1px solid #e3e5eb;line-height:1.25">' + _esc(f) + '</div>'; }).join('') : '<div style="color:#888">Sem obra identificada em festival neste país.</div>';
    html += '</div><div><div style="font-size:9px;font-weight:700;color:#5B6BB5;margin-bottom:5px;text-transform:uppercase;letter-spacing:.06em">VOD Europa</div>';
    html += vod.length ? vod.map(function(f){ return '<div style="padding:4px 0;border-bottom:1px solid #e3e5eb;line-height:1.25">' + _esc(f) + '</div>'; }).join('') : '<div style="color:#888">Sem título identificado em VOD neste país.</div>';
    html += '</div></div>';
    box.innerHTML = html;
  }

  function renderIntlMapFallback() {
    if(!window.Plotly) return;
    var host = _target();
    if(!host) return;
    var rows = _activeRows();
    var mode = (document.getElementById('chk-fest') && document.getElementById('chk-fest').checked && !(document.getElementById('chk-vod') && document.getElementById('chk-vod').checked))
      ? 'festivais'
      : ((document.getElementById('chk-vod') && document.getElementById('chk-vod').checked && !(document.getElementById('chk-fest') && document.getElementById('chk-fest').checked)) ? 'vod' : 'presenca');
    var scale = mode === 'festivais'
      ? [[0, '#2a1b12'], [0.35, '#E8702A'], [1, '#f5c842']]
      : [[0, '#151a2b'], [0.45, '#5B6BB5'], [1, '#00e5c8']];
    var data = [{
      type: 'choropleth',
      locationmode: 'ISO-3',
      locations: rows.map(function(r){ return r.iso3; }),
      z: rows.map(function(r){ return r.value; }),
      text: rows.map(function(r){ return r.name; }),
      customdata: rows.map(function(r){ return [r.iso2, r.fest, r.vod, r.sample, r.festList, r.vodList, r.name]; }),
      colorscale: scale,
      marker: {line: {color: '#0e1018', width: 0.7}},
      colorbar: {
        title: 'obras/títulos',
        tickfont: {color: '#e8eaf2'},
        titlefont: {color: '#e8eaf2'}
      },
      hovertemplate:
        '<b>%{text}</b> (%{customdata[0]})<br>' +
        'Festivais: %{customdata[1]}<br>' +
        'VOD Europa: %{customdata[2]}<br>' +
        '%{customdata[3]}<extra></extra>'
    }];
    var layout = {
      margin: {l: 0, r: 0, t: 6, b: 0},
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      font: {color: '#e8eaf2', family: 'DM Mono, monospace', size: 10},
      geo: {
        projection: {type: 'natural earth'},
        bgcolor: 'rgba(0,0,0,0)',
        showland: true,
        landcolor: '#111827',
        showcountries: true,
        countrycolor: '#2a2d42',
        lakecolor: '#07080f',
        coastlinecolor: '#2a2d42',
        fitbounds: 'locations'
      }
    };
    Plotly.react(host, data, layout, {responsive: true, displaylogo: false});
    host.on('plotly_click', function(ev) {
      if(!ev || !ev.points || !ev.points.length) return;
      var cd = ev.points[0].customdata || [];
      _renderCountryDetail({
        iso2: cd[0],
        fest: cd[1],
        vod: cd[2],
        festList: cd[4] || [],
        vodList: cd[5] || [],
        name: cd[6] || cd[0]
      });
    });
  }

  window.intlMapEnsure = renderIntlMapFallback;
  window.intlMapUpdate = renderIntlMapFallback;
  if(document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', renderIntlMapFallback);
  } else {
    renderIntlMapFallback();
  }
})();
</script>
""".replace('__MAP_DATA__', _intl_map_data_js).replace('__CNAMES__', _intl_cnames_js).replace('__ISO3__', json.dumps(_iso2_to_iso3, ensure_ascii=False))

print(f"Comparativo: {len(tab_divs)} tab divs extracted")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — CONCENTRACAO
# ─────────────────────────────────────────────────────────────────────────────
conc_ids = ['t1', 't2', 't3', 'hist_ticket', 'thr_curve', 'tier_comp', 'tier_pie',
            'tier_tbody', 'tier_bar_sust', 'lorenz_prod', 'scatter_tick']

conc_body = get_body(conc_raw)

conc_css = extract_css(conc_raw)

conc_scripts_all = re.findall(r'<script[^>]*>([\s\S]*?)</script>', conc_raw)
conc_data_script = conc_scripts_all[0] if len(conc_scripts_all) > 0 else ''
conc_logic_script = conc_scripts_all[1] if len(conc_scripts_all) > 1 else ''

conc_panel_t1 = extract_tab_panel(conc_body, 't1')
conc_panel_t2 = extract_tab_panel(conc_body, 't2')
conc_panel_t3 = extract_tab_panel(conc_body, 't3')

conc_panel_t1 = prefix_ids(conc_panel_t1, conc_ids, 'conc-').replace('class="tab-panel"', 'class="conc-panel"')
conc_panel_t2 = prefix_ids(conc_panel_t2, conc_ids, 'conc-').replace('class="tab-panel"', 'class="conc-panel"')
conc_panel_t3 = prefix_ids(conc_panel_t3, conc_ids, 'conc-').replace('class="tab-panel"', 'class="conc-panel"')

conc_data_script = prefix_ids(conc_data_script, conc_ids, 'conc-')
conc_logic_script = prefix_ids(conc_logic_script, conc_ids, 'conc-')

conc_logic_script = conc_logic_script.replace('function showTab(', 'function concShowTab_internal(')
conc_logic_script = re.sub(r'\bshowTab\(', 'concShowTab_internal(', conc_logic_script)

print("Concentracao: panels extracted")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — CRITERIO SELECAO
# ─────────────────────────────────────────────────────────────────────────────
cs_ids = [
    'canvas-quadrant', 'canvas-timeline', 'cat-desc-grid', 'cat-pills-rank',
    'filter-fase1', 'filter-obras-cat', 'kpi-bar', 'modal-content', 'modal-meta',
    'modal-overlay', 'modal-title', 'obras-count', 'obras-table-wrap',
    'quad-axis-btns', 'rank-count', 'rank-dom', 'rank-sort-btns', 'rank-sort-label',
    'rank-table-wrap', 'search-obras', 'search-rank', 'sint-table', 'sort-obras',
    'sort-rank', 'tab-comp', 'tab-meto', 'tab-obras', 'tab-rank', 'tab-sint',
    'timeline-legend', 'tooltip',
]

cs_body = get_body(cs_raw)

tabs_start = cs_body.find('<div class="tabs">')
if tabs_start >= 0:
    cs_body = cs_body[tabs_start:]

cs_scripts_all = re.findall(r'<script[^>]*>([\s\S]*?)</script>', cs_raw)
cs_script = cs_scripts_all[-1] if cs_scripts_all else ''

cs_body = re.sub(r'<script[^>]*>[\s\S]*?</script>', '', cs_body)
cs_body = re.sub(r'<style[^>]*>[\s\S]*?</style>', '', cs_body)

cs_body = prefix_ids(cs_body, cs_ids, 'cs-')
cs_script = prefix_ids(cs_script, cs_ids, 'cs-')

# prefix_ids não cobre string literals em argumentos de função
cs_script = cs_script.replace("renderRank('rank-dom'", "renderRank('cs-rank-dom'")
cs_script = cs_script.replace("buildPills('cat-pills-rank'", "buildPills('cs-cat-pills-rank'")

cs_script = cs_script.replace('function showTab(', 'function csShowTab(')
cs_script = re.sub(r'(?<![a-zA-Z])showTab\(', 'csShowTab(', cs_script)
cs_body = cs_body.replace("onclick=\"showTab(", "onclick=\"csShowTab(")

# Fix: prefix ID string literals in onclick args and function body comparisons
# (not caught by prefix_ids which only handles id="", getElementById(), querySelector())
for _tab_id in ['tab-comp', 'tab-rank', 'tab-obras', 'tab-meto', 'tab-sint']:
    cs_body = cs_body.replace(f"csShowTab('{_tab_id}',", f"csShowTab('cs-{_tab_id}',")
    cs_script = cs_script.replace(f"==='{_tab_id}'", f"==='cs-{_tab_id}'")


cs_css = extract_css(cs_raw)

print("Criterio selecao: body extracted")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — PRODUTORAS
# ─────────────────────────────────────────────────────────────────────────────
pr_ids = [
    'bar-chart', 'canvas-wrap', 'cl-grid', 'cl-panel', 'cl-panel-dot',
    'cl-panel-title', 'cl-prod-list',
    'film-cl-chips', 'film-clr', 'film-count', 'film-crn', 'film-search',
    'film-sort', 'film-tbody', 'film-thead',
    'kpi-bar', 'obras-drawer', 'obras-drawer-tbody', 'obras-drawer-thead',
    'obras-drawer-title', 'ov-clr', 'ov-count', 'ov-search', 'ov-sort',
    'prod-data', 'quad-canvas', 'quad-size', 'quad-size-legend',
    'r-clr', 'r-search', 'rank-chips', 'rank-count', 'rank-sort',
    'rank-tbody', 'rank-thead', 's-clr', 's-lin', 's-log',
    'scatter-plotly', 'sx', 'sy',
    'tab0', 'tab1', 'tab2', 'tab3', 'tip',
]

pr_body = get_body(pr_raw)

pr_body_stripped = re.sub(r'<div class="hdr">[\s\S]*?</div>', '', pr_body, count=1)
tab_bar_start = pr_body_stripped.find('<div class="tab-bar">')
if tab_bar_start > 0:
    pr_body = pr_body_stripped[tab_bar_start:]
else:
    pr_body = pr_body_stripped

pr_scripts = re.findall(r'<script([^>]*)>([\s\S]*?)</script>', pr_raw)
pr_data_script = ''
pr_logic_script = ''
for attrs, content in pr_scripts:
    if 'prod-data' in attrs:
        pr_data_script = content
    else:
        pr_logic_script = content

pr_body = re.sub(r'<script[^>]*>[\s\S]*?</script>', '', pr_body)

pr_body = prefix_ids(pr_body, pr_ids, 'pr-')
pr_data_script = prefix_ids(pr_data_script, pr_ids, 'pr-')
pr_logic_script = prefix_ids(pr_logic_script, pr_ids, 'pr-')

pr_logic_script = pr_logic_script.replace('function switchTab(', 'function prSwitchTab(')
pr_logic_script = re.sub(r'(?<![a-zA-Z])switchTab\(', 'prSwitchTab(', pr_logic_script)
pr_body = pr_body.replace('onclick="switchTab(', 'onclick="prSwitchTab(')
pr_body = pr_body.replace("onclick='switchTab(", "onclick='prSwitchTab(")

# Add "Por Cluster" init + hide extra panels when any standard tab is clicked
pr_logic_script = pr_logic_script.replace(
    'if(i===3) renderFilms();',
    'if(i===1) setTimeout(function(){_initFigsInContainer("pr-tab-clusters");},80);\n  if(i===3) renderFilms();\n  ["pr-ticket-panel","pr-conc-panel"].forEach(function(id){var p=document.getElementById(id);if(p){p.style.display="none";p.classList.remove("active");}});'
)

pr_css = extract_css(pr_raw)

# ─── POST-PROCESSING: Normalize CSS across all sources ───

# 1+2. Remove duplicate global rules (:root, *, body) — mega :root is authoritative
def _strip_global_rules(css):
    css = re.sub(r':root\s*\{[^}]*\}', '', css)
    css = re.sub(r'\*\s*\{[^}]*box-sizing[^}]*\}', '', css)
    css = re.sub(r'(?:html,)?body\s*\{[^}]*\}', '', css)
    return css

cmp_css = _strip_global_rules(cmp_css)
cs_css = _strip_global_rules(cs_css)
pr_css = _strip_global_rules(pr_css)
conc_css = _strip_global_rules(conc_css)

# 3. Normalize shorthand CSS variables → base theme names
def _normalize_layout_vars(css):
    """Map layout variables (surfaces, text) to base theme."""
    css = css.replace('var(--s1)', 'var(--surface)')
    css = css.replace('var(--s2)', 'var(--surface2)')
    css = css.replace('var(--s3)', 'var(--border)')
    css = css.replace('var(--s4)', 'var(--surface2)')
    css = css.replace('var(--txt)', 'var(--text)')
    return css

# PR: --acc is primary accent (cyan)
pr_css = _normalize_layout_vars(pr_css)
pr_css = pr_css.replace('var(--acc2)', 'var(--gold)')
pr_css = pr_css.replace('var(--acc)', 'var(--accent)')

# Conc: --acc is warning/orange, --acc3 is OK/green, --acc4 is info/blue
conc_css = _normalize_layout_vars(conc_css)
conc_css = conc_css.replace('var(--acc4)', 'var(--muted-blue)')
conc_css = conc_css.replace('var(--acc3)', 'var(--accent)')
conc_css = conc_css.replace('var(--acc2)', 'var(--gold)')
conc_css = conc_css.replace('var(--acc)', 'var(--coral)')

# 4. Replace hardcoded fonts with CSS variables
def _normalize_fonts(css):
    css = css.replace("font-family:'Syne',sans-serif", "font-family:var(--font-head)")
    css = css.replace("font-family: 'Syne', sans-serif", "font-family: var(--font-head)")
    css = css.replace("font-family:'Segoe UI',sans-serif", "font-family:var(--font-mono)")
    css = css.replace("font-family: 'Segoe UI', sans-serif", "font-family: var(--font-mono)")
    css = css.replace("font-family:'DM Sans',sans-serif", "font-family:var(--font-mono)")
    css = css.replace("font-family: 'DM Sans', sans-serif", "font-family: var(--font-mono)")
    return css

cmp_css = _normalize_fonts(cmp_css)
cs_css = _normalize_fonts(cs_css)
pr_css = _normalize_fonts(pr_css)
conc_css = _normalize_fonts(conc_css)

# 5. Scope conc_css rules to #conc-section to avoid leaking into other sections
def _scope_css(css, scope):
    """Prefix every CSS selector with a scope selector."""
    result = []
    i = 0
    while i < len(css):
        # Skip whitespace
        if css[i] in ' \t\n\r':
            result.append(css[i])
            i += 1
            continue
        # Skip comments
        if css[i:i+2] == '/*':
            end = css.find('*/', i+2)
            if end == -1: end = len(css)
            else: end += 2
            result.append(css[i:end])
            i = end
            continue
        # Skip @rules
        if css[i] == '@':
            brace = css.find('{', i)
            if brace == -1:
                result.append(css[i:])
                break
            result.append(css[i:brace+1])
            # Find matching closing brace
            depth = 1
            j = brace + 1
            while j < len(css) and depth > 0:
                if css[j] == '{': depth += 1
                elif css[j] == '}': depth -= 1
                j += 1
            result.append(css[brace+1:j])
            i = j
            continue
        # Find next { — everything before it is the selector
        brace = css.find('{', i)
        if brace == -1:
            result.append(css[i:])
            break
        selector = css[i:brace].strip()
        if selector:
            # Prefix each comma-separated selector
            parts = selector.split(',')
            scoped = ', '.join(f'{scope} {p.strip()}' for p in parts if p.strip())
            result.append(scoped)
        result.append('{')
        # Find matching closing brace
        depth = 1
        j = brace + 1
        while j < len(css) and depth > 0:
            if css[j] == '{': depth += 1
            elif css[j] == '}': depth -= 1
            j += 1
        result.append(css[brace+1:j])
        i = j
    return ''.join(result)

conc_css_scoped = _scope_css(conc_css, '#conc-section')

print("Produtoras: body extracted")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — Build Concentracao section (for Ticket panel inside Produtoras)
# ─────────────────────────────────────────────────────────────────────────────
conc_section = f'''
<div id="conc-section" style="padding:20px 0;border-top:1px solid var(--border);margin-top:24px">
  <div style="font-family:'DM Serif Display',serif;font-size:14px;font-style:italic;color:var(--accent);margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid var(--border)">
    Concentração e Distribuição de Produtoras
  </div>
  <div class="conc-subnav" style="display:flex;gap:0;border-bottom:1px solid var(--border);margin-bottom:20px">
    <button class="conc-tab active" onclick="concShow('conc-t1')">Ticket Médio Anual</button>
    <button class="conc-tab" onclick="concShow('conc-t2')">Análise por Tiers</button>
    <button class="conc-tab" onclick="concShow('conc-t3')">Concentração Lorenz</button>
  </div>
  {conc_panel_t1}
  {conc_panel_t2}
  {conc_panel_t3}
</div>
'''

conc_section = conc_section.replace('id="conc-t2"', 'id="conc-t2" style="display:none"')
conc_section = conc_section.replace('id="conc-t3"', 'id="conc-t3" style="display:none"')

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — Build Ticket panel for Produtoras
# ticket content is in tab_divs_prefixed['ticket'] (from cmp section)
# ─────────────────────────────────────────────────────────────────────────────
ticket_div = tab_divs_prefixed.get('ticket', '<p>Ticket não encontrado</p>')

# This panel will be added as a new tab in the produtoras section
pr_ticket_panel = f'''<div id="pr-ticket-panel" class="tab-panel" style="overflow-y:auto;padding:14px 18px 18px">
<div style="font-family:\'DM Serif Display\',serif;font-size:15px;font-style:italic;color:var(--accent);margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid var(--border)">
  Ticket Médio &amp; Capital Parado
</div>
{ticket_div}
</div>'''

# Concentração panel — aba separada
pr_conc_panel = f'''<div id="pr-conc-panel" class="tab-panel" style="overflow-y:auto;padding:14px 18px 18px">
{conc_section}
</div>'''

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6b — Build "Por Cluster" panel from PROD data (same source as Visão Geral)
# ─────────────────────────────────────────────────────────────────────────────
# Extract PROD JSON embedded in painel_produtoras.html
_prod_match = re.search(r'<script[^>]*id="prod-data"[^>]*>([\s\S]*?)</script>', pr_raw)
_prod_records = []
if _prod_match:
    _json_match = re.search(r'const PROD\s*=\s*(\[[\s\S]*?\]);', _prod_match.group(1))
    if _json_match:
        _prod_records = json.loads(_json_match.group(1))
        print(f"  PROD extraído: {len(_prod_records)} produtoras")
    else:
        print("  AVISO: PROD JSON não encontrado no script prod-data")
else:
    print("  AVISO: script prod-data não encontrado em painel_produtoras.html")

# Build cluster figures from PROD data
_CL_MAP = {
    'duplo':       ('Duplo Retorno',         '#f5c842'),
    'dom':         ('Retorno Doméstico',      '#f09020'),
    'intl':        ('Retorno Internacional',  '#4fa3e0'),
    'sem_retorno': ('Fomento Baixo Retorno',  '#e05050'),
    'pequeno':     ('Pequeno Porte',          '#4a4a60'),
}
_CL_ORDER = ['duplo', 'dom', 'intl', 'sem_retorno', 'pequeno']
_CL_LABELS = [_CL_MAP[k][0] for k in _CL_ORDER]
_CL_COLORS = [_CL_MAP[k][1] for k in _CL_ORDER]

_df_pr = pd.DataFrame(_prod_records) if _prod_records else pd.DataFrame()

def _cl_stats(key):
    if _df_pr.empty:
        return {'n_prod':0,'n_obras':0,'inv_bi':0,'rec_bi':0,'roi_agr':0,'roi_dom_med':0,'roi_intl_med':0,'roi_intl_soma':0,'pct_intl':0}
    g = _df_pr[_df_pr['cl'] == key]
    if len(g) == 0:
        return {'n_prod':0,'n_obras':0,'inv_bi':0,'rec_bi':0,'roi_agr':0,'roi_dom_med':0,'roi_intl_med':0,'roi_intl_soma':0,'pct_intl':0}
    inv = g['inv_def'].sum()
    rec = g['rec_def'].sum()
    inv_w = g['inv_def']
    roi_dom_vals = g[g['rda'] > 0]
    roi_dom_w = (roi_dom_vals['rda'] * roi_dom_vals['inv_def']).sum() / roi_dom_vals['inv_def'].sum() if roi_dom_vals['inv_def'].sum() > 0 else 0
    roi_intl_vals = g[g['ria'] > 0]['ria']
    return {
        'n_prod':      len(g),
        'n_obras':     int(g['n'].sum()),
        'inv_bi':      inv / 1e9,
        'rec_bi':      rec / 1e9,
        'roi_agr':     rec / inv if inv > 0 else 0,
        'roi_dom_med': roi_dom_w,
        'roi_intl_med': roi_intl_vals.mean() if len(roi_intl_vals) > 0 else 0,
        'roi_intl_soma': roi_intl_vals.sum() if len(roi_intl_vals) > 0 else 0,
        'pct_intl':    (g['rim'] >= 13).mean() * 100,
    }

_cl_resumo = {k: _cl_stats(k) for k in _CL_ORDER}

_LAYOUT_BASE = dict(
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(20,20,20,0.5)',
    font=dict(color='#ddd8cc', family='DM Mono, monospace', size=11),
    margin=dict(l=50, r=20, t=80, b=120),
)

def _make_cl_kpi():
    from plotly.subplots import make_subplots as _msp
    fig = _msp(rows=1, cols=4,
               subplot_titles=['Nº de produtoras','Nº de obras','Investimento (R$ bi)','Receita (R$ bi)'],
               horizontal_spacing=0.08)
    cores = _CL_COLORS
    for vals, col, fmt in [
        ([_cl_resumo[k]['n_prod']  for k in _CL_ORDER], 1, lambda v: str(int(v))),
        ([_cl_resumo[k]['n_obras'] for k in _CL_ORDER], 2, lambda v: str(int(v))),
        ([_cl_resumo[k]['inv_bi']  for k in _CL_ORDER], 3, lambda v: f'R$ {v:.2f}bi'),
        ([_cl_resumo[k]['rec_bi']  for k in _CL_ORDER], 4, lambda v: f'R$ {v:.2f}bi'),
    ]:
        fig.add_trace(go.Bar(x=_CL_LABELS, y=vals, marker_color=cores,
                             text=[fmt(v) for v in vals], textposition='outside', showlegend=False),
                      row=1, col=col)
    fig.update_layout(title='KPIs por Cluster de Produtora', height=460,
                      **{k:v for k,v in _LAYOUT_BASE.items()})
    fig.update_xaxes(tickangle=-25)
    return fig

def _make_cl_roi():
    from plotly.subplots import make_subplots as _msp
    fig = _msp(rows=1, cols=3,
               subplot_titles=['ROI agregado (rec/inv)','ROI dom. pond. (rda)','ROI intl médio (ria)'],
               horizontal_spacing=0.1)
    for vals, col, fmt in [
        ([_cl_resumo[k]['roi_agr']      for k in _CL_ORDER], 1, lambda v: f'{v:.2f}x'),
        ([_cl_resumo[k]['roi_dom_med']  for k in _CL_ORDER], 2, lambda v: f'{v:.2f}x'),
        ([_cl_resumo[k]['roi_intl_med'] for k in _CL_ORDER], 3, lambda v: f'{v:.2f}'),
    ]:
        fig.add_trace(go.Bar(x=_CL_LABELS, y=vals, marker_color=_CL_COLORS,
                             text=[fmt(v) for v in vals], textposition='outside', showlegend=False),
                      row=1, col=col)
    fig.update_layout(title='Retorno por Cluster de Produtora', height=460,
                      **{k:v for k,v in _LAYOUT_BASE.items()})
    fig.update_xaxes(tickangle=-25)
    return fig

def _make_cl_box():
    fig = go.Figure()
    if not _df_pr.empty:
        for k in _CL_ORDER:
            g = _df_pr[(_df_pr['cl'] == k) & (_df_pr['rda'] > 0)]['rda']
            if len(g) < 3: continue
            fig.add_trace(go.Box(y=g, name=_CL_MAP[k][0], marker_color=_CL_MAP[k][1], boxmean='sd'))
    fig.update_layout(title='Distribuição do ROI Doméstico por Cluster<br><sup>(produtoras com receita · escala log)</sup>',
                      yaxis=dict(title='ROI doméstico (rda)', type='log'),
                      height=460, showlegend=False, **{k:v for k,v in _LAYOUT_BASE.items()})
    return fig

def _make_cl_box_intl():
    fig = go.Figure()
    if not _df_pr.empty:
        for k in _CL_ORDER:
            g = _df_pr[(_df_pr['cl'] == k) & (_df_pr['ria'] > 0)]['ria']
            if len(g) < 3: continue
            fig.add_trace(go.Box(y=g, name=_CL_MAP[k][0], marker_color=_CL_MAP[k][1], boxmean='sd'))
    fig.update_layout(title='Distribuição do ROI Internacional por Cluster<br><sup>(produtoras com ROI Intl ≥ 13)</sup>',
                      yaxis=dict(title='ROI intl médio (ria, 0–100)', type='log'),
                      height=460, showlegend=False, **{k:v for k,v in _LAYOUT_BASE.items()})
    return fig

def _make_cl_scat():
    fig = go.Figure()
    if not _df_pr.empty:
        for k in _CL_ORDER:
            g = _df_pr[(_df_pr['cl'] == k) & (_df_pr['inv_def'] > 0) & (_df_pr['rda'] > 0)]
            if len(g) == 0: continue
            sz = (g['n'].clip(1)**0.5 * 4).clip(5, 25)
            fig.add_trace(go.Scatter(
                x=g['inv_def'] / 1e6, y=g['rda'], mode='markers', name=_CL_MAP[k][0],
                marker=dict(color=_CL_MAP[k][1], size=sz, opacity=0.65, line=dict(width=0.5, color='white')),
                text=g['nm'],
                hovertemplate='<b>%{text}</b><br>Inv: R$ %{x:.1f}mi<br>ROI dom: %{y:.2f}x<extra></extra>',
            ))
    fig.update_layout(title='Investimento × ROI Doméstico<br><sup>(tamanho = nº obras · eixos log)</sup>',
                      xaxis=dict(title='Investimento (R$ mi, log)', type='log'),
                      yaxis=dict(title='ROI doméstico (log)', type='log'),
                      height=560, legend=dict(orientation='h', y=-0.15),
                      **{k:v for k,v in _LAYOUT_BASE.items()})
    return fig

def _make_cl_scat_intl():
    fig = go.Figure()
    if not _df_pr.empty:
        for k in _CL_ORDER:
            g = _df_pr[(_df_pr['cl'] == k) & (_df_pr['inv_def'] > 0) & (_df_pr['ria'] > 0)]
            if len(g) == 0: continue
            sz = (g['n'].clip(1)**0.5 * 4).clip(5, 25)
            fig.add_trace(go.Scatter(
                x=g['inv_def'] / 1e6, y=g['ria'], mode='markers', name=_CL_MAP[k][0],
                marker=dict(color=_CL_MAP[k][1], size=sz, opacity=0.65, line=dict(width=0.5, color='white')),
                text=g['nm'],
                hovertemplate='<b>%{text}</b><br>Inv: R$ %{x:.1f}mi<br>ROI intl: %{y:.1f}<extra></extra>',
            ))
    fig.update_layout(title='Investimento × ROI Internacional<br><sup>(tamanho = nº obras · eixo x log)</sup>',
                      xaxis=dict(title='Investimento (R$ mi, log)', type='log'),
                      yaxis=dict(title='ROI intl médio (ria, 0–100)'),
                      height=560, legend=dict(orientation='h', y=-0.15),
                      **{k:v for k,v in _LAYOUT_BASE.items()})
    return fig

def _make_cl_tab():
    cats = _CL_LABELS
    rows = [
        ('Nº Produtoras',      [str(_cl_resumo[k]['n_prod'])                          for k in _CL_ORDER]),
        ('Nº Obras',           [str(_cl_resumo[k]['n_obras'])                         for k in _CL_ORDER]),
        ('Invest. (R$ bi)',    [f"R$ {_cl_resumo[k]['inv_bi']:.2f}bi"                 for k in _CL_ORDER]),
        ('Receita (R$ bi)',    [f"R$ {_cl_resumo[k]['rec_bi']:.2f}bi"                 for k in _CL_ORDER]),
        ('ROI Agregado',       [f"{_cl_resumo[k]['roi_agr']:.2f}x"                    for k in _CL_ORDER]),
        ('ROI Dom. Pond.',     [f"{_cl_resumo[k]['roi_dom_med']:.2f}x"                for k in _CL_ORDER]),
        ('ROI Intl Médio',     [f"{_cl_resumo[k]['roi_intl_med']:.2f}"                for k in _CL_ORDER]),
        ('% c/ ROI Intl >=13', [f"{_cl_resumo[k]['pct_intl']:.0f}%"                  for k in _CL_ORDER]),
    ]
    fig = go.Figure(go.Table(
        header=dict(values=['Métrica'] + cats, fill_color='#1c1c1c', font=dict(color='#ddd8cc', size=10),
                    line_color='#363636', align='left'),
        cells=dict(values=[[r[0] for r in rows]] + [[r[1][i] for r in rows] for i in range(len(cats))],
                   fill_color='#141414', font=dict(color='#ddd8cc', size=10),
                   line_color='#272727', align='left'),
    ))
    fig.update_layout(title='Resumo por Cluster — Fonte: Visão Geral Produtoras', height=340,
                      **{k:v for k,v in _LAYOUT_BASE.items()})
    return fig

def _to_lazy_pr(fig, div_id, height, search=None):
    fig_json = fig.to_json()
    search_html = ''
    data_attr = ''
    if search:
        data_attr = f' data-search-mode="{search}"'
        placeholder = 'Buscar por produtora…'
        search_html = (f'<div class="search-wrap">'
                       f'<input type="text" class="search-input" placeholder="{placeholder}" '
                       f'oninput="filterFig(\'{div_id}\', this.value)">'
                       f'<span class="search-count" id="count_{div_id}"></span>'
                       f'</div>')
    return (f'{search_html}'
            f'<div id="{div_id}"{data_attr} style="width:100%;height:{height}px;"></div>\n'
            f'<script>window.__fig_{div_id} = {fig_json};</script>')

_fig_cl_kpi      = _make_cl_kpi()
_fig_cl_roi      = _make_cl_roi()
_fig_cl_box      = _make_cl_box()
_fig_cl_box_intl = _make_cl_box_intl()
_fig_cl_scat     = _make_cl_scat()
_fig_cl_scat_intl= _make_cl_scat_intl()
_fig_cl_tab      = _make_cl_tab()

clusters_content = f'''<div id="pr-tab-clusters" class="tab-panel main-tab" style="overflow-y:auto;padding:20px 24px">
  <div class="card">{_to_lazy_pr(_fig_cl_scat,      "pr_cl_scat",      580, search="text")}</div>
  <div class="card">{_to_lazy_pr(_fig_cl_kpi,       "pr_cl_kpi",       460)}</div>
  <div class="card">{_to_lazy_pr(_fig_cl_roi,       "pr_cl_roi",       460)}</div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
    <div class="card">{_to_lazy_pr(_fig_cl_box,     "pr_cl_box",       460)}</div>
    <div class="card">{_to_lazy_pr(_fig_cl_box_intl,"pr_cl_box_intl",  460)}</div>
  </div>
  <div class="card">{_to_lazy_pr(_fig_cl_scat_intl, "pr_cl_scat_intl", 560, search="text")}</div>
  <div class="card">{_to_lazy_pr(_fig_cl_tab,       "pr_cl_tab",       340)}</div>
</div>'''
print("Por Cluster: figuras geradas a partir do PROD (mesma fonte da Visão Geral)")

# Inject Por Cluster button between Visão Geral and Ranking (no Ticket/Conc buttons)
pr_body = pr_body.replace(
    '>Vis\u00e3o Geral</button>\n  <button class="tab-btn" onclick="prSwitchTab(2)">Ranking',
    '>Vis\u00e3o Geral</button>\n  <button class="tab-btn" onclick="prSwitchTab(1)">Por Cluster</button>\n  <button class="tab-btn" onclick="prSwitchTab(2)">Ranking'
)

# Insert clusters panel before pr-tab1 (Ranking) so DOM order = [tab0, pr-tab-clusters, tab1, tab2]
pr_body = pr_body.replace(
    '<div class="tab-panel main-tab" id="pr-tab1">',
    clusters_content + '\n<div class="tab-panel main-tab" id="pr-tab1">'
)

# Append ticket and conc panels (needed for sidebar Concentração via prShowConc, no buttons)
pr_body = pr_body.rstrip()
pr_body = pr_body + '\n' + pr_ticket_panel + '\n' + pr_conc_panel

print("Por Cluster + Ticket + Concentração panels built for Produtoras")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7 — Build CMP (Visão Geral) section HTML
# Sub-tabs: visao-geral + financeiro only
# (categorias → CS section; clusters → PR section)
# ─────────────────────────────────────────────────────────────────────────────

cmp_subnav = '''
<div class="vg-subnav">
  <button class="vg-subtab active" onclick="vgShow('visao-geral')">Visão Geral</button>
  <button class="vg-subtab" onclick="vgShow('financeiro')">Retorno Doméstico</button>
  <button class="vg-subtab" onclick="vgShow('ret-intl')">Retorno Internacional</button>
</div>
'''

cmp_panels_html = ''
for tname in tab_names_vg:
    div = tab_divs_prefixed[tname]
    display = '' if tname == 'visao-geral' else 'display:none;'
    panel_html = f'<div id="cmp-panel-{tname}" class="cmp-tab-panel" style="{display}">\n{div}\n</div>\n'
    cmp_panels_html += panel_html

cmp_section_html = f'<div id="mega-section-cmp" class="mega-panel active">\n{cmp_subnav}\n{cmp_panels_html}\n</div>'

print("CMP (Visão Geral) section built")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 8 — Build CS section HTML
# ─────────────────────────────────────────────────────────────────────────────
cs_section_html = f'''<div id="mega-section-cs" class="mega-panel" style="display:none">
{cs_body}
</div>'''


# ─────────────────────────────────────────────────────────────────────────────
# STEP 9 — Build PR section HTML
# ─────────────────────────────────────────────────────────────────────────────
pr_section_html = f'''<div id="mega-section-pr" class="mega-panel" style="display:none">
{pr_body}
</div>'''

def _build_curtas_longas_html():
    cl_xlsx = os.path.join(BASE, 'dados', 'curtas_brasileiros_festivais_internacionais.xlsx')
    df_curtas = pd.read_excel(cl_xlsx, sheet_name='Dados')
    df_fest = pd.read_excel(cl_xlsx, sheet_name='Por Festival')
    df_prem = pd.read_excel(cl_xlsx, sheet_name='Premiados')

    def _norm_dir(s):
        s = str(s).lower().strip()
        s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        s = re.sub(r'\b(filho|jr\.?|junior)\b', '', s)
        s = re.sub(r'[^a-z0-9]+', ' ', s)
        return re.sub(r'\s+', ' ', s).strip()

    def _norm_person_proxy(s):
        s = str(s).lower().strip()
        s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        s = re.sub(r'\b(jr\.?|junior)\b', '', s)
        s = re.sub(r'[^a-z0-9]+', ' ', s)
        return re.sub(r'\s+', ' ', s).strip()

    cl_alias_groups = {
        'andre novais': ['andre novais', 'andre novais oliveira', 'andre de novais oliveira'],
        'thais fujinaga': ['thais fujinaga'],
        'marcelo caetano': ['marcelo caetano', 'marcelo cateano', 'marcelo batista caetano'],
        'gabriel martins': ['gabriel martins', 'gabriel martins alves'],
        'guto parente': ['guto parente', 'gustavo parente', 'gustavo parente lima'],
        'eva randolph': ['eva randolph'],
        'carolina markowicz': ['carolina markowicz', 'carolina markowicz bastos'],
    }
    cl_alias_norm = {k: [_norm_dir(v) for v in vals] for k, vals in cl_alias_groups.items()}

    def _proxy_dir_key(s):
        n = _norm_dir(s)
        if not n:
            return ''
        for canonical, aliases in cl_alias_norm.items():
            for alias in aliases:
                if n == alias or n.startswith(alias + ' ') or alias in n:
                    return canonical
        return n

    def _split_dirs(s):
        parts = re.split(r'\s+e\s+|\s*&\s*|\s*/\s*|,', str(s))
        return [p.strip() for p in parts if p.strip()]

    def _norm_title(s):
        s = str(s).lower().strip()
        s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        s = re.sub(r'[^a-z0-9]+', ' ', s)
        return re.sub(r'\s+', ' ', s).strip()

    def _norm_cpb(s):
        if pd.isna(s):
            return ''
        return str(s).strip().upper()

    def _year_from_cpb(s):
        m = re.match(r'^B(\d{2})', _norm_cpb(s))
        if not m:
            return np.nan
        yy = int(m.group(1))
        return 2000 + yy if yy <= 40 else 1900 + yy

    def _pick_col(df, wanted_norm):
        for col in df.columns:
            if _norm_title(col) == wanted_norm:
                return col
        return None

    def _join_unique(values, limit=None):
        out = sorted({
            str(v).strip()
            for v in values
            if str(v).strip() and str(v).strip().lower() != 'nan'
        })
        if limit:
            out = out[:limit]
        return ' | '.join(out)

    def _cell_text(value):
        if pd.isna(value):
            return ''
        text = str(value).strip()
        return '' if text.lower() == 'nan' else text

    def _person_tokens(s):
        particles = {'de', 'da', 'do', 'das', 'dos', 'e'}
        return [t for t in _norm_person_proxy(s).split() if t and t not in particles]

    def _is_artist_name_expansion(artist_name, official_name):
        artist = _person_tokens(artist_name)
        official = _person_tokens(official_name)
        if len(artist) < 2 or len(official) < 2 or len(artist) > len(official):
            return False
        if artist[0] != official[0]:
            return False
        pos = []
        cursor = 0
        for token in artist:
            try:
                idx = official.index(token, cursor)
            except ValueError:
                return False
            pos.append(idx)
            cursor = idx + 1
        if artist[-1] == official[-1]:
            return True
        # Handles cases like Carolina Markowicz Bastos and Felipe Sholl Machado:
        # the public/artist name is a prefix or ordered subset of the official name.
        return len(artist) == 2 and pos[-1] <= 2

    df_curtas['ano'] = pd.to_numeric(df_curtas['ano'], errors='coerce')
    df_curtas['premiado'] = df_curtas['premiado'].fillna(False).astype(bool)
    cl_rows = []
    for _, row in df_curtas.iterrows():
        for diretor in _split_dirs(row['diretor']):
            item = row.to_dict()
            item['diretor_individual'] = diretor
            item['diretor_norm'] = _proxy_dir_key(diretor)
            item['fonte_curta'] = 'excel_curtas'
            item['selecao_contagem'] = 1
            cl_rows.append(item)
    df_curtas_exp = pd.DataFrame(cl_rows)

    longas_frames = []
    part_path = os.path.join(BASE, 'dados', 'participacoes_festivais_diretores.csv')
    if os.path.exists(part_path):
        df_part = pd.read_csv(part_path)
        part = pd.DataFrame({
            'diretor': df_part.get('DIRETOR', ''),
            'diretor_norm': df_part.get('DIRETOR', df_part.get('DIRETOR_NORM', '')).map(_proxy_dir_key),
            'FILME': df_part.get('FILME', ''),
            'ANO_FILME': pd.to_numeric(df_part.get('ANO_FILME', np.nan), errors='coerce'),
            'CPB': df_part.get('CPB', ''),
            'FESTIVAL': df_part.get('FESTIVAL', ''),
            'FONTE_CRUZAMENTO': 'participacoes_festivais_diretores.csv',
        })
        longas_frames.append(part)

    ata_path = os.path.join(BASE, 'resultados', 'dataset', 'base_festivais_obras_ata.csv')
    raw_dir_path = os.path.join(BASE, 'raw', 'diretores-de-obras-nao-publicitarias-brasileiras.csv')
    cons_path = os.path.join(BASE, 'resultados', 'tabela_consolidada_obras.xlsx')
    if os.path.exists(ata_path) and os.path.exists(raw_dir_path):
        df_ata = pd.read_csv(ata_path, dtype=str)
        df_raw_dir = pd.read_csv(raw_dir_path, sep=';', dtype=str, on_bad_lines='skip')
        df_ata['cpb_norm'] = df_ata.get('cpb', '').map(_norm_cpb)
        df_ata['titulo_norm_join'] = df_ata.get('titulo_norm', df_ata.get('titulo', '')).map(_norm_title)
        df_raw_dir['cpb_norm_raw'] = df_raw_dir.get('CPB', '').map(_norm_cpb)
        df_raw_dir['titulo_norm_join'] = df_raw_dir.get('TITULO_ORIGINAL', '').map(_norm_title)

        cons_cpb = pd.DataFrame()
        cons_title = pd.DataFrame()
        if os.path.exists(cons_path):
            df_cons = pd.read_excel(cons_path, sheet_name='Obras', dtype=str)
            cons_cpb_col = _pick_col(df_cons, 'cpb')
            cons_tit_col = _pick_col(df_cons, 'projeto')
            cons_ano_col = _pick_col(df_cons, 'ano')
            if cons_cpb_col and cons_ano_col:
                cons_cpb = df_cons[[cons_cpb_col, cons_ano_col]].copy()
                cons_cpb.columns = ['cpb_norm_raw', 'ANO_CONS']
                cons_cpb['cpb_norm_raw'] = cons_cpb['cpb_norm_raw'].map(_norm_cpb)
                cons_cpb['ANO_CONS'] = pd.to_numeric(cons_cpb['ANO_CONS'], errors='coerce')
                cons_cpb = cons_cpb.dropna(subset=['ANO_CONS']).drop_duplicates('cpb_norm_raw')
            if cons_tit_col and cons_ano_col:
                cons_title = df_cons[[cons_tit_col, cons_ano_col]].copy()
                cons_title.columns = ['titulo_norm_join', 'ANO_CONS_TIT']
                cons_title['titulo_norm_join'] = cons_title['titulo_norm_join'].map(_norm_title)
                cons_title['ANO_CONS_TIT'] = pd.to_numeric(cons_title['ANO_CONS_TIT'], errors='coerce')
                cons_title = cons_title.dropna(subset=['ANO_CONS_TIT']).drop_duplicates('titulo_norm_join')

        raw_cols = ['DIRETOR', 'TITULO_ORIGINAL', 'cpb_norm_raw', 'titulo_norm_join']
        ata_cpb = (df_ata[df_ata['cpb_norm'].astype(bool)]
                   .merge(df_raw_dir[raw_cols].drop_duplicates(), left_on='cpb_norm', right_on='cpb_norm_raw', how='inner'))
        ata_title = (df_ata[(~df_ata['cpb_norm'].astype(bool)) & df_ata['titulo_norm_join'].astype(bool)]
                     .merge(df_raw_dir[raw_cols].drop_duplicates(), on='titulo_norm_join', how='inner'))
        ata_join = pd.concat([ata_cpb, ata_title], ignore_index=True, sort=False)
        if len(ata_join):
            if len(cons_cpb):
                ata_join = ata_join.merge(cons_cpb, on='cpb_norm_raw', how='left')
            else:
                ata_join['ANO_CONS'] = np.nan
            if len(cons_title):
                ata_join = ata_join.merge(cons_title, on='titulo_norm_join', how='left')
            else:
                ata_join['ANO_CONS_TIT'] = np.nan
            ata_join['ANO_FILME'] = ata_join['ANO_CONS'].combine_first(ata_join['ANO_CONS_TIT'])
            ata_join['ANO_FILME'] = ata_join['ANO_FILME'].combine_first(ata_join['cpb_norm_raw'].map(_year_from_cpb))
            ata_join['FESTIVAL'] = ata_join.get('festival', '').fillna('')
            ata_join.loc[ata_join['FESTIVAL'].astype(str).str.strip().eq(''), 'FESTIVAL'] = 'Base ATA sem festival detalhado'
            ata_join['FILME'] = ata_join.get('titulo', '').fillna('').where(
                ata_join.get('titulo', '').fillna('').astype(str).str.strip().ne(''),
                ata_join.get('TITULO_ORIGINAL', '')
            )
            ata_longas = pd.DataFrame({
                'diretor': ata_join['DIRETOR'],
                'diretor_norm': ata_join['DIRETOR'].map(_proxy_dir_key),
                'FILME': ata_join['FILME'],
                'ANO_FILME': pd.to_numeric(ata_join['ANO_FILME'], errors='coerce'),
                'CPB': ata_join['cpb_norm_raw'],
                'FESTIVAL': ata_join['FESTIVAL'],
                'FONTE_CRUZAMENTO': 'base_festivais_obras_ata.csv + diretores ANCINE',
            })
            longas_frames.append(ata_longas)

    if longas_frames:
        df_longas = pd.concat(longas_frames, ignore_index=True, sort=False)
    else:
        df_longas = pd.DataFrame(columns=['diretor_norm', 'FILME', 'ANO_FILME', 'FESTIVAL', 'FONTE_CRUZAMENTO'])
    df_longas = df_longas[df_longas['diretor_norm'].astype(str).str.len() > 0].copy()
    df_longas['FILME'] = df_longas['FILME'].map(lambda v: _cell_text(v).upper())
    df_longas = df_longas[df_longas['FILME'].astype(str).str.len() > 0].copy()
    df_longas['ANO_FILME'] = pd.to_numeric(df_longas['ANO_FILME'], errors='coerce')
    df_longas = df_longas.drop_duplicates(['diretor_norm', 'FILME', 'ANO_FILME', 'FESTIVAL', 'FONTE_CRUZAMENTO'])

    manual_longas = pd.DataFrame([
        {
            'diretor': 'Leonardo Mouramateus',
            'diretor_norm': _proxy_dir_key('Leonardo Mouramateus'),
            'FILME': 'A VIDA SÃO DOIS DIAS',
            'ANO_FILME': 2022,
            'CPB': '',
            'FESTIVAL': 'FIDMarseille',
            'FONTE_CRUZAMENTO': 'proxy manual usuario | FIDMarseille 2022',
        },
        {
            'diretor': 'Leonardo Mouramateus',
            'diretor_norm': _proxy_dir_key('Leonardo Mouramateus'),
            'FILME': 'GREICE',
            'ANO_FILME': 2024,
            'CPB': '',
            'FESTIVAL': 'IFFR Rotterdam',
            'FONTE_CRUZAMENTO': 'proxy manual usuario | IFFR 2024 Harbour',
        },
    ])
    df_longas = pd.concat([df_longas, manual_longas], ignore_index=True, sort=False)
    df_longas = df_longas.drop_duplicates(['diretor_norm', 'FILME', 'ANO_FILME', 'FESTIVAL', 'FONTE_CRUZAMENTO'])

    dynamic_proxy = {}
    short_names = df_curtas_exp[['diretor_individual', 'diretor_norm']].drop_duplicates()
    long_names = df_longas[['diretor', 'diretor_norm']].drop_duplicates()
    existing_short_keys = set(short_names['diretor_norm'])
    for _, long_row in long_names.iterrows():
        long_key = long_row['diretor_norm']
        if long_key in existing_short_keys:
            continue
        candidates = []
        for _, short_row in short_names.iterrows():
            if _is_artist_name_expansion(short_row['diretor_individual'], long_row['diretor']):
                candidates.append(short_row['diretor_norm'])
        candidates = sorted(set(candidates))
        if len(candidates) == 1:
            dynamic_proxy[long_key] = candidates[0]
    if dynamic_proxy:
        dynamic_mask = df_longas['diretor_norm'].isin(dynamic_proxy)
        df_longas.loc[dynamic_mask, 'FONTE_CRUZAMENTO'] = (
            df_longas.loc[dynamic_mask, 'FONTE_CRUZAMENTO'].astype(str)
            + ' | proxy automático nome artístico/nome oficial'
        )
        df_longas['diretor_norm'] = df_longas['diretor_norm'].map(lambda k: dynamic_proxy.get(k, k))

    proxy_names = ['andre novais', 'thais fujinaga', 'marcelo caetano', 'gabriel martins',
                   'guto parente', 'eva randolph', 'carolina markowicz']
    proxy_rows = []
    existing_proxy_keys = set(df_curtas_exp['diretor_norm'])
    raw_proxy_year = {}
    if os.path.exists(raw_dir_path):
        try:
            df_raw_proxy = pd.read_csv(raw_dir_path, sep=';', dtype=str, on_bad_lines='skip')
            df_raw_proxy['diretor_norm'] = df_raw_proxy.get('DIRETOR', '').map(_proxy_dir_key)
            df_raw_proxy['ano_raw'] = df_raw_proxy.get('CPB', '').map(_year_from_cpb)
            raw_proxy_year = (df_raw_proxy.dropna(subset=['ano_raw'])
                              .groupby('diretor_norm')['ano_raw'].min()
                              .to_dict())
        except Exception:
            raw_proxy_year = {}
    for proxy in proxy_names:
        key = _proxy_dir_key(proxy)
        if key in existing_proxy_keys:
            continue
        proxy_rows.append({
            'titulo': 'proxy nominal informado pelo usuario',
            'diretor': proxy,
            'ano': raw_proxy_year.get(key, np.nan),
            'festival': 'Proxy nominal',
            'secao': '',
            'resultado': 'curta informado pelo usuario',
            'premiado': False,
            'diretor_individual': proxy.title(),
            'diretor_norm': key,
            'fonte_curta': 'proxy_usuario',
            'selecao_contagem': 0,
        })
    if proxy_rows:
        df_curtas_match = pd.concat([df_curtas_exp, pd.DataFrame(proxy_rows)], ignore_index=True, sort=False)
    else:
        df_curtas_match = df_curtas_exp.copy()

    longas_agg = (df_longas.groupby('diretor_norm')
                  .agg(primeiro_longa=('ANO_FILME', 'min'),
                       n_longas_fest=('FILME', 'nunique'),
                       longas_obras=('FILME', lambda x: _join_unique(x, 8)),
                       longas_festivais=('FESTIVAL', lambda x: _join_unique(x, 8)),
                       fontes_longa=('FONTE_CRUZAMENTO', _join_unique))
                  .reset_index())
    by_dir = (df_curtas_match.groupby(['diretor_norm', 'diretor_individual'])
              .agg(primeira_curta=('ano', 'min'),
                   ultima_curta=('ano', 'max'),
                   n_curtas=('titulo', 'nunique'),
                   n_selecoes=('selecao_contagem', 'sum'),
                   n_premios=('premiado', 'sum'),
                   festivais=('festival', _join_unique),
                   fontes_curta=('fonte_curta', _join_unique))
              .reset_index()
              .merge(longas_agg, on='diretor_norm', how='left'))
    by_dir['n_longas_fest'] = by_dir['n_longas_fest'].fillna(0).astype(int)
    by_dir['tem_longa_festival'] = by_dir['n_longas_fest'] > 0
    by_dir['gap_anos'] = by_dir['primeiro_longa'] - by_dir['primeira_curta']
    by_dir['tem_longa_pos_curta'] = by_dir['tem_longa_festival'] & by_dir['gap_anos'].notna() & (by_dir['gap_anos'] >= 0)
    matched_gap = by_dir[by_dir['tem_longa_pos_curta']]

    timeline_people = by_dir[by_dir['tem_longa_pos_curta']][['diretor_norm', 'diretor_individual']].drop_duplicates()
    timeline_short_src = df_curtas_exp[df_curtas_exp['selecao_contagem'] > 0].merge(
        timeline_people, on='diretor_norm', how='inner', suffixes=('', '_match')
    ).copy()
    timeline_short_src['diretor_individual'] = timeline_short_src['diretor_individual_match'].fillna(timeline_short_src['diretor_individual'])
    timeline_short_src['ano_evento'] = pd.to_numeric(timeline_short_src['ano'], errors='coerce')
    timeline_short_src['obra_evento'] = timeline_short_src['titulo'].map(_cell_text)
    timeline_short_src['festival_evento'] = timeline_short_src['festival'].map(_cell_text)
    timeline_short = (timeline_short_src.groupby(['diretor_norm', 'diretor_individual', 'ano_evento', 'obra_evento'], dropna=False)
                      .agg(festival_evento=('festival_evento', _join_unique))
                      .reset_index())
    timeline_short['tipo_evento'] = 'Curta selecionado'
    timeline_short = timeline_short[['diretor_norm', 'diretor_individual', 'tipo_evento', 'ano_evento', 'obra_evento', 'festival_evento']]

    timeline_long = df_longas.merge(
        by_dir[by_dir['tem_longa_pos_curta']][['diretor_norm', 'diretor_individual', 'primeira_curta']],
        on='diretor_norm',
        how='inner'
    ).copy()
    timeline_long['ano_evento'] = pd.to_numeric(timeline_long['ANO_FILME'], errors='coerce')
    timeline_long = timeline_long[timeline_long['ano_evento'].notna()]
    timeline_long = timeline_long[timeline_long['primeira_curta'].isna() | (timeline_long['ano_evento'] >= timeline_long['primeira_curta'])]
    timeline_long['obra_evento'] = timeline_long['FILME'].map(_cell_text)
    timeline_long['festival_evento'] = timeline_long['FESTIVAL'].map(_cell_text)
    timeline_long = (timeline_long.groupby(['diretor_norm', 'diretor_individual', 'ano_evento', 'obra_evento'], dropna=False)
                     .agg(festival_evento=('festival_evento', _join_unique))
                     .reset_index())
    timeline_long['tipo_evento'] = 'Longa selecionado'
    timeline_long = timeline_long[['diretor_norm', 'diretor_individual', 'tipo_evento', 'ano_evento', 'obra_evento', 'festival_evento']]

    timeline = pd.concat([timeline_short, timeline_long], ignore_index=True, sort=False)
    timeline = timeline[timeline['ano_evento'].notna()].drop_duplicates(
        ['diretor_norm', 'tipo_evento', 'ano_evento', 'obra_evento', 'festival_evento']
    )
    timeline_order = (by_dir[by_dir['tem_longa_pos_curta']]
                      .sort_values(['primeira_curta', 'primeiro_longa', 'diretor_individual'])
                      ['diretor_individual'].tolist())
    timeline['diretor_individual'] = pd.Categorical(timeline['diretor_individual'], categories=timeline_order, ordered=True)
    timeline = timeline.sort_values(['diretor_individual', 'ano_evento', 'tipo_evento'])

    n_selecoes = int(len(df_curtas))
    n_filmes = int(df_curtas['titulo'].nunique())
    n_diretores = int(by_dir['diretor_norm'].nunique())
    n_festivais = int(df_curtas['festival'].nunique())
    n_premiados = int(df_curtas['premiado'].sum())
    ano_min = int(df_curtas['ano'].min())
    ano_max = int(df_curtas['ano'].max())
    n_transicao = int(by_dir['tem_longa_pos_curta'].sum())
    taxa_transicao = (n_transicao / n_diretores * 100) if n_diretores else 0
    gap_mediano = float(matched_gap['gap_anos'].median()) if len(matched_gap) else 0
    longas_total = int(df_longas['FILME'].nunique())
    longas_pos = int(timeline_long['obra_evento'].nunique()) if len(timeline_long) else 0
    prob_longas_selecionados = (longas_pos / longas_total * 100) if longas_total else 0
    matriz_a = n_transicao
    matriz_b = max(n_diretores - n_transicao, 0)
    if n_diretores:
        _p = n_transicao / n_diretores
        _z = 1.96
        _den = 1 + (_z ** 2 / n_diretores)
        _center = (_p + (_z ** 2 / (2 * n_diretores))) / _den
        _half = (_z * np.sqrt((_p * (1 - _p) / n_diretores) + (_z ** 2 / (4 * n_diretores ** 2)))) / _den
        trans_ci_low = max(0, (_center - _half) * 100)
        trans_ci_high = min(100, (_center + _half) * 100)
    else:
        trans_ci_low = 0
        trans_ci_high = 0

    base_dirs_total = 0
    base_dirs_fest = 0
    base_filmes_total = 0
    base_filmes_fest = 0
    taxa_base_fest = 0
    ganho_pp = 0
    ganho_mult = 0
    try:
        if os.path.exists(cons_path) and os.path.exists(raw_dir_path):
            df_cons_base = pd.read_excel(cons_path, sheet_name='Obras', dtype=str)
            df_raw_base = pd.read_csv(raw_dir_path, sep=';', dtype=str, on_bad_lines='skip')
            cons_cpb_col = _pick_col(df_cons_base, 'cpb')
            cons_tit_col = _pick_col(df_cons_base, 'projeto')
            cons_fest_col = _pick_col(df_cons_base, 'pontuacao festivais')
            if cons_cpb_col and cons_tit_col and cons_fest_col:
                df_cons_base['cpb_norm_join'] = df_cons_base[cons_cpb_col].map(_norm_cpb)
                df_cons_base['titulo_norm_join'] = df_cons_base[cons_tit_col].map(_norm_title)
                df_cons_base['fest_score_base'] = pd.to_numeric(df_cons_base[cons_fest_col], errors='coerce').fillna(0)
                df_raw_base['cpb_norm_raw'] = df_raw_base.get('CPB', '').map(_norm_cpb)
                df_raw_base['titulo_norm_join'] = df_raw_base.get('TITULO_ORIGINAL', '').map(_norm_title)
                df_raw_base['diretor_norm'] = df_raw_base.get('DIRETOR', '').map(_proxy_dir_key)
                raw_base_cols = ['DIRETOR', 'diretor_norm', 'cpb_norm_raw', 'titulo_norm_join']
                base_cpb = (
                    df_cons_base[df_cons_base['cpb_norm_join'].astype(bool)]
                    .merge(df_raw_base[raw_base_cols].drop_duplicates(),
                           left_on='cpb_norm_join', right_on='cpb_norm_raw', how='left')
                )
                base_title = (
                    df_cons_base[(~df_cons_base['cpb_norm_join'].astype(bool)) & df_cons_base['titulo_norm_join'].astype(bool)]
                    .merge(df_raw_base[raw_base_cols].drop_duplicates(), on='titulo_norm_join', how='left')
                )
                base_join = pd.concat([base_cpb, base_title], ignore_index=True, sort=False)
                base_join = base_join[base_join['diretor_norm'].fillna('').astype(str).str.len() > 0].copy()
                base_dirs_total = int(base_join['diretor_norm'].nunique())
                base_dirs_fest = int(base_join[base_join['fest_score_base'] > 0]['diretor_norm'].nunique())
                base_filmes_total = int(df_cons_base[cons_tit_col].nunique())
                base_filmes_fest = int((df_cons_base['fest_score_base'] > 0).sum())
                taxa_base_fest = (base_dirs_fest / base_dirs_total * 100) if base_dirs_total else 0
                ganho_pp = taxa_transicao - taxa_base_fest
                ganho_mult = (taxa_transicao / taxa_base_fest) if taxa_base_fest else 0
    except Exception:
        base_dirs_total = base_dirs_fest = base_filmes_total = base_filmes_fest = 0
        taxa_base_fest = ganho_pp = ganho_mult = 0

    cl_layout = dict(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10),
        margin=dict(l=54, r=18, t=34, b=46),
    )

    by_year = (df_curtas.groupby('ano')
               .agg(selecoes=('titulo', 'count'), filmes=('titulo', 'nunique'), premios=('premiado', 'sum'))
               .reset_index()
               .sort_values('ano'))
    fig_year = go.Figure()
    fig_year.add_trace(go.Bar(
        x=by_year['ano'].astype(int), y=by_year['selecoes'],
        name='Seleções', marker_color='#00e5c8',
        hovertemplate='Ano %{x}<br>Seleções: %{y}<extra></extra>'
    ))
    fig_year.add_trace(go.Scatter(
        x=by_year['ano'].astype(int), y=by_year['premios'],
        name='Prêmios', mode='lines+markers',
        line=dict(color='#f5c842', width=3), marker=dict(size=8),
        hovertemplate='Ano %{x}<br>Prêmios: %{y}<extra></extra>'
    ))
    fig_year.update_layout(
        title='Evolução histórica: curtas brasileiros em festivais internacionais',
        xaxis=dict(title='Ano', dtick=2, gridcolor='#1e2035'),
        yaxis=dict(title='Seleções / prêmios', gridcolor='#1e2035'),
        height=330,
        barmode='group',
        legend=dict(orientation='h', y=-0.2),
        **cl_layout
    )

    fig_timeline = go.Figure()
    timeline_height = max(360, min(720, 120 + 34 * max(1, n_transicao)))
    for tipo, color, symbol, size in [
        ('Curta selecionado', '#00e5c8', 'circle', 8),
        ('Longa selecionado', '#f5c842', 'diamond', 10),
    ]:
        part = timeline[timeline['tipo_evento'] == tipo]
        if len(part):
            fig_timeline.add_trace(go.Scatter(
                x=part['ano_evento'],
                y=part['diretor_individual'],
                mode='markers',
                name=tipo,
                marker=dict(color=color, symbol=symbol, size=size, line=dict(width=1, color='rgba(0,0,0,.45)')),
                customdata=part[['obra_evento', 'festival_evento']],
                hovertemplate='<b>%{y}</b><br>%{fullData.name}: %{x:.0f}<br>%{customdata[0]}<br>%{customdata[1]}<extra></extra>'
            ))
    fig_timeline.update_layout(
        title='Evolução histórica: diretores com curta e longa selecionados',
        xaxis=dict(title='Ano', dtick=2, gridcolor='#1e2035'),
        yaxis=dict(title='', gridcolor='#1e2035', autorange='reversed'),
        height=timeline_height,
        legend=dict(orientation='h', y=-0.12),
        **{**cl_layout, 'margin': dict(l=185, r=18, t=34, b=54)}
    )

    df_fest = df_fest.sort_values('total', ascending=True)
    fig_fest = go.Figure()
    fig_fest.add_trace(go.Bar(
        y=df_fest['festival'], x=df_fest['total'],
        name='Seleções', orientation='h', marker_color='#5fd1ff',
        hovertemplate='<b>%{y}</b><br>Seleções: %{x}<extra></extra>'
    ))
    fig_fest.add_trace(go.Bar(
        y=df_fest['festival'], x=df_fest['premiados'],
        name='Premiados', orientation='h', marker_color='#f5c842',
        hovertemplate='<b>%{y}</b><br>Premiados: %{x}<extra></extra>'
    ))
    fig_fest.update_layout(
        title='Distribuição por festival',
        xaxis=dict(title='Curtas', gridcolor='#1e2035'),
        yaxis=dict(title='', gridcolor='#1e2035'),
        height=330,
        barmode='group',
        legend=dict(orientation='h', y=-0.18),
        **cl_layout
    )

    fig_trans = go.Figure(go.Bar(
        y=['Longa posterior identificada', 'Sem longa posterior identificada'],
        x=[n_transicao, max(n_diretores - n_transicao, 0)],
        orientation='h',
        marker_color=['#00e5c8', '#5a6080'],
        text=[f'{n_transicao} ({taxa_transicao:.1f}%)',
              f'{max(n_diretores - n_transicao, 0)} ({100 - taxa_transicao:.1f}%)'],
        textposition='outside',
        hovertemplate='%{y}<br>Diretores: %{x}<extra></extra>'
    ))
    fig_trans.update_layout(
        title='Cruzamento nominal: curta em festival -> longa posterior em festival internacional',
        xaxis=dict(title='Diretores', range=[0, max(n_diretores * 1.16, 1)], gridcolor='#1e2035'),
        yaxis=dict(title='', gridcolor='#1e2035'),
        height=320,
        showlegend=False,
        **{**cl_layout, 'margin': dict(l=210, r=58, t=34, b=46)}
    )

    taxa_sem_transicao = max(0, 100 - taxa_transicao)
    fig_prob = go.Figure()
    fig_prob.add_trace(go.Bar(
        y=['Diretores do recorte'],
        x=[taxa_transicao],
        name='Com longa posterior',
        orientation='h',
        marker_color='#00e5c8',
        text=[f'{taxa_transicao:.1f}%<br>{n_transicao}/{n_diretores}'],
        textposition='inside',
        insidetextanchor='middle',
        hovertemplate='Com longa internacional posterior: ' +
                      f'{n_transicao}/{n_diretores} ({taxa_transicao:.1f}%)<br>' +
                      f'IC 95% Wilson: {trans_ci_low:.1f}% a {trans_ci_high:.1f}%<extra></extra>',
    ))
    fig_prob.add_trace(go.Bar(
        y=['Diretores do recorte'],
        x=[taxa_sem_transicao],
        name='Sem longa posterior na base',
        orientation='h',
        marker_color='#5a6080',
        text=[f'{taxa_sem_transicao:.1f}%<br>{matriz_b}/{n_diretores}'],
        textposition='inside',
        insidetextanchor='middle',
        hovertemplate='Sem longa internacional posterior identificada: ' +
                      f'{matriz_b}/{n_diretores} ({taxa_sem_transicao:.1f}%)<extra></extra>',
    ))
    fig_prob.update_layout(
        title='Destino observado após a seleção internacional do curta',
        xaxis=dict(title='% dos diretores', range=[0, 100], ticksuffix='%', gridcolor='#1e2035'),
        yaxis=dict(title='', gridcolor='#1e2035'),
        height=330,
        barmode='stack',
        legend=dict(orientation='h', y=-0.18),
        annotations=[dict(
            x=taxa_transicao,
            y='Diretores do recorte',
            text=f'IC 95%: {trans_ci_low:.1f}% a {trans_ci_high:.1f}%',
            showarrow=True,
            arrowhead=2,
            ax=42,
            ay=-54,
            font=dict(color='#f5c842', size=10),
            arrowcolor='#f5c842',
        )],
        **{**cl_layout, 'margin': dict(l=78, r=28, t=42, b=74)}
    )

    fig_uplift = go.Figure(go.Bar(
        x=['Base geral de diretores de longas', 'Diretores com curta selecionado'],
        y=[taxa_base_fest, taxa_transicao],
        marker_color=['#5a6080', '#00e5c8'],
        text=[
            f'{base_dirs_fest}/{base_dirs_total}<br>{taxa_base_fest:.1f}%',
            f'{n_transicao}/{n_diretores}<br>{taxa_transicao:.1f}%'
        ],
        textposition='outside',
        customdata=[
            'Diretores da base consolidada FSA/ANCINE com ao menos um longa pontuado em festivais internacionais',
            'Diretores do recorte de curtas com longa internacional posterior identificado'
        ],
        hovertemplate='<b>%{x}</b><br>%{text}<br>%{customdata}<extra></extra>',
    ))
    fig_uplift.update_layout(
        title='Quanto aumenta a chance em relação à base geral?',
        xaxis=dict(title='', gridcolor='#1e2035'),
        yaxis=dict(title='% com longa em festival internacional', range=[0, min(100, max(taxa_transicao, taxa_base_fest) * 1.45 + 4)], ticksuffix='%', gridcolor='#1e2035'),
        height=340,
        showlegend=False,
        annotations=[dict(
            x=1,
            y=taxa_transicao,
            text=f'+{ganho_pp:.1f} p.p. | {ganho_mult:.1f}x',
            showarrow=True,
            arrowhead=2,
            ax=-58,
            ay=-58,
            font=dict(color='#f5c842', size=12),
            arrowcolor='#f5c842',
        )] if ganho_mult else [],
        **{**cl_layout, 'margin': dict(l=62, r=28, t=44, b=74)}
    )

    fig_gap = go.Figure()
    if len(matched_gap):
        gap_plot = matched_gap.sort_values('gap_anos')
        fig_gap.add_trace(go.Bar(
            x=gap_plot['diretor_individual'],
            y=gap_plot['gap_anos'],
            marker_color='#a78bfa',
            customdata=gap_plot[['primeira_curta', 'primeiro_longa', 'n_longas_fest', 'longas_obras']],
            hovertemplate='<b>%{x}</b><br>1o curta: %{customdata[0]:.0f}<br>1o longa intl: %{customdata[1]:.0f}<br>Gap: %{y:.0f} anos<br>Longas em festivais: %{customdata[2]}<br>%{customdata[3]}<extra></extra>'
        ))
        fig_gap.add_hline(y=gap_mediano, line_dash='dot', line_color='#f5c842',
                          annotation_text=f'mediana {gap_mediano:.0f} anos',
                          annotation_font_color='#f5c842')
    fig_gap.update_layout(
        title='Gap entre primeira seleção de curta e longa internacional identificada',
        xaxis=dict(title='', tickangle=-25, gridcolor='#1e2035'),
        yaxis=dict(title='Anos', gridcolor='#1e2035'),
        height=320,
        showlegend=False,
        **cl_layout
    )

    top_dir = by_dir.sort_values(['n_premios', 'n_selecoes', 'n_curtas'], ascending=False).head(18).sort_values('n_selecoes')
    fig_dirs = go.Figure(go.Bar(
        y=top_dir['diretor_individual'],
        x=top_dir['n_selecoes'],
        orientation='h',
        marker_color='#00e5c8',
        customdata=top_dir[['n_curtas', 'n_premios', 'festivais']],
        hovertemplate='<b>%{y}</b><br>Seleções: %{x}<br>Curtas: %{customdata[0]}<br>Prêmios: %{customdata[1]}<br>%{customdata[2]}<extra></extra>'
    ))
    fig_dirs.update_layout(
        title='Diretores com mais seleções/premiações de curtas',
        xaxis=dict(title='Seleções de curtas', gridcolor='#1e2035'),
        yaxis=dict(title='', gridcolor='#1e2035'),
        height=420,
        showlegend=False,
        margin=dict(l=185, r=18, t=34, b=44),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10),
    )

    prem_rows = ''
    for _, r in df_prem.sort_values(['ano', 'festival', 'titulo']).iterrows():
        prem_rows += (
            '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
            f'<td style="padding:6px 8px">{html.escape(str(r["titulo"]))}</td>'
            f'<td style="padding:6px 8px;color:var(--muted)">{html.escape(str(r["diretor"]))}</td>'
            f'<td style="text-align:center;padding:6px 8px">{int(r["ano"])}</td>'
            f'<td style="padding:6px 8px;color:var(--accent)">{html.escape(str(r["festival"]))}</td>'
            f'<td style="padding:6px 8px;color:var(--muted);font-size:11px">{html.escape(str(r["resultado"]))}</td>'
            '</tr>'
        )

    trans_rows = ''
    for _, r in by_dir[by_dir['tem_longa_pos_curta']].sort_values('gap_anos').iterrows():
        fonte_curta = _cell_text(r.get('fontes_curta', ''))
        fonte_longa = _cell_text(r.get('fontes_longa', ''))
        longa_txt = _cell_text(r.get('longas_obras', ''))
        festival_longa_txt = _cell_text(r.get('longas_festivais', ''))
        trans_rows += (
            '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
            f'<td style="padding:6px 8px">{html.escape(str(r["diretor_individual"]))}</td>'
            f'<td style="text-align:center;padding:6px 8px">{int(r["primeira_curta"])}</td>'
            f'<td style="text-align:center;padding:6px 8px">{int(r["primeiro_longa"]) if pd.notna(r["primeiro_longa"]) else ""}</td>'
            f'<td style="padding:6px 8px;color:var(--muted);font-size:11px">{html.escape(str(r["festivais"]))}</td>'
            f'<td style="padding:6px 8px;color:var(--text);font-size:11px">{html.escape(longa_txt)}</td>'
            f'<td style="padding:6px 8px;color:var(--muted);font-size:10px">{html.escape(festival_longa_txt)}</td>'
            f'<td style="padding:6px 8px;color:var(--muted);font-size:10px">{html.escape(fonte_curta)}<br>{html.escape(fonte_longa)}</td>'
            '</tr>'
        )

    proxy_audit_rows = ''
    proxy_keys = [_proxy_dir_key(p) for p in proxy_names]
    for _, r in by_dir[by_dir['diretor_norm'].isin(proxy_keys)].sort_values('diretor_norm').iterrows():
        status = 'contado' if bool(r['tem_longa_pos_curta']) else 'sem longa posterior em festival na base'
        if bool(r['tem_longa_festival']) and not bool(r['tem_longa_pos_curta']):
            status = 'longa anterior ou ano insuficiente'
        proxy_audit_rows += (
            '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
            f'<td style="padding:6px 8px">{html.escape(str(r["diretor_individual"]))}</td>'
            f'<td style="padding:6px 8px;color:var(--muted);font-size:11px">{html.escape(_cell_text(r.get("fontes_curta", "")))}</td>'
            f'<td style="text-align:center;padding:6px 8px">{int(r["primeira_curta"]) if pd.notna(r["primeira_curta"]) else ""}</td>'
            f'<td style="text-align:center;padding:6px 8px">{int(r["primeiro_longa"]) if pd.notna(r["primeiro_longa"]) else ""}</td>'
            f'<td style="padding:6px 8px;color:var(--text);font-size:11px">{html.escape(_cell_text(r.get("longas_obras", "")))}</td>'
            f'<td style="padding:6px 8px;color:var(--muted);font-size:10px">{html.escape(_cell_text(r.get("fontes_longa", "")))}</td>'
            f'<td style="padding:6px 8px;color:var(--accent);font-size:11px">{html.escape(status)}</td>'
            '</tr>'
        )

    prob_rows = (
        '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
        '<td style="padding:7px 8px">Taxa observada no recorte</td>'
        f'<td style="padding:7px 8px;text-align:right">{n_transicao}/{n_diretores}</td>'
        f'<td style="padding:7px 8px;text-align:right;color:var(--accent);font-weight:700">{taxa_transicao:.1f}%</td>'
        '<td style="padding:7px 8px;color:var(--muted)">proporção de diretores com curta internacional que têm longa posterior internacional identificada</td>'
        '</tr>'
        '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
        '<td style="padding:7px 8px">Intervalo de confiança</td>'
        f'<td style="padding:7px 8px;text-align:right">Wilson 95%</td>'
        f'<td style="padding:7px 8px;text-align:right;color:var(--text);font-weight:700">{trans_ci_low:.1f}% a {trans_ci_high:.1f}%</td>'
        '<td style="padding:7px 8px;color:var(--muted)">incerteza amostral calculada sobre o denominador observado de diretores, sem linha de base externa</td>'
        '</tr>'
        '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
        '<td style="padding:7px 8px">Sem longa posterior identificada</td>'
        f'<td style="padding:7px 8px;text-align:right">{matriz_b}/{n_diretores}</td>'
        f'<td style="padding:7px 8px;text-align:right;color:var(--text);font-weight:700">{100 - taxa_transicao:.1f}%</td>'
        '<td style="padding:7px 8px;color:var(--muted)">casos não encontrados na base de longas internacionais do projeto</td>'
        '</tr>'
        '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
        '<td style="padding:7px 8px">Mediana do intervalo temporal</td>'
        f'<td style="padding:7px 8px;text-align:right">curta -> longa</td>'
        f'<td style="padding:7px 8px;text-align:right;color:var(--accent);font-weight:700">{gap_mediano:.0f} anos</td>'
        '<td style="padding:7px 8px;color:var(--muted)">calculada apenas entre diretores com transição posterior observada</td>'
        '</tr>'
        '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
        '<td style="padding:7px 8px">Leitura inversa pelos longas</td>'
        f'<td style="padding:7px 8px;text-align:right">{longas_pos}/{longas_total}</td>'
        f'<td style="padding:7px 8px;text-align:right;color:var(--text);font-weight:700">{prob_longas_selecionados:.1f}%</td>'
        '<td style="padding:7px 8px;color:var(--muted)">parcela dos longas selecionados cujo diretor(a) tinha curta anterior no recorte</td>'
        '</tr>'
    )

    curtas_rows = ''
    for _, r in df_curtas.sort_values(['ano', 'festival', 'titulo']).iterrows():
        premiado_txt = 'sim' if bool(r.get('premiado', False)) else ''
        curtas_rows += (
            '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
            f'<td style="padding:6px 8px">{html.escape(str(r["titulo"]))}</td>'
            f'<td style="padding:6px 8px;color:var(--muted)">{html.escape(str(r["diretor"]))}</td>'
            f'<td style="text-align:center;padding:6px 8px">{int(r["ano"]) if pd.notna(r["ano"]) else ""}</td>'
            f'<td style="padding:6px 8px;color:var(--accent)">{html.escape(str(r["festival"]))}</td>'
            f'<td style="padding:6px 8px;color:var(--muted)">{html.escape(_cell_text(r.get("secao", "")))}</td>'
            f'<td style="padding:6px 8px;color:var(--muted);font-size:11px">{html.escape(_cell_text(r.get("resultado", "")))}</td>'
            f'<td style="text-align:center;padding:6px 8px">{premiado_txt}</td>'
            '</tr>'
        )

    section_html = f'''
<div style="overflow-y:auto;height:100%;padding:24px 28px 48px">
  <div style="margin-bottom:18px">
    <div style="font-family:var(--font-head);font-size:21px;font-style:italic;color:var(--text);margin-bottom:5px">Curtas como Porta de Entrada</div>
    <div style="font-size:10px;color:var(--muted);letter-spacing:.04em">{n_filmes} curtas brasileiros · {n_selecoes} seleções internacionais · {n_diretores} diretores creditados · {ano_min}-{ano_max}</div>
  </div>

  <div class="card" style="margin-bottom:16px;border-color:rgba(0,229,200,.35)">
    <div class="card-t" style="color:var(--accent);font-size:17px">Pergunta comparativa: quanto o curta selecionado aumenta a chance?</div>
    <div class="grid2" style="align-items:stretch;margin-bottom:12px">
      <div><div id="cl-curtas-uplift" style="height:340px"></div></div>
      <div style="background:rgba(0,229,200,.07);border:1px solid rgba(0,229,200,.22);border-left:3px solid var(--accent);border-radius:8px;padding:14px 16px;font-size:12px;line-height:1.75;color:var(--muted)">
        <div style="font-family:var(--font-head);font-size:16px;font-style:italic;color:var(--text);margin-bottom:8px">Resposta principal</div>
        <p style="margin:0 0 10px">
          Na base geral, <b style="color:var(--text)">{base_dirs_fest}</b> de
          <b style="color:var(--text)">{base_dirs_total}</b> diretores de longas chegam a festival internacional:
          <b style="color:var(--text)">{taxa_base_fest:.1f}%</b>.
        </p>
        <p style="margin:0 0 10px">
          Entre diretores com curta selecionado internacionalmente, <b style="color:var(--accent)">{n_transicao}</b> de
          <b style="color:var(--accent)">{n_diretores}</b> aparecem depois com longa internacional:
          <b style="color:var(--accent)">{taxa_transicao:.1f}%</b>.
        </p>
        <div style="font-family:var(--font-head);font-size:26px;color:var(--accent);line-height:1.15;margin:8px 0 6px">
          +{ganho_pp:.1f} p.p. · {ganho_mult:.1f}x
        </div>
        <div style="font-size:11px;color:var(--muted);line-height:1.65">
          O curta selecionado funciona como marcador antecipado: ele não garante causalidade, mas identifica um grupo com probabilidade bem maior de chegar ao circuito internacional de longas.
        </div>
      </div>
    </div>
    <div class="kpi-bar">
      <div class="kpi mid"><div class="kpi-l">Base geral</div><div class="kpi-v">{taxa_base_fest:.1f}<span class="kpi-u">%</span></div><div class="kpi-sub">{base_dirs_fest}/{base_dirs_total} diretores de longas</div></div>
      <div class="kpi ok"><div class="kpi-l">Com curta selecionado</div><div class="kpi-v">{taxa_transicao:.1f}<span class="kpi-u">%</span></div><div class="kpi-sub">{n_transicao}/{n_diretores} diretores do recorte</div></div>
      <div class="kpi warn"><div class="kpi-l">Aumento absoluto</div><div class="kpi-v">+{ganho_pp:.1f}<span class="kpi-u"> p.p.</span></div><div class="kpi-sub">diferença entre as taxas</div></div>
      <div class="kpi warn"><div class="kpi-l">Multiplicador</div><div class="kpi-v">{ganho_mult:.1f}<span class="kpi-u">x</span></div><div class="kpi-sub">taxa do recorte / taxa base</div></div>
    </div>
  </div>

  <div style="border-left:3px solid var(--accent);padding:14px 16px;background:var(--surface);border-radius:6px;font-size:12px;line-height:1.7;margin-bottom:20px;color:var(--muted)">
    <b style="color:var(--text)">Base de análise:</b> curtas brasileiros selecionados ou premiados em festivais
    internacionais mapeados em <code>dados/curtas_brasileiros_festivais_internacionais.xlsx</code>. O cruzamento
    com longas usa alias/proxy de nome artístico e combina <code>participacoes_festivais_diretores.csv</code>
    com <code>base_festivais_obras_ata.csv</code> cruzada por CPB/título contra a base ANCINE de diretores.
    Além dos proxies manuais, há regra automática conservadora quando o nome artístico está contido, em ordem,
    no nome oficial da base. Proxies informados manualmente entram sinalizados na auditoria abaixo.
  </div>

  <div class="kpi-bar" style="margin-bottom:20px">
    <div class="kpi ok"><div class="kpi-l">Curtas mapeados</div><div class="kpi-v">{n_filmes}</div><div class="kpi-sub">{n_selecoes} seleções em {n_festivais} festivais</div></div>
    <div class="kpi ok"><div class="kpi-l">Premiações</div><div class="kpi-v">{n_premiados}</div><div class="kpi-sub">{(n_premiados / n_selecoes * 100):.1f}% das seleções registradas</div></div>
    <div class="kpi mid"><div class="kpi-l">Diretores creditados</div><div class="kpi-v">{n_diretores}</div><div class="kpi-sub">créditos individualizados + proxies nominais</div></div>
    <div class="kpi warn"><div class="kpi-l">Transição posterior</div><div class="kpi-v">{n_transicao}<span class="kpi-u"> dirs.</span></div><div class="kpi-sub">{taxa_transicao:.1f}% com longa intl posterior na base do projeto</div></div>
  </div>

  <div class="grid2" style="margin-bottom:16px">
    <div class="card"><div id="cl-curtas-year" style="height:330px"></div><div class="info">Volume anual de seleções e prêmios de curtas brasileiros em festivais internacionais.</div></div>
    <div class="card"><div id="cl-curtas-fest" style="height:330px"></div><div class="info">Ranking dos festivais no Excel corrigido: Cannes, Berlinale, Annecy, Rotterdam, Locarno, Veneza e Clermont-Ferrand.</div></div>
  </div>

  <div class="card" style="margin-bottom:16px">
    <div id="cl-curtas-timeline" style="height:{timeline_height}px"></div>
    <div class="info">Mostra apenas diretores com curta e longa posterior identificados. Curta e longa ficam na mesma linha do(a) diretor(a).</div>
  </div>

  <div class="card" style="margin-bottom:16px">
    <div class="card-t" style="color:var(--accent)">Detalhe do recorte de curtas</div>
    <div class="grid2" style="align-items:stretch;margin-bottom:12px">
      <div><div id="cl-curtas-prob" style="height:330px"></div></div>
      <div style="background:rgba(14,16,24,.55);border:1px solid var(--border);border-radius:8px;padding:14px 16px;font-size:12px;line-height:1.75;color:var(--muted)">
        <div style="font-family:var(--font-head);font-size:15px;font-style:italic;color:var(--text);margin-bottom:8px">Como ler</div>
        <p style="margin:0 0 10px">
          Este bloco mostra apenas o recorte interno dos diretores com curta selecionado. A unidade é o(a) diretor(a),
          não cada seleção de festival, para evitar duplicação de casos.
        </p>
        <p style="margin:0 0 10px">
          Dentro desse grupo, <b style="color:var(--accent)">{n_transicao}/{n_diretores}</b> chegaram depois a longa
          internacional identificado na base do projeto. O intervalo Wilson 95% é
          <b style="color:var(--text)">{trans_ci_low:.1f}%</b> a <b style="color:var(--text)">{trans_ci_high:.1f}%</b>.
        </p>
        <div style="border-top:1px solid var(--border);padding-top:10px;font-size:11px;color:var(--muted)">
          Leitura: associação observada, não prova causal. A comparação principal com a base geral está no bloco de abertura.
        </div>
      </div>
    </div>
  </div>

  <div class="card" style="margin-bottom:16px">
    <div class="card-t" style="color:var(--accent)">Listagem completa dos curtas selecionados</div>
    <div style="overflow-x:auto;max-height:520px"><table style="width:100%;border-collapse:collapse;font-size:11px">
      <thead><tr style="border-bottom:1px solid var(--border)">
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Curta</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Direção</th>
        <th style="padding:6px 8px;text-align:center;color:var(--muted)">Ano</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Festival</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Seção</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Resultado</th>
        <th style="padding:6px 8px;text-align:center;color:var(--muted)">Premiado</th>
      </tr></thead><tbody>{curtas_rows}</tbody>
    </table></div>
  </div>

  <div class="card" style="margin-bottom:16px">
    <div class="card-t" style="color:var(--accent)">Premiados no recorte de curtas</div>
    <div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:11px">
      <thead><tr style="border-bottom:1px solid var(--border)">
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Curta</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Direção</th>
        <th style="padding:6px 8px;text-align:center;color:var(--muted)">Ano</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Festival</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Resultado</th>
      </tr></thead><tbody>{prem_rows}</tbody>
    </table></div>
  </div>

  <div class="card">
    <div class="card-t" style="color:var(--accent)">Diretores com match nominal/proxy curta -> longa internacional</div>
    <div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:11px">
      <thead><tr style="border-bottom:1px solid var(--border)">
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Diretor(a)</th>
        <th style="padding:6px 8px;text-align:center;color:var(--muted)">1º curta</th>
        <th style="padding:6px 8px;text-align:center;color:var(--muted)">1º longa intl</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Festivais dos curtas</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Longas encontrados</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Festivais dos longas</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Fonte do match</th>
      </tr></thead><tbody>{trans_rows}</tbody>
    </table></div>
  </div>

  <div class="card" style="margin-top:16px">
    <div class="card-t" style="color:var(--accent)">Auditoria dos proxies nominais informados</div>
    <div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:11px">
      <thead><tr style="border-bottom:1px solid var(--border)">
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Diretor(a)</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Fonte curta/proxy</th>
        <th style="padding:6px 8px;text-align:center;color:var(--muted)">Ano curta/proxy</th>
        <th style="padding:6px 8px;text-align:center;color:var(--muted)">1º longa</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Longas encontrados</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Fonte longa</th>
        <th style="padding:6px 8px;text-align:left;color:var(--muted)">Status</th>
      </tr></thead><tbody>{proxy_audit_rows}</tbody>
    </table></div>
  </div>
</div>
<script>
(function(){{
  var figs = {{
    'cl-curtas-year': {pio.to_json(fig_year)},
    'cl-curtas-fest': {pio.to_json(fig_fest)},
    'cl-curtas-timeline': {pio.to_json(fig_timeline)},
    'cl-curtas-prob': {pio.to_json(fig_prob)},
    'cl-curtas-uplift': {pio.to_json(fig_uplift)}
  }};
  window.clRender = function(){{
    if(!window.Plotly) return;
    Object.keys(figs).forEach(function(id){{
      var el = document.getElementById(id);
      if(el) Plotly.react(el, figs[id].data, figs[id].layout, {{responsive:true,displaylogo:false}});
    }});
    _resizeVisibleCharts('mega-section-cl');
  }};
  window.clRender();
}})();
</script>'''
    summary = dict(n_filmes=n_filmes, n_selecoes=n_selecoes, n_diretores=n_diretores, n_transicao=n_transicao)
    return section_html, summary

# ─────────────────────────────────────────────────────────────────────────────
# STEP 10 — Build CURTAS→LONGAS section (festival trajectories)
# ─────────────────────────────────────────────────────────────────────────────
_cl_html = ''
try:
    _cl_html, _cl_summary = _build_curtas_longas_html()
    print(
        f"Curtas-Longas: {_cl_summary['n_filmes']} curtas, "
        f"{_cl_summary['n_selecoes']} seleções, "
        f"{_cl_summary['n_diretores']} diretores, "
        f"{_cl_summary['n_transicao']} transições nominais"
    )
except Exception as e:
    print(f"AVISO Curtas-Longas: {e}")
    _cl_html = '<div style="padding:24px 28px"><p style="color:var(--muted)">Dados de curtas não disponíveis.</p></div>'

# ─────────────────────────────────────────────────────────────────────────────
# STEP 11 — Build SOFT POWER section
# ─────────────────────────────────────────────────────────────────────────────
_sp_html = ''
try:
    sp_data = json.load(open(os.path.join(BASE, 'outputs', 'softpower_data.json'), encoding='utf-8'))
    kpi = sp_data['kpi']
    top_obras = sp_data['top_obras'][:20]
    top_cit = sp_data['top_cit'][:15]

    # Scatter: critica vs roi_intl
    scat = sp_data['scat_data']
    scat_x = [d.get('c', 0) for d in scat]
    scat_y = [d.get('ri', 0) for d in scat]
    scat_text = [d.get('t', '') for d in scat]

    fig_sp = go.Figure(go.Scatter(
        x=scat_x, y=scat_y, text=scat_text,
        mode='markers',
        marker=dict(color='#a78bfa', size=7, opacity=0.6, line=dict(width=0.5, color='white')),
        hovertemplate='<b>%{text}</b><br>Crítica: %{x:.2f}<br>ROI Intl: %{y:.1f}<extra></extra>'
    ))
    fig_sp.update_layout(
        title='Índice Crítico × ROI Internacional',
        xaxis_title='Índice Crítico (1–5)', yaxis_title='ROI Internacional (0–100)',
        height=460, margin=dict(l=60, r=20, t=50, b=50),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=11),
        xaxis=dict(gridcolor='#1e2035'), yaxis=dict(gridcolor='#1e2035'),
    )
    fig_sp_json = pio.to_json(fig_sp)

    # Top obras table
    top_obras_rows = ''.join(
        f'<tr><td>{i+1}</td><td>{o["t"]}</td><td>{o.get("a","")}</td>'
        f'<td>{o.get("c",0):.2f}</td><td>{o.get("n",0)}</td></tr>'
        for i, o in enumerate(top_obras)
    )

    _sp_html = f'''
    <div style="padding:24px 28px;overflow-y:auto;height:100%">
      <div style="font-family:var(--font-head);font-size:21px;font-style:italic;margin-bottom:6px">
        Soft Power do Cinema Brasileiro
      </div>
      <p style="color:var(--muted);font-size:11px;margin-bottom:16px;line-height:1.6">
        Crítica cinematográfica e citação acadêmica —
        <b style="color:var(--accent)">{kpi["n_obras"]}</b> obras analisadas,
        cobertura <b style="color:var(--accent)">{kpi["cob_pct"]:.0f}%</b>,
        média geral <b style="color:var(--accent)">{kpi["media"]:.2f}</b>/5.
      </p>
      <div id="sp-fig-scatter" style="background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:12px;margin-bottom:20px"></div>
      <div style="font-family:var(--font-head);font-size:14px;font-style:italic;color:var(--accent);margin-bottom:10px">Top 20 Obras — Índice Crítico</div>
      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse;font-size:11px">
          <thead><tr style="border-bottom:1px solid var(--border)">
            <th style="padding:6px 8px;text-align:left;color:var(--muted)">#</th>
            <th style="padding:6px 8px;text-align:left;color:var(--muted)">Título</th>
            <th style="padding:6px 8px;text-align:left;color:var(--muted)">Ano</th>
            <th style="padding:6px 8px;text-align:right;color:var(--muted)">Índice</th>
            <th style="padding:6px 8px;text-align:right;color:var(--muted)">Fontes</th>
          </tr></thead>
          <tbody>{top_obras_rows}</tbody>
        </table>
      </div>
    </div>
    <script>
    (function(){{
      var d={fig_sp_json};
      Plotly.newPlot('sp-fig-scatter',d.data,d.layout,{{responsive:true,displaylogo:false}});
    }})();
    </script>'''
    print(f"Soft Power: {kpi['n_obras']} obras, {len(top_cit)} diretores citados")
except Exception as e:
    print(f"Soft Power legacy block skipped: {e}")
    _sp_html = '<div style="padding:24px 28px"><p style="color:var(--muted)">Dados de soft power não disponíveis.</p></div>'

# Runtime renderers for sections that were static shells in the source HTML.
# They only aggregate existing project data and keep the source datasets unchanged.
_div_render_script = r'''
<script>
(function(){
  var baseLayout = {
    paper_bgcolor:'rgba(0,0,0,0)',
    plot_bgcolor:'rgba(14,16,24,0.6)',
    font:{color:'#e8eaf2',family:'DM Mono, monospace',size:10},
    margin:{l:48,r:18,t:18,b:44},
    xaxis:{gridcolor:'#1e2035',zerolinecolor:'#2a2d45'},
    yaxis:{gridcolor:'#1e2035',zerolinecolor:'#2a2d45',ticksuffix:'%'},
    legend:{orientation:'h',y:-0.22}
  };
  function mergeLayout(extra){
    return Object.assign({}, baseLayout, extra || {});
  }
  function plot(id, data, layout){
    var el = document.getElementById(id);
    if(!el || !window.Plotly) return;
    Plotly.react(el, data, mergeLayout(layout), {responsive:true,displaylogo:false});
  }
  window.divRender = function(){
    plot('div-chart-raca', [{
      type:'bar',
      x:['Inscritos sem PA','Inscritos com PA','Selecionados com PA'],
      y:[15.2,22.8,32.4],
      marker:{color:['#5a6080','#f5c842','#00e5c8']},
      text:['15,2%','22,8%','32,4%'],
      textposition:'outside',
      hovertemplate:'%{x}<br>%{y:.1f}%<extra></extra>'
    }], {yaxis:{range:[0,38],ticksuffix:'%',gridcolor:'#1e2035'}, showlegend:false});

    plot('div-chart-genero', [{
      type:'bar',
      x:['Inscritas com PA','Selecionadas com PA','Direcao global','Producao executiva'],
      y:[37.0,52.6,30.0,66.0],
      marker:{color:['#a78bfa','#00e5c8','#5fd1ff','#f5c842']},
      text:['37,0%','52,6%','30,0%','66,0%'],
      textposition:'outside',
      hovertemplate:'%{x}<br>%{y:.1f}%<extra></extra>'
    }], {yaxis:{range:[0,74],ticksuffix:'%',gridcolor:'#1e2035'}, showlegend:false});

    plot('div-chart-taxa', [
      {type:'bar',name:'Negros',x:['Sem PA','Com PA'],y:[26.8,14.8],marker:{color:'#f5c842'},hovertemplate:'Negros - %{x}<br>%{y:.1f}%<extra></extra>'},
      {type:'bar',name:'Brancos',x:['Sem PA','Com PA'],y:[29.9,9.1],marker:{color:'#5fd1ff'},hovertemplate:'Brancos - %{x}<br>%{y:.1f}%<extra></extra>'},
      {type:'bar',name:'Mulheres',x:['Com PA'],y:[15.0],marker:{color:'#a78bfa'},hovertemplate:'Mulheres - %{x}<br>%{y:.1f}%<extra></extra>'},
      {type:'bar',name:'Homens',x:['Com PA'],y:[7.9],marker:{color:'#ff7c6e'},hovertemplate:'Homens - %{x}<br>%{y:.1f}%<extra></extra>'}
    ], {barmode:'group', yaxis:{range:[0,34],ticksuffix:'%',gridcolor:'#1e2035'}});

    plot('div-chart-raca-global', [{
      type:'scatter',
      mode:'lines+markers+text',
      x:['Sem PA: inscritos negros','Com PA: inscritos negros','Com PA: selecionados negros'],
      y:[15.2,22.8,32.4],
      line:{color:'#00e5c8',width:3},
      marker:{size:11,color:['#5a6080','#f5c842','#00e5c8']},
      text:['15,2%','22,8%','32,4%'],
      textposition:'top center',
      hovertemplate:'%{x}<br>%{y:.1f}%<extra></extra>'
    }], {yaxis:{range:[0,38],ticksuffix:'%',gridcolor:'#1e2035'}, showlegend:false});

    plot('div-chart-genero-global', [{
      type:'bar',
      x:['Direcao','Producao executiva','Inscritas com PA','Selecionadas com PA'],
      y:[30.0,66.0,37.0,52.6],
      marker:{color:['#5fd1ff','#f5c842','#a78bfa','#00e5c8']},
      text:['30,0%','66,0%','37,0%','52,6%'],
      textposition:'outside',
      hovertemplate:'%{x}<br>%{y:.1f}%<extra></extra>'
    }], {yaxis:{range:[0,74],ticksuffix:'%',gridcolor:'#1e2035'}, showlegend:false});

    _resizeVisibleCharts('mega-section-div');
  };
})();
</script>
'''

_sp_render_script = ''
try:
    df_crit_sp = pd.read_csv(os.path.join(BASE, 'dados', 'critica_obras.csv'))
    df_obras_sp = pd.read_excel(os.path.join(BASE, 'resultados', 'tabela_consolidada_obras.xlsx'))
    df_cit_sp = pd.read_csv(os.path.join(BASE, 'dados', 'citacoes_diretores.csv'))

    df_crit_sp['CRITICA_INDICE_1_5'] = pd.to_numeric(df_crit_sp.get('CRITICA_INDICE_1_5'), errors='coerce')
    df_crit_sp['CRITICA_N_FONTES'] = pd.to_numeric(df_crit_sp.get('CRITICA_N_FONTES'), errors='coerce')
    df_crit_sp['Ano'] = pd.to_numeric(df_crit_sp.get('Ano'), errors='coerce')
    df_crit_ok = df_crit_sp[df_crit_sp['CRITICA_INDICE_1_5'].notna() & (df_crit_sp['CRITICA_N_FONTES'] >= 2)].copy()

    cat_stats = (df_crit_ok.groupby('Categoria', dropna=True)
                 .agg(media=('CRITICA_INDICE_1_5', 'mean'), n=('Projeto', 'count'))
                 .reset_index()
                 .sort_values('media', ascending=True)
                 .tail(14))
    fig_sp_cat = go.Figure(go.Bar(
        x=cat_stats['media'].round(2),
        y=cat_stats['Categoria'],
        customdata=cat_stats['n'],
        orientation='h',
        marker_color='#00e5c8',
        hovertemplate='<b>%{y}</b><br>Indice medio: %{x:.2f}<br>Obras: %{customdata}<extra></extra>'
    ))
    fig_sp_cat.update_layout(
        margin=dict(l=210, r=16, t=10, b=36),
        xaxis=dict(title='Indice medio (1-5)', range=[max(0, float(cat_stats['media'].min()) - .25) if len(cat_stats) else 0, 5],
                   gridcolor='#1e2035'),
        yaxis=dict(title='', gridcolor='#1e2035'),
        height=260,
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10),
    )

    ev_stats = (df_crit_ok[df_crit_ok['Ano'].between(2012, 2023)]
                .groupby('Ano')
                .agg(media=('CRITICA_INDICE_1_5', 'mean'), n=('Projeto', 'count'))
                .reset_index()
                .sort_values('Ano'))
    fig_sp_ev = go.Figure(go.Scatter(
        x=ev_stats['Ano'].astype(int) if len(ev_stats) else [],
        y=ev_stats['media'].round(2) if len(ev_stats) else [],
        customdata=ev_stats['n'] if len(ev_stats) else [],
        mode='lines+markers',
        line=dict(color='#f5c842', width=3),
        marker=dict(size=8, color='#f5c842'),
        hovertemplate='Ano %{x}<br>Indice medio: %{y:.2f}<br>Obras: %{customdata}<extra></extra>',
    ))
    fig_sp_ev.update_layout(
        margin=dict(l=44, r=16, t=10, b=36),
        xaxis=dict(title='Ano', dtick=1, gridcolor='#1e2035'),
        yaxis=dict(title='Indice medio', range=[1, 5], gridcolor='#1e2035'),
        height=260,
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10),
    )

    for col in ['CRITICA_INDICE_1_5', 'CRITICA_N_FONTES', 'ROI Internacional (0-100)']:
        df_obras_sp[col] = pd.to_numeric(df_obras_sp.get(col), errors='coerce')
    scat_sp = df_obras_sp[
        df_obras_sp['CRITICA_INDICE_1_5'].notna() &
        (df_obras_sp['CRITICA_N_FONTES'] >= 2) &
        df_obras_sp['ROI Internacional (0-100)'].notna()
    ].copy()
    fig_sp_scat = go.Figure(go.Scatter(
        x=scat_sp['CRITICA_INDICE_1_5'],
        y=scat_sp['ROI Internacional (0-100)'],
        text=scat_sp['Projeto'],
        customdata=scat_sp[['Ano', 'Categoria']],
        mode='markers',
        marker=dict(color='#a78bfa', size=7, opacity=.62, line=dict(width=.4, color='white')),
        hovertemplate='<b>%{text}</b><br>Ano: %{customdata[0]}<br>%{customdata[1]}<br>Critica: %{x:.2f}<br>ROI intl: %{y:.1f}<extra></extra>',
    ))
    fig_sp_scat.update_layout(
        margin=dict(l=54, r=18, t=10, b=44),
        xaxis=dict(title='Indice critico (1-5)', range=[1, 5], gridcolor='#1e2035'),
        yaxis=dict(title='ROI Internacional (0-100)', gridcolor='#1e2035'),
        height=320,
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10),
    )

    df_cit_sp['CITA_SOMA_CIT'] = pd.to_numeric(df_cit_sp.get('CITA_SOMA_CIT'), errors='coerce').fillna(0)
    df_cit_nonzero = df_cit_sp[df_cit_sp['CITA_SOMA_CIT'] > 0].copy()
    topcit_sp = df_cit_nonzero.nlargest(25, 'CITA_SOMA_CIT').sort_values('CITA_SOMA_CIT', ascending=True)
    fig_sp_topcit = go.Figure(go.Bar(
        x=topcit_sp['CITA_SOMA_CIT'],
        y=topcit_sp['DIRETOR'],
        orientation='h',
        marker_color='#00e5c8',
        hovertemplate='<b>%{y}</b><br>Citacoes: %{x:.0f}<extra></extra>'
    ))
    fig_sp_topcit.update_layout(
        margin=dict(l=190, r=16, t=10, b=38),
        xaxis=dict(title='Citacoes OpenAlex', gridcolor='#1e2035'),
        yaxis=dict(title='', gridcolor='#1e2035'),
        height=340,
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10),
    )

    fig_sp_dist = go.Figure(go.Histogram(
        x=df_cit_nonzero['CITA_SOMA_CIT'],
        nbinsx=30,
        marker_color='#f5c842',
        hovertemplate='Faixa de citacoes: %{x}<br>Diretores: %{y}<extra></extra>'
    ))
    fig_sp_dist.update_layout(
        margin=dict(l=48, r=16, t=10, b=42),
        xaxis=dict(title='Citacoes totais por diretor', gridcolor='#1e2035'),
        yaxis=dict(title='Diretores', gridcolor='#1e2035'),
        height=340,
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,0.6)',
        font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10),
    )

    _sp_render_script = f'''
<script>
(function(){{
  var figCat={pio.to_json(fig_sp_cat)};
  var figEv={pio.to_json(fig_sp_ev)};
  var figScat={pio.to_json(fig_sp_scat)};
  var figTopCit={pio.to_json(fig_sp_topcit)};
  var figDist={pio.to_json(fig_sp_dist)};
  function plotFig(id, fig){{
    var el = document.getElementById(id);
    if(!el || !window.Plotly) return;
    Plotly.react(el, fig.data, fig.layout, {{responsive:true,displaylogo:false}});
  }}
  window.spRender = function(which){{
    if(which === 'citacao') {{
      plotFig('sp-chart-topcit', figTopCit);
      plotFig('sp-chart-dist', figDist);
    }} else {{
      plotFig('sp-chart-cat', figCat);
      plotFig('sp-chart-ev', figEv);
      plotFig('sp-chart-scat', figScat);
    }}
    _resizeVisibleCharts('mega-section-sp');
  }};
  window.spShow = function(which){{
    var critica = which !== 'citacao';
    var pc = document.getElementById('sp-panel-critica');
    var pa = document.getElementById('sp-panel-citacao');
    var tc = document.getElementById('sp-tab-critica');
    var ta = document.getElementById('sp-tab-citacao');
    if(pc) pc.style.display = critica ? 'block' : 'none';
    if(pa) pa.style.display = critica ? 'none' : 'block';
    if(tc) {{
      tc.style.color = critica ? 'var(--accent)' : 'var(--muted)';
      tc.style.borderBottomColor = critica ? 'var(--accent)' : 'transparent';
    }}
    if(ta) {{
      ta.style.color = critica ? 'var(--muted)' : 'var(--accent)';
      ta.style.borderBottomColor = critica ? 'transparent' : 'var(--accent)';
    }}
    window.spRender(critica ? 'critica' : 'citacao');
  }};
}})();
</script>
'''
    print("Soft Power runtime charts: renderers generated from critica/citacoes/consolidado")
except Exception as e:
    print(f"AVISO Soft Power renderers: {e}")
    _sp_render_script = '''
<script>
(function(){
  window.spShow = window.spShow || function(which){
    var critica = which !== 'citacao';
    var pc = document.getElementById('sp-panel-critica');
    var pa = document.getElementById('sp-panel-citacao');
    if(pc) pc.style.display = critica ? 'block' : 'none';
    if(pa) pa.style.display = critica ? 'none' : 'block';
  };
})();
</script>
'''

def _build_soft_power_panel_v2():
    def _norm_name(value):
        value = '' if pd.isna(value) else str(value)
        value = unicodedata.normalize('NFD', value.upper())
        value = ''.join(ch for ch in value if unicodedata.category(ch) != 'Mn')
        value = re.sub(r'[^A-Z0-9 ]+', ' ', value)
        return re.sub(r'\s+', ' ', value).strip()

    def _title_case_name(value):
        parts = []
        small = {'DE', 'DA', 'DAS', 'DO', 'DOS', 'E'}
        for p in str(value).split():
            parts.append(p.lower() if p.upper() in small else p[:1].upper() + p[1:].lower())
        return ' '.join(parts)

    artistic_alias = {
        'BRUNO GULARTE BARRETO': 'Bruno Barreto',
        'BRUNO VILLELA BARRETO BORGES': 'Bruno Barreto',
        'FABIO VILLELA BARRETO BORGES': 'Fabio Barreto',
        'EDUARDO DE OLIVEIRA COUTINHO': 'Eduardo Coutinho',
        'CARLA DE ANDRADE CAMURATI': 'Carla Camurati',
        'CARLOS OSCAR REICHENBACH FILHO': 'Carlos Reichenbach',
        'CARLOS OSCAR REICHENBACH': 'Carlos Reichenbach',
        'CACÃƒÂ DIEGUES': 'Caca Diegues',
        'CACA DIEGUES': 'Caca Diegues',
        'FERNANDO FERREIRA MEIRELLES': 'Fernando Meirelles',
        'JOSE BASTOS PADILHA NETO': 'Jose Padilha',
        'CARLOS IMPERIO HAMBURGER': 'Cao Hamburger',
        'BRENO DA SILVEIRA SOUZA': 'Breno Silveira',
        'KLEBER DE MENDONCA VASCONCELLOS FILHO': 'Kleber Mendonca Filho',
        'GABRIEL MASCARO SEABRA DE MELO': 'Gabriel Mascaro',
        'LUIZ FERNANDO CARVALHO DE ALMEIDA': 'Luiz Fernando Carvalho',
        'NELSON PEREIRA DOS SANTOS': 'Nelson Pereira dos Santos',
        'ARNALDO JABOR': 'Arnaldo Jabor',
        'ANDREA TONACCI': 'Andrea Tonacci',
        'ANSELMO DUARTE': 'Anselmo Duarte',
        'WALTER SALLES': 'Walter Salles',
        'KARIM AINOUZ': 'Karim Ainouz',
        'MARCELO GOMES': 'Marcelo Gomes',
    }

    display_map = {}
    for path in [
        os.path.join(BASE, 'dados', 'perfil_festivais_diretores.csv'),
        os.path.join(BASE, 'dados', 'prestigio_diretores.csv'),
    ]:
        if os.path.exists(path):
            try:
                dfx = pd.read_csv(path, sep=None, engine='python', encoding='utf-8-sig').fillna('')
                for _, row in dfx.iterrows():
                    raw = str(row.get('DIRETOR', '')).strip()
                    key = _norm_name(row.get('DIRETOR_NORM', raw))
                    if raw and key and key not in display_map:
                        display_map[key] = raw
            except Exception:
                pass

    def _display_director(raw):
        key = _norm_name(raw)
        if key in artistic_alias:
            return artistic_alias[key]
        if key in display_map:
            return _title_case_name(display_map[key])
        return _title_case_name(raw)

    def _num_col(df, col):
        if col in df.columns:
            return pd.to_numeric(df[col], errors='coerce').fillna(0)
        return pd.Series([0] * len(df), index=df.index, dtype=float)

    def _norm_plain(value):
        value = '' if pd.isna(value) else str(value)
        value = unicodedata.normalize('NFD', value.upper())
        value = ''.join(ch for ch in value if unicodedata.category(ch) != 'Mn')
        value = re.sub(r'[^A-Z0-9 ]+', ' ', value)
        return re.sub(r'\s+', ' ', value).strip()

    def _load_cinema_metadata():
        rows = []
        raw_dir = os.path.join(BASE, 'raw', 'obras-nao-pub-brasileiras-csv')
        if not os.path.isdir(raw_dir):
            return pd.DataFrame(columns=['CPB_key', 'is_cinema'])
        usecols = [
            'CPB', 'TIPO_OBRA', 'ORGANIZACAO_TEMPORAL', 'DURACAO_TOTAL_MINUTOS',
            'QUANTIDADE_EPISODIOS', 'SEGMENTO_DESTINACAO_INICIAL'
        ]
        for name in os.listdir(raw_dir):
            if not name.lower().endswith('.csv'):
                continue
            path = os.path.join(raw_dir, name)
            try:
                part = pd.read_csv(path, sep=';', encoding='utf-8-sig', dtype=str, usecols=lambda c: c in usecols).fillna('')
            except Exception:
                continue
            rows.append(part)
        if not rows:
            return pd.DataFrame(columns=['CPB_key', 'is_cinema'])
        meta = pd.concat(rows, ignore_index=True).drop_duplicates('CPB')
        meta['CPB_key'] = meta['CPB'].astype(str).str.strip()
        meta['tipo_norm'] = meta.get('TIPO_OBRA', '').map(_norm_plain)
        meta['org_norm'] = meta.get('ORGANIZACAO_TEMPORAL', '').map(_norm_plain)
        meta['seg_norm'] = meta.get('SEGMENTO_DESTINACAO_INICIAL', '').map(_norm_plain)
        meta['duracao_min'] = pd.to_numeric(
            meta.get('DURACAO_TOTAL_MINUTOS', '').astype(str).str.replace(',', '.', regex=False),
            errors='coerce'
        ).fillna(0)
        meta['episodios'] = pd.to_numeric(meta.get('QUANTIDADE_EPISODIOS', ''), errors='coerce').fillna(0)
        long_form = meta['duracao_min'] >= 60
        type_ok = meta['tipo_norm'].isin(['FICCAO', 'DOCUMENTARIO', 'ANIMACAO'])
        non_serial = meta['org_norm'].str.contains('NAO SERIADA', na=False) & (meta['episodios'] <= 1)
        theatrical = meta['seg_norm'].str.contains('SALAS DE EXIBICAO', na=False)
        not_tv = ~meta['seg_norm'].str.contains('TV ABERTA|TV PAGA|RADIODIFUSAO|COMUNICACAO ELETRONICA|VIDEO DOMESTICO|VIDEO POR DEMANDA', na=False)
        meta['is_cinema'] = theatrical | (type_ok & non_serial & long_form & not_tv)
        return meta[['CPB_key', 'is_cinema', 'TIPO_OBRA', 'DURACAO_TOTAL_MINUTOS', 'SEGMENTO_DESTINACAO_INICIAL']]

    master = pd.read_excel(os.path.join(BASE, 'resultados', 'tabela_consolidada_obras.xlsx'), sheet_name='Obras')
    master['Projeto'] = master.get('Projeto', '').fillna('').astype(str)
    master['Categoria'] = master.get('Categoria', '').fillna('').astype(str)
    master['Ano'] = pd.to_numeric(master.get('Ano'), errors='coerce')
    master['CPB_key'] = master.get('CPB', '').fillna('').astype(str)
    master['titulo_key'] = master['Projeto'].map(lambda x: re.sub(r'\s+', ' ', unicodedata.normalize('NFD', str(x).upper())).strip())
    master = master.sort_values(['CPB_key', 'titulo_key', 'Ano'])
    obra = master.drop_duplicates(subset=['CPB_key', 'titulo_key'], keep='first').copy()
    # Recorte amplo: mesmo universo historico usado no painel de produtoras,
    # restrito a filmes de cinema por metadados ANCINE do CPB.
    cinema_meta = _load_cinema_metadata()
    obra = obra.merge(cinema_meta, on='CPB_key', how='left')
    category_not_tv = ~obra['Categoria'].str.upper().str.contains('_TV_EXCLUIR| TV|TV-|TV_|SERIE|S[ÉE]RIE', na=False)
    obra = obra[category_not_tv & (obra['is_cinema'].fillna(False))].copy()

    crit_raw_path = os.path.join(BASE, 'dados', 'critica_obras.csv')
    if os.path.exists(crit_raw_path):
        crit_raw_cols = [
            'CPB', 'ADORO_SCORE_CRITICO', 'ADORO_SCORE_USUARIO', 'ADORO_N_CRITICOS',
            'ADORO_N_USUARIOS', 'ADORO_URL', 'ADORO_MATCH_CONF', 'MC_METASCORE',
            'MC_USER_SCORE', 'MC_N_REVIEWS', 'MC_MATCH_CONF', 'IMDB_RATING',
            'IMDB_N_VOTES', 'LB_RATING', 'RT_TOMATOMETER',
        ]
        crit_raw = pd.read_csv(crit_raw_path, sep=None, engine='python', encoding='utf-8-sig', dtype=str).fillna('')
        crit_raw = crit_raw[[c for c in crit_raw_cols if c in crit_raw.columns]].copy()
        crit_raw['CPB_key'] = crit_raw.get('CPB', '').astype(str).str.strip()
        crit_raw = crit_raw.drop(columns=['CPB'], errors='ignore').drop_duplicates('CPB_key')
        obra = obra.merge(crit_raw, on='CPB_key', how='left')

    obra['critica'] = _num_col(obra, 'CRITICA_INDICE_1_5')
    obra['critica_n'] = _num_col(obra, 'CRITICA_N_FONTES')
    obra['adoro_critica'] = _num_col(obra, 'ADORO_SCORE_CRITICO')
    obra['adoro_usuario'] = _num_col(obra, 'ADORO_SCORE_USUARIO')
    obra['adoro_n_criticas'] = _num_col(obra, 'ADORO_N_CRITICOS')
    obra['adoro_n_usuarios'] = _num_col(obra, 'ADORO_N_USUARIOS')
    obra['mc_meta'] = _num_col(obra, 'MC_METASCORE')
    obra['mc_user'] = _num_col(obra, 'MC_USER_SCORE')
    obra['mc_reviews'] = _num_col(obra, 'MC_N_REVIEWS')
    obra['imdb_rating'] = _num_col(obra, 'IMDB_RATING')
    obra['imdb_votes'] = _num_col(obra, 'IMDB_N_VOTES')
    obra['lb_rating'] = _num_col(obra, 'LB_RATING')
    obra['rt_tomato'] = _num_col(obra, 'RT_TOMATOMETER')
    obra['critica_reviews'] = np.where(obra['adoro_critica'] > 0, obra['adoro_n_criticas'], 0) + np.where(obra['mc_meta'] > 0, obra['mc_reviews'], 0)
    professional_sum = (
        np.where(obra['adoro_critica'] > 0, obra['adoro_critica'], 0) +
        np.where(obra['mc_meta'] > 0, obra['mc_meta'] / 20, 0) +
        np.where(obra['rt_tomato'] > 0, obra['rt_tomato'] / 20, 0)
    )
    professional_n = (
        (obra['adoro_critica'] > 0).astype(int) +
        (obra['mc_meta'] > 0).astype(int) +
        (obra['rt_tomato'] > 0).astype(int)
    )
    obra['critica_profissional'] = np.where(professional_n > 0, professional_sum / professional_n, 0)
    obra['critica_confianca'] = obra.get('CRITICA_CONFIANCA', '').fillna('').astype(str).str.lower().str.strip()
    obra['roi_intl'] = _num_col(obra, 'ROI Internacional (0-100)')
    obra['fest_score'] = _num_col(obra, 'Pontuação Festivais')
    obra['adm_eu'] = _num_col(obra, 'Adm. EU — Lumière')
    obra['vod_paises'] = _num_col(obra, 'VOD Intl — N Países')
    obra['alcance_paises'] = _num_col(obra, 'Total Países Alcançados')
    obra['tem_festival'] = obra['fest_score'] > 0
    obra['tem_alcance'] = obra['tem_festival'] | (obra['adm_eu'] > 0) | (obra['vod_paises'] > 0)
    obra['critica'] = obra['critica_profissional']
    crit_all = obra[obra['critica'] > 0].copy()
    crit_ok = crit_all[
        (crit_all['critica_confianca'] == 'alta') &
        (crit_all['critica_reviews'] >= 5)
    ].copy()
    crit_low = crit_all.drop(crit_ok.index).copy()

    cat_stats = (
        obra.groupby('Categoria', dropna=False)
        .agg(
            n_obras=('Projeto', 'count'),
            n_festival=('tem_festival', 'sum'),
            n_alcance=('tem_alcance', 'sum'),
            fest_media=('fest_score', 'mean'),
            roi_intl=('roi_intl', 'mean'),
            critica_media=('critica', lambda s: s[obra.loc[s.index, 'critica_n'] >= 2].replace(0, np.nan).mean()),
        )
        .reset_index()
    )
    cat_stats = cat_stats[cat_stats['n_obras'] >= 5].copy()
    cat_stats['pct_festival'] = cat_stats['n_festival'] / cat_stats['n_obras'] * 100
    cat_stats = cat_stats.sort_values(['pct_festival', 'fest_media'], ascending=True).tail(14)

    top_crit = crit_ok.sort_values(['critica', 'critica_reviews', 'roi_intl', 'Projeto'], ascending=[False, False, False, True])
    low_crit = crit_low.sort_values(['critica', 'critica_reviews', 'roi_intl', 'Projeto'], ascending=[False, False, False, True])

    allowed_terms = [
        'cinema', 'film', 'screen', 'audiovisual', 'imagem', 'contracampo',
        'rebeca', 'significacao', 'comunicacao', 'comunica', 'galaxia',
        'famecos', 'devires', 'fotocinema', 'aniki', 'alphaville',
        'avanca', 'doc online', 'libero', 'alceu', 'social text',
        'cultura audiovisual', 'teatro', 'theatre', 'media',
    ]
    blocked_terms = [
        'microbiology', 'materials', 'chem', 'medical entomology', 'lepidopter',
        'cancer epidemiology', 'functional materials', 'advanced materials',
        'physical review', 'frontiers in microbiology',
    ]

    cit = pd.read_csv(os.path.join(BASE, 'dados', 'citacoes_diretores.csv'), sep=None, engine='python', encoding='utf-8-sig').fillna('')
    for col in ['CITA_N_PAPERS', 'CITA_SOMA_CIT', 'CITA_MAX_CIT']:
        cit[col] = pd.to_numeric(cit.get(col), errors='coerce').fillna(0)
    cit['diretor_norm'] = cit['DIRETOR'].map(_norm_name)
    cit['diretor_artistico'] = cit['DIRETOR'].map(_display_director)
    cit['venues_lower'] = cit['CITA_VENUES'].astype(str).str.lower()
    cit['venue_area'] = cit['venues_lower'].map(lambda v: any(t in v for t in allowed_terms))
    cit['venue_bloqueado'] = cit['venues_lower'].map(lambda v: any(t in v for t in blocked_terms))
    cit['identidade_curada'] = cit['diretor_norm'].map(lambda k: k in display_map or k in artistic_alias)
    cit['_assinatura_biblio'] = (
        cit['CITA_SOMA_CIT'].astype(int).astype(str) + '|' +
        cit['CITA_N_PAPERS'].astype(int).astype(str) + '|' +
        cit['venues_lower']
    )
    sig_counts = cit['_assinatura_biblio'].value_counts()
    cit['assinatura_repetida'] = cit['_assinatura_biblio'].map(sig_counts).fillna(0).astype(int) > 2
    cit['aprovado'] = (cit['CITA_SOMA_CIT'] > 0) & cit['venue_area'] & (~cit['assinatura_repetida'] | cit['identidade_curada'])
    cit['alerta_misto'] = cit['venue_area'] & cit['venue_bloqueado']
    cit_ok = cit[cit['aprovado']].sort_values('CITA_SOMA_CIT', ascending=False).drop_duplicates('diretor_artistico', keep='first').copy()
    cit_rej = cit[(cit['CITA_SOMA_CIT'] > 0) & ~cit['aprovado']].sort_values('CITA_SOMA_CIT', ascending=False).head(18).copy()
    cit_top = cit_ok.head(25).sort_values('CITA_SOMA_CIT', ascending=True)
    cit_ok['h_proxy'] = np.minimum(cit_ok['CITA_N_PAPERS'], np.sqrt(cit_ok['CITA_SOMA_CIT']).astype(int))
    cit_top_table = cit_ok.sort_values(['CITA_SOMA_CIT', 'CITA_N_PAPERS'], ascending=False).head(25)

    kpi_filmes = len(obra)
    kpi_crit = len(crit_ok)
    kpi_crit_low = len(crit_low)
    kpi_fest = int(obra['tem_festival'].sum())
    kpi_cit = len(cit_ok)
    kpi_rej = int(((cit['CITA_SOMA_CIT'] > 0) & ~cit['aprovado']).sum())
    crit_media = float(crit_ok['critica'].mean()) if len(crit_ok) else 0
    pct_crit = kpi_crit / kpi_filmes * 100 if kpi_filmes else 0

    def _esc(v):
        return html.escape('' if pd.isna(v) else str(v))

    def _fmt(v, dec=0):
        try:
            return f'{float(v):,.{dec}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
        except Exception:
            return ''

    def _source_detail_rows(r):
        rows = []

        def _to_float(v):
            try:
                return float(v)
            except Exception:
                return 0.0

        def add(source, kind, score, count='', obs=''):
            score_v = _to_float(score)
            if score_v <= 0:
                return
            count_txt = _fmt(count) if str(count).strip() not in ('', 'nan') else ''
            rows.append(
                '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
                f'<td style="padding:5px 8px;color:var(--text)">{_esc(source)}</td>'
                f'<td style="padding:5px 8px;color:var(--muted)">{_esc(kind)}</td>'
                f'<td style="padding:5px 8px;text-align:center;color:var(--gold)">{_fmt(score_v, 2)}</td>'
                f'<td style="padding:5px 8px;text-align:center;color:var(--muted)">{count_txt}</td>'
                f'<td style="padding:5px 8px;color:var(--muted);font-size:9px">{_esc(obs)}</td>'
                '</tr>'
            )

        add('AdoroCinema', 'critica profissional', r.get('adoro_critica', 0), r.get('adoro_n_criticas', 0), 'nota original 1-5; N = criticas no raw; usada no indice')
        add('Metacritic', 'metascore', _to_float(r.get('mc_meta', 0)) / 20, r.get('mc_reviews', 0), 'metascore convertido de 0-100 para 1-5; usado no indice')
        add('Rotten Tomatoes', 'tomatometer', _to_float(r.get('rt_tomato', 0)) / 20, '', 'tomatometer convertido de 0-100 para 1-5; usado no indice')
        if not rows:
            return '<tr><td colspan="5" style="padding:6px 8px;color:var(--muted)">Sem detalhe bruto disponivel.</td></tr>'
        return ''.join(rows)

    def _rows_crit_table(data, table_key):
        rows = []
        for i, (_, r) in enumerate(data.iterrows(), 1):
            try:
                ano_txt = str(int(float(r["Ano"])))
            except Exception:
                ano_txt = _esc(r["Ano"])
            conf_txt = str(r.get('critica_confianca', '') or '').strip()
            if not conf_txt:
                conf_txt = 'sem class.'
            detail_id = f'sp-crit-detail-{table_key}-{i}'
            raw_sources = [s.strip() for s in str(r.get('CRITICA_FONTES', '') or '').split('|') if s.strip()]
            professional_sources = [s for s in raw_sources if s in ('adoro_critico', 'mc_meta', 'rt')]
            raw_fontes = ' | '.join(professional_sources)
            if not raw_fontes and float(r.get('adoro_critica', 0) or 0) > 0:
                raw_fontes = 'adoro_critico'
            detail_rows = _source_detail_rows(r)
            rows.append(
                f'<tr onclick="spToggleCritDetail(\'{detail_id}\')" style="border-bottom:1px solid rgba(255,255,255,.05);cursor:pointer">'
                f'<td style="padding:6px 8px;color:var(--muted)">{i}</td>'
                f'<td style="padding:6px 8px;color:var(--text)">{_esc(r["Projeto"])}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--muted)">{ano_txt}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--gold)">{_fmt(r["critica"], 2)}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--muted)">{_fmt(r["critica_reviews"])}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--muted)">{_esc(conf_txt)}</td>'
                f'<td style="padding:6px 8px;color:var(--muted);font-size:9px">{_esc(r["Categoria"])}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--accent)">{_fmt(r["fest_score"], 1)}</td>'
                '</tr>'
                f'<tr id="{detail_id}" style="display:none;background:rgba(255,255,255,.025)"><td></td><td colspan="7" style="padding:10px 8px 12px">'
                f'<div style="font-size:9px;color:var(--muted);line-height:1.55;margin-bottom:8px">Indice recalculado somente com critica profissional. Fontes profissionais no raw: {_esc(raw_fontes)}. URL AdoroCinema: {_esc(r.get("ADORO_URL", ""))}. Match Adoro: {_esc(r.get("ADORO_MATCH_CONF", ""))}.</div>'
                '<table style="width:100%;border-collapse:collapse;font-size:9.5px"><thead><tr style="border-bottom:1px solid var(--border)">'
                '<th style="padding:5px 8px;text-align:left;color:var(--muted)">Veiculo/fonte</th>'
                '<th style="padding:5px 8px;text-align:left;color:var(--muted)">Tipo</th>'
                '<th style="padding:5px 8px;text-align:center;color:var(--muted)">Nota 1-5</th>'
                '<th style="padding:5px 8px;text-align:center;color:var(--muted)">N</th>'
                '<th style="padding:5px 8px;text-align:left;color:var(--muted)">Observacao</th>'
                f'</tr></thead><tbody>{detail_rows}</tbody></table></td></tr>'
            )
        return ''.join(rows)

    def _crit_table_html(data, title, subtitle, table_key, max_height='720px'):
        return f'''
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:16px;margin-bottom:14px">
      <div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">{title}</div>
      <div style="font-size:10px;color:var(--muted);line-height:1.55;margin-bottom:12px">{subtitle}</div>
      <div style="overflow-x:auto;max-height:{max_height}">
        <table style="width:100%;border-collapse:collapse;font-size:10px">
          <thead><tr style="border-bottom:1px solid var(--border)">
            <th style="padding:6px 8px;text-align:left;color:var(--muted)">#</th>
            <th style="padding:6px 8px;text-align:left;color:var(--muted)">Titulo</th>
            <th style="padding:6px 8px;text-align:center;color:var(--muted)">Ano</th>
            <th style="padding:6px 8px;text-align:center;color:var(--muted)">Indice</th>
            <th style="padding:6px 8px;text-align:center;color:var(--muted)">N criticas</th>
            <th style="padding:6px 8px;text-align:center;color:var(--muted)">Conf.</th>
            <th style="padding:6px 8px;text-align:left;color:var(--muted)">Categoria</th>
            <th style="padding:6px 8px;text-align:center;color:var(--muted)">Fest.</th>
          </tr></thead>
          <tbody>{_rows_crit_table(data, table_key)}</tbody>
        </table>
      </div>
    </div>'''

    def _rows_top_cit():
        rows = []
        for i, (_, r) in enumerate(cit_top_table.iterrows(), 1):
            venues = str(r.get('CITA_VENUES', '')).split(' | ')
            venues = ' | '.join(venues[:3])
            if bool(r.get('alerta_misto')):
                venues = 'MISTO: ' + venues
            rows.append(
                '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
                f'<td style="padding:6px 8px;color:var(--muted)">{i}</td>'
                f'<td style="padding:6px 8px;color:var(--text)">{_esc(r["diretor_artistico"])}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--accent)">{_fmt(r["CITA_SOMA_CIT"])}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--muted)">{_fmt(r["CITA_N_PAPERS"])}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--purple)">{_fmt(r["h_proxy"])}</td>'
                f'<td style="padding:6px 8px;color:var(--muted);font-size:9px">{_esc(venues)}</td>'
                '</tr>'
            )
        return ''.join(rows)

    def _rows_audit():
        rows = []
        for _, r in cit_rej.iterrows():
            reason = 'sem venue da area'
            if bool(r.get('venue_bloqueado')):
                reason = 'termo tecnico externo sem evidencia audiovisual'
            if bool(r.get('assinatura_repetida')) and not bool(r.get('identidade_curada')):
                reason = 'assinatura bibliografica repetida em homonimos'
            rows.append(
                '<tr style="border-bottom:1px solid rgba(255,255,255,.05)">'
                f'<td style="padding:6px 8px;color:var(--text)">{_esc(_display_director(r["DIRETOR"]))}</td>'
                f'<td style="padding:6px 8px;text-align:center;color:var(--coral)">{_fmt(r["CITA_SOMA_CIT"])}</td>'
                f'<td style="padding:6px 8px;color:var(--muted)">{reason}</td>'
                f'<td style="padding:6px 8px;color:var(--muted);font-size:9px">{_esc(r["CITA_VENUES"])}</td>'
                '</tr>'
            )
        return ''.join(rows)

    fig_cat = go.Figure(go.Bar(
        x=cat_stats['pct_festival'].round(1),
        y=cat_stats['Categoria'],
        customdata=np.column_stack([cat_stats['n_obras'], cat_stats['n_festival'], cat_stats['fest_media'].round(1)]),
        orientation='h',
        marker_color='#00e5c8',
        hovertemplate='<b>%{y}</b><br>% festival: %{x:.1f}%<br>Obras: %{customdata[0]}<br>Com festival: %{customdata[1]}<br>Score medio: %{customdata[2]}<extra></extra>',
    ))
    fig_cat.update_layout(height=320, margin=dict(l=220, r=18, t=10, b=42),
                          xaxis=dict(title='% de filmes com festival', gridcolor='#1e2035'),
                          yaxis=dict(title='', gridcolor='#1e2035'),
                          paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,.6)',
                          font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10))

    yearly = crit_ok[crit_ok['Ano'].notna()].groupby('Ano').agg(media=('critica', 'mean'), n=('Projeto', 'count')).reset_index().sort_values('Ano')
    fig_year = go.Figure(go.Scatter(
        x=yearly['Ano'].astype(int) if len(yearly) else [],
        y=yearly['media'].round(2) if len(yearly) else [],
        customdata=yearly['n'] if len(yearly) else [],
        mode='lines+markers',
        line=dict(color='#f5c842', width=3),
        marker=dict(size=8, color='#f5c842'),
        hovertemplate='Ano %{x}<br>Critica media: %{y:.2f}<br>Filmes: %{customdata}<extra></extra>',
    ))
    fig_year.update_layout(height=320, margin=dict(l=48, r=18, t=10, b=42),
                           xaxis=dict(title='Ano', gridcolor='#1e2035'),
                           yaxis=dict(title='Indice critico', range=[1, 5], gridcolor='#1e2035'),
                           paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,.6)',
                           font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10))

    scat = crit_ok[crit_ok['roi_intl'].notna()].copy()
    fig_scat = go.Figure(go.Scatter(
        x=scat['critica'],
        y=scat['roi_intl'],
        text=scat['Projeto'],
        customdata=np.column_stack([scat['Ano'].fillna(''), scat['Categoria'], scat['fest_score']]),
        mode='markers',
        marker=dict(color='#a78bfa', size=7, opacity=.62, line=dict(width=.4, color='white')),
        hovertemplate='<b>%{text}</b><br>Ano: %{customdata[0]}<br>%{customdata[1]}<br>Critica: %{x:.2f}<br>ROI intl: %{y:.1f}<br>Festival: %{customdata[2]:.1f}<extra></extra>',
    ))
    fig_scat.update_layout(height=330, margin=dict(l=54, r=18, t=10, b=44),
                           xaxis=dict(title='Indice critico (1-5)', range=[1, 5], gridcolor='#1e2035'),
                           yaxis=dict(title='ROI internacional (0-100)', gridcolor='#1e2035'),
                           paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,.6)',
                           font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10))

    fig_cit = go.Figure(go.Bar(
        x=cit_top['CITA_SOMA_CIT'],
        y=cit_top['diretor_artistico'],
        orientation='h',
        marker_color='#00e5c8',
        hovertemplate='<b>%{y}</b><br>Citacoes: %{x:.0f}<extra></extra>',
    ))
    fig_cit.update_layout(height=380, margin=dict(l=180, r=18, t=10, b=42),
                          xaxis=dict(title='Citacoes OpenAlex aprovadas por venue', gridcolor='#1e2035'),
                          yaxis=dict(title='', gridcolor='#1e2035'),
                          paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(14,16,24,.6)',
                          font=dict(color='#e8eaf2', family='DM Mono, monospace', size=10))

    section_html = f'''
<div id="mega-section-sp" class="mega-panel" style="display:none">
<div style="overflow-y:auto;height:100%;padding:24px 28px 48px" id="sp-scroll">
  <div style="margin-bottom:20px">
    <div style="font-family:var(--font-head);font-size:21px;font-style:italic;color:var(--text);margin-bottom:5px">Soft Power do Cinema Brasileiro</div>
    <div style="font-size:10px;color:var(--muted);letter-spacing:.04em">Recorte amplo de filmes do painel de produtoras · categorias de fomento/festivais · citações OpenAlex auditadas por area</div>
  </div>
  <div style="display:flex;gap:8px;margin-bottom:20px;border-bottom:1px solid var(--border)">
    <button id="sp-tab-critica" onclick="spShow('critica')" style="padding:8px 16px;background:none;border:none;border-bottom:2px solid var(--accent);color:var(--accent);font-family:var(--font-mono);font-size:11px;cursor:pointer;margin-bottom:-1px">Critica + festivais</button>
    <button id="sp-tab-citacao" onclick="spShow('citacao')" style="padding:8px 16px;background:none;border:none;border-bottom:2px solid transparent;color:var(--muted);font-family:var(--font-mono);font-size:11px;cursor:pointer;margin-bottom:-1px">Citacao academica auditada</button>
  </div>

  <div id="sp-panel-critica">
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:20px">
      <div style="background:var(--surface);border:1px solid var(--border);border-top:2px solid var(--accent);border-radius:6px;padding:14px 16px"><div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">Filmes no recorte amplo</div><div style="font-family:var(--font-head);font-size:26px;color:var(--accent)">{_fmt(kpi_filmes)}</div><div style="font-size:10px;color:var(--muted);margin-top:3px">base historica, sem TV/series</div></div>
      <div style="background:var(--surface);border:1px solid var(--border);border-top:2px solid var(--gold);border-radius:6px;padding:14px 16px"><div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">Critica alta</div><div style="font-family:var(--font-head);font-size:26px;color:var(--gold)">{_fmt(kpi_crit)}</div><div style="font-size:10px;color:var(--muted);margin-top:3px">{_fmt(pct_crit,1)}% de cobertura</div></div>
      <div style="background:var(--surface);border:1px solid var(--border);border-top:2px solid var(--purple);border-radius:6px;padding:14px 16px"><div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">Indice medio</div><div style="font-family:var(--font-head);font-size:26px;color:var(--purple)">{_fmt(crit_media,2)}<span style="font-size:14px;color:var(--muted)">/5</span></div><div style="font-size:10px;color:var(--muted);margin-top:3px">apenas filmes com 2+ fontes</div></div>
      <div style="background:var(--surface);border:1px solid var(--border);border-top:2px solid var(--coral);border-radius:6px;padding:14px 16px"><div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">Com festival</div><div style="font-family:var(--font-head);font-size:26px;color:var(--coral)">{_fmt(kpi_fest)}</div><div style="font-size:10px;color:var(--muted);margin-top:3px">pontuacao de festivais > 0</div></div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px">
      <div style="background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:16px"><div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--accent);margin-bottom:10px">Categoria de fomento por presenca em festivais</div><div id="sp-chart-cat" style="height:320px"></div></div>
      <div style="background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:16px"><div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--gold);margin-bottom:10px">Evolucao da critica no recorte amplo</div><div id="sp-chart-ev" style="height:320px"></div></div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:16px;margin-bottom:14px"><div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--purple);margin-bottom:10px">Critica x ROI internacional</div><div id="sp-chart-scat" style="height:330px"></div></div>
    {_crit_table_html(top_crit, 'Filmes de cinema por indice critico - confianca alta', f'{_fmt(len(top_crit))} filmes de cinema com critica profissional, match alto e pelo menos 5 criticas profissionais. Clique em uma linha para ver fonte profissional, nota e contagem bruta.', 'top', '720px')}
    {_crit_table_html(low_crit, 'Filmes com confianca media/baixa - consultar com cautela', f'{_fmt(kpi_crit_low)} filmes de cinema com critica profissional, mas sem match alto ou sem volume minimo de criticas. Clique em uma linha para auditar o detalhe bruto.', 'low', '520px')}
  </div>

  <div id="sp-panel-citacao" style="display:none">
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:20px">
      <div style="background:var(--surface);border:1px solid var(--border);border-top:2px solid var(--accent);border-radius:6px;padding:14px 16px"><div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">Diretores aprovados</div><div style="font-family:var(--font-head);font-size:26px;color:var(--accent)">{_fmt(kpi_cit)}</div><div style="font-size:10px;color:var(--muted);margin-top:3px">com venue de cinema/comunicacao</div></div>
      <div style="background:var(--surface);border:1px solid var(--border);border-top:2px solid var(--coral);border-radius:6px;padding:14px 16px"><div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">Descartados</div><div style="font-family:var(--font-head);font-size:26px;color:var(--coral)">{_fmt(kpi_rej)}</div><div style="font-size:10px;color:var(--muted);margin-top:3px">homonimia/area externa provavel</div></div>
      <div style="background:var(--surface);border:1px solid var(--border);border-top:2px solid var(--purple);border-radius:6px;padding:14px 16px"><div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">Criterio</div><div style="font-family:var(--font-head);font-size:20px;color:var(--purple);line-height:1.15">venue<br>auditado</div><div style="font-size:10px;color:var(--muted);margin-top:3px">OpenAlex nao foi aceito no bruto</div></div>
      <div style="background:var(--surface);border:1px solid var(--border);border-top:2px solid var(--gold);border-radius:6px;padding:14px 16px"><div style="font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:6px">Nome exibido</div><div style="font-family:var(--font-head);font-size:20px;color:var(--gold);line-height:1.15">artistico</div><div style="font-size:10px;color:var(--muted);margin-top:3px">alias de festivais/prestigio</div></div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:16px;margin-bottom:14px"><div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--accent);margin-bottom:10px">Top diretores por citacoes aprovadas</div><div id="sp-chart-topcit" style="height:380px"></div></div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:16px;margin-bottom:14px"><div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:12px">Ranking auditado - nome artistico da direcao</div><div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:10px"><thead><tr style="border-bottom:1px solid var(--border)"><th style="padding:6px 8px;text-align:left;color:var(--muted)">#</th><th style="padding:6px 8px;text-align:left;color:var(--muted)">Direcao</th><th style="padding:6px 8px;text-align:center;color:var(--muted)">Cit.</th><th style="padding:6px 8px;text-align:center;color:var(--muted)">Papers</th><th style="padding:6px 8px;text-align:center;color:var(--muted)">h*</th><th style="padding:6px 8px;text-align:left;color:var(--muted)">Venues aceitos</th></tr></thead><tbody>{_rows_top_cit()}</tbody></table></div></div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:16px"><div style="font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--coral);margin-bottom:12px">Auditoria - registros removidos do ranking</div><div style="overflow-x:auto;max-height:360px"><table style="width:100%;border-collapse:collapse;font-size:10px"><thead><tr style="border-bottom:1px solid var(--border)"><th style="padding:6px 8px;text-align:left;color:var(--muted)">Nome consultado</th><th style="padding:6px 8px;text-align:center;color:var(--muted)">Cit.</th><th style="padding:6px 8px;text-align:left;color:var(--muted)">Motivo</th><th style="padding:6px 8px;text-align:left;color:var(--muted)">Venues</th></tr></thead><tbody>{_rows_audit()}</tbody></table></div><div style="font-size:9px;color:var(--muted);line-height:1.6;margin-top:10px">A auditoria remove registros sem venue de cinema, comunicacao, audiovisual ou humanidades correlatas. Registros com venues mistos ficam no ranking porque ha evidencia da area, mas aparecem marcados como MISTO para indicar risco de contagem inflada por homonimia no OpenAlex agregado.</div></div>
  </div>
</div>
</div>'''

    render_script = f'''
<script>
(function(){{
  var figCat={pio.to_json(fig_cat)};
  var figEv={pio.to_json(fig_year)};
  var figScat={pio.to_json(fig_scat)};
  var figTopCit={pio.to_json(fig_cit)};
  window.spToggleCritDetail = function(id){{
    var el = document.getElementById(id);
    if(!el) return;
    el.style.display = el.style.display === 'none' ? 'table-row' : 'none';
  }};
  function plotFig(id, fig){{
    var el = document.getElementById(id);
    if(!el || !window.Plotly) return;
    Plotly.react(el, fig.data, fig.layout, {{responsive:true,displaylogo:false}});
  }}
  window.spRender = function(which){{
    if(which === 'citacao') {{
      plotFig('sp-chart-topcit', figTopCit);
    }} else {{
      plotFig('sp-chart-cat', figCat);
      plotFig('sp-chart-ev', figEv);
      plotFig('sp-chart-scat', figScat);
    }}
    _resizeVisibleCharts('mega-section-sp');
  }};
  window.spShow = function(which){{
    var critica = which !== 'citacao';
    var pc = document.getElementById('sp-panel-critica');
    var pa = document.getElementById('sp-panel-citacao');
    var tc = document.getElementById('sp-tab-critica');
    var ta = document.getElementById('sp-tab-citacao');
    if(pc) pc.style.display = critica ? 'block' : 'none';
    if(pa) pa.style.display = critica ? 'none' : 'block';
    if(tc) {{
      tc.style.color = critica ? 'var(--accent)' : 'var(--muted)';
      tc.style.borderBottomColor = critica ? 'var(--accent)' : 'transparent';
    }}
    if(ta) {{
      ta.style.color = critica ? 'var(--muted)' : 'var(--accent)';
      ta.style.borderBottomColor = critica ? 'transparent' : 'var(--accent)';
    }}
    window.spRender(critica ? 'critica' : 'citacao');
  }};
}})();
</script>
'''
    print(f"Soft Power v2: {kpi_filmes} filmes no recorte amplo, {kpi_cit} diretores aprovados, {kpi_rej} descartados na auditoria")
    return section_html, render_script

try:
    _sp_section, _sp_render_script = _build_soft_power_panel_v2()
except Exception as e:
    print(f"AVISO Soft Power v2: {e}")

print("Sections built")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 12 — Assemble mega HTML
# ─────────────────────────────────────────────────────────────────────────────

mega_html = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>FSA Cinema — Painel Integrado de Análise</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Mono:wght@300;400;500&display=swap" rel="stylesheet">
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<script src="https://d3js.org/d3.v7.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/topojson@3/dist/topojson.min.js"></script>
<style>
/* ── BASE: dark theme ── */
:root{{
  --bg:#07080f;--surface:#0e1018;--surface2:#141620;--border:#1e2035;
  --accent:#00e5c8;--accent-dim:rgba(0,229,200,.15);
  --gold:#f5c842;--coral:#ff7c6e;--purple:#a78bfa;--muted-blue:#5fd1ff;
  --text:#e8eaf2;--muted:#5a6080;--dim:#1e2035;
  --font-head:'DM Serif Display',serif;--font-mono:'DM Mono',monospace;
  --sidebar-w:210px;--sidebar-collapsed-w:0px;
  --transition-speed:.25s;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:var(--font-mono);font-size:13px;min-height:100vh}}

/* ── Header ── */
.mega-header{{
  padding:0;border-bottom:1px solid var(--border);
  background:linear-gradient(180deg,#0c0e18 0%,var(--bg) 100%);
  display:flex;flex-direction:column;flex-shrink:0;
}}
.mega-header-top{{
  display:flex;align-items:center;gap:20px;padding:16px 32px 12px;
}}
.mega-header-icon{{
  width:44px;height:44px;border-radius:12px;background:var(--accent);
  display:flex;align-items:center;justify-content:center;flex-shrink:0;
  box-shadow:0 0 20px rgba(0,229,200,.15);
}}
.mega-header-icon svg{{width:20px;height:20px}}
.mega-header-titles{{flex:1}}
.mega-header-titles h1{{
  font-family:var(--font-head);font-size:21px;font-weight:400;letter-spacing:-.3px;line-height:1.15
}}
.mega-header-titles p{{font-size:11px;color:var(--muted);margin-top:3px;letter-spacing:.02em}}
.mega-header-meta{{text-align:right;font-size:10px;color:var(--muted);line-height:1.7}}

/* Methodology ribbon — collapsible */
.mega-meto-bar{{
  display:flex;align-items:center;gap:0;padding:0 32px;
  border-top:1px solid var(--border);background:rgba(255,255,255,.02);
  overflow:hidden;max-height:60px;transition:max-height var(--transition-speed) ease;
}}
.mega-meto-bar.collapsed{{max-height:0;border-top-color:transparent}}
.meto-toggle{{
  position:absolute;right:32px;top:0;height:100%;
  display:flex;align-items:center;background:none;border:none;
  color:var(--muted);cursor:pointer;font-family:var(--font-mono);font-size:9px;
  letter-spacing:.08em;text-transform:uppercase;gap:4px;padding:0 4px;
  transition:color .15s;
}}
.meto-toggle:hover{{color:var(--text)}}
.meto-toggle svg{{width:10px;height:10px;transition:transform var(--transition-speed)}}
.mega-meto-bar.collapsed ~ .meto-toggle svg{{transform:rotate(180deg)}}
.mega-meto-wrap{{position:relative;border-top:1px solid var(--border)}}
.mega-meto-wrap .meto-toggle{{top:0;height:32px}}
.meto-item{{
  display:flex;align-items:baseline;gap:7px;
  padding:8px 18px 8px 0;
  border-right:1px solid var(--border);margin-right:18px;
  white-space:nowrap;font-size:10px;
}}
.meto-item:last-child{{border-right:none;margin-right:0}}
.meto-label{{
  font-size:9px;letter-spacing:.1em;text-transform:uppercase;
  color:var(--accent);font-weight:600;
}}
.meto-def{{color:var(--muted);line-height:1.5}}
.meto-def strong{{color:var(--text)}}

/* ── Layout ── */
body{{display:flex;flex-direction:column;height:100vh;overflow:hidden}}
.mega-layout{{display:flex;flex:1;overflow:hidden}}

/* ── Sidebar navigation ── */
.mega-sidebar{{
  width:var(--sidebar-w);min-width:var(--sidebar-w);
  background:var(--surface);border-right:1px solid var(--border);
  display:flex;flex-direction:column;overflow-y:auto;padding:8px 0;
  transition:width var(--transition-speed) ease,min-width var(--transition-speed) ease,
             opacity var(--transition-speed) ease,transform var(--transition-speed) ease;
}}
body.sidebar-collapsed .mega-sidebar{{
  width:0;min-width:0;opacity:0;overflow:hidden;padding:0;border-right:none;
}}
.mega-sidebar-label{{padding:14px 18px 6px;font-size:9px;color:var(--muted);letter-spacing:.1em;text-transform:uppercase}}

.mega-tab{{
  display:flex;align-items:center;gap:10px;width:100%;text-align:left;
  padding:10px 18px;font-size:10.5px;color:var(--muted);cursor:pointer;
  border-left:2px solid transparent;transition:all .15s;
  letter-spacing:.04em;text-transform:uppercase;font-family:var(--font-mono);
  white-space:normal;background:none;border-top:none;border-bottom:none;border-right:none;
  line-height:1.4;
}}
.mega-tab .tab-icon{{
  width:16px;height:16px;flex-shrink:0;opacity:.5;transition:opacity .15s;
}}
.mega-tab.active .tab-icon{{opacity:1}}
.mega-tab .tab-key{{
  margin-left:auto;font-size:9px;color:var(--dim);
  border:1px solid var(--border);border-radius:3px;
  padding:1px 5px;font-family:var(--font-mono);line-height:1.4;
}}
.mega-tab.active{{color:var(--accent);border-left-color:var(--accent);background:var(--accent-dim)}}
.mega-tab:hover:not(.active){{color:var(--text);background:rgba(255,255,255,.03)}}
.mega-tab:hover .tab-key{{color:var(--muted)}}

/* Sidebar toggle */
.sidebar-toggle{{
  position:fixed;left:var(--sidebar-w);top:50%;transform:translateY(-50%);
  z-index:50;width:18px;height:38px;background:var(--surface);
  border:1px solid var(--border);border-left:none;
  border-radius:0 6px 6px 0;cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  color:var(--muted);transition:left var(--transition-speed) ease,color .15s,background .15s;
}}
.sidebar-toggle:hover{{color:var(--accent);background:var(--surface2)}}
.sidebar-toggle svg{{width:12px;height:12px;transition:transform var(--transition-speed)}}
body.sidebar-collapsed .sidebar-toggle{{left:0}}
body.sidebar-collapsed .sidebar-toggle svg{{transform:rotate(180deg)}}

.mega-content{{flex:1;overflow-y:auto;position:relative}}
.mega-panel{{display:none;min-height:100%}}
.mega-panel.active{{display:block}}

/* Scroll-to-top button */
.scroll-top{{
  position:fixed;bottom:24px;right:24px;z-index:100;
  width:38px;height:38px;border-radius:50%;
  background:var(--surface2);border:1px solid var(--border);
  color:var(--muted);cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  opacity:0;pointer-events:none;
  transition:opacity .2s,background .15s,color .15s,transform .15s;
  box-shadow:0 4px 12px rgba(0,0,0,.3);
}}
.scroll-top.visible{{opacity:1;pointer-events:auto}}
.scroll-top:hover{{background:var(--accent);color:#000;transform:scale(1.08)}}
.scroll-top svg{{width:16px;height:16px}}

/* ── Visão Geral sub-nav (dark) ── */
.vg-subnav{{
  display:flex;padding:0 28px;background:var(--surface);
  border-bottom:1px solid var(--border);position:sticky;top:0;z-index:10;
}}
.vg-subtab{{
  padding:11px 18px;font-size:11px;color:var(--muted);cursor:pointer;
  border-bottom:2px solid transparent;transition:all .2s;
  letter-spacing:.06em;text-transform:uppercase;font-family:var(--font-mono);
  background:none;border-top:none;border-left:none;border-right:none;
  margin-bottom:-1px;white-space:nowrap;
}}
.vg-subtab.active{{color:var(--accent);border-bottom-color:var(--accent)}}
.vg-subtab:hover:not(.active){{color:var(--text)}}

/* ── Visão Geral: dark overrides (match panel theme) ── */
#mega-section-cmp{{background:var(--bg)}}
#mega-section-cmp .tab-content{{display:block!important;background:var(--bg)!important;padding:24px 28px!important}}
#mega-section-cmp .card{{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important;border-radius:8px!important}}
#mega-section-cmp .section-title{{color:var(--accent)!important;border-bottom-color:var(--border)!important}}
#mega-section-cmp .kpi-card{{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important}}
#mega-section-cmp .kpi-card .val{{color:var(--accent)!important}}
#mega-section-cmp .kpi-card .label{{color:var(--muted)!important}}
#mega-section-cmp .kpi-card .sub{{color:var(--muted)!important}}
#mega-section-cmp table thead th{{color:var(--muted)!important;border-bottom-color:var(--border)!important;background:var(--bg)!important}}
#mega-section-cmp table tbody td{{border-bottom-color:var(--border)!important;color:var(--text)!important}}
#mega-section-cmp table tbody tr:hover{{background:var(--surface2)!important}}
#mega-section-cmp h3,#mega-section-cmp h4{{color:var(--text)!important}}
#mega-section-cmp .insight-box,#mega-section-cmp .alert-box{{background:var(--surface2)!important;border-color:var(--border)!important;color:var(--text)!important}}
#mega-section-cmp select{{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}}
#mega-section-cmp input{{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}}
#mega-section-cmp .topbar{{display:none!important}}
#mega-section-cmp .tabs{{display:none!important}}
#mega-section-cmp .legenda{{background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:8px}}
#mega-section-cmp .legenda-item{{color:var(--text)!important}}
#mega-section-cmp .legenda-item span{{color:var(--muted)!important}}
#mega-section-cmp .chart-wrap{{background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:12px!important}}
#mega-section-cmp .chart-title{{color:var(--text)!important;font-family:var(--font-head)!important}}
#mega-section-cmp .ctrl-btn{{border-color:var(--border)!important;color:var(--muted)!important;background:transparent!important}}
#mega-section-cmp .ctrl-btn.active,#mega-section-cmp .ctrl-btn:hover{{background:var(--accent)!important;color:#000!important;border-color:var(--accent)!important;font-weight:600!important}}
#mega-section-cmp .tog-btn{{border-color:var(--border)!important;color:var(--muted)!important;background:var(--surface2)!important}}
#mega-section-cmp .tog-btn.active{{background:var(--accent)!important;color:#000!important;border-color:var(--accent)!important}}
#mega-section-cmp .tog-btn:hover:not(.active){{background:var(--surface)!important;color:var(--text)!important}}
#mega-section-cmp .search-input{{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}}
#mega-section-cmp .search-input:focus{{border-color:var(--accent)!important;box-shadow:0 0 0 2px rgba(0,229,200,.18)!important}}
#mega-section-cmp .search-count{{color:var(--muted)!important}}
#mega-section-cmp .mod-bar-track{{background:var(--surface2)!important}}
#mega-section-cmp .mod-bar-fill{{color:#000!important}}
/* cluster cards */
#mega-section-cmp .cl-card{{background:var(--surface)!important;border:1px solid var(--border)!important}}
#mega-section-cmp .cl-card:hover{{background:var(--surface2)!important}}
#mega-section-cmp .cl-name{{color:var(--text)!important}}
#mega-section-cmp .cl-n{{color:var(--text)!important}}
#mega-section-cmp .cl-desc,#mega-section-cmp .cl-stat span,#mega-section-cmp .cl-top,#mega-section-cmp .cl-pct{{color:var(--muted)!important}}
#mega-section-cmp .cl-stat b{{color:var(--text)!important}}
#mega-section-cmp .cl-top b{{color:var(--accent)!important}}
#mega-section-cmp .ov-rank-head{{border-bottom-color:var(--border)!important}}
#mega-section-cmp .bc-nm{{color:var(--text)!important}}
#mega-section-cmp .bc-val{{color:var(--text)!important}}
#mega-section-cmp .bc-rank,#mega-section-cmp .bc-sub{{color:var(--muted)!important}}
#mega-section-cmp .bc-bar-wrap{{background:var(--surface2)!important}}
#mega-section-cmp .quad-wrap{{background:var(--surface)!important;border:1px solid var(--border)!important}}
#mega-section-cmp .filter-bar select,#mega-section-cmp .filter-bar input{{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}}
#mega-section-cmp .bar-bg{{background:rgba(30,32,53,.8)!important}}
#mega-section-cmp .comp-card{{background:var(--surface)!important;border:1px solid var(--border)!important}}
#mega-section-cmp .comp-row .label{{color:var(--muted)!important}}
#mega-section-cmp .comp-row .val{{color:var(--text)!important}}
#mega-section-cmp .quad-ctrl .lbl{{color:var(--muted)!important}}
#mega-section-cmp .obras-drawer{{background:var(--surface)!important;border:1px solid var(--border)!important}}
#mega-section-cmp .obras-drawer-title{{color:var(--text)!important}}
#mega-section-cmp .cl-panel{{background:var(--surface2)!important;border:1px solid var(--border)!important}}
#mega-section-cmp .cat-pill{{background:var(--surface)!important;border-color:var(--border)!important;color:var(--muted)!important}}
#mega-section-cmp .cat-pill.active{{background:var(--accent-dim)!important;border-color:var(--accent)!important;color:var(--accent)!important}}

/* ── Retorno Internacional: override inline light-theme styles ── */
#cmp-panel-ret-intl *{{color:var(--text)}}
#cmp-panel-ret-intl [style*="background:#fff"],
#cmp-panel-ret-intl [style*="background: #fff"]{{background:var(--bg)!important;color:var(--text)!important}}
#cmp-panel-ret-intl [style*="background:#f7f8fb"]{{background:var(--surface)!important;border-color:var(--border)!important}}
#cmp-panel-ret-intl [style*="background:#f0f4f8"]{{background:var(--surface)!important;border-color:var(--border)!important}}
#cmp-panel-ret-intl [style*="background:#eef0f4"]{{background:var(--surface2)!important}}
#cmp-panel-ret-intl [style*="background:rgba(255,255,255"]{{background:var(--surface)!important;border-color:var(--border)!important}}
#cmp-panel-ret-intl [style*="color:#222"]{{color:var(--text)!important}}
#cmp-panel-ret-intl [style*="color:#333"]{{color:var(--text)!important}}
#cmp-panel-ret-intl [style*="color:#444"]{{color:var(--text)!important}}
#cmp-panel-ret-intl [style*="color:#555"]{{color:var(--muted)!important}}
#cmp-panel-ret-intl [style*="color:#666"]{{color:var(--muted)!important}}
#cmp-panel-ret-intl [style*="color:#888"]{{color:var(--muted)!important}}
#cmp-panel-ret-intl [style*="border:1px solid #dde0e8"]{{border-color:var(--border)!important}}
#cmp-panel-ret-intl [style*="border:1px solid #ccd0da"]{{border-color:var(--border)!important}}
#cmp-panel-ret-intl [style*="border:1px solid #ddd"]{{border-color:var(--border)!important}}
#cmp-panel-ret-intl [style*="border:1px solid #5B6BB5"]{{border-color:var(--accent)!important}}
#cmp-panel-ret-intl #intl-tooltip{{background:var(--surface2)!important;border-color:var(--accent)!important;color:var(--text)!important}}

/* ── Concentracao section (inside Produtoras) ── */
#conc-section .conc-subnav{{display:flex;gap:0;border-bottom:1px solid var(--border);margin-bottom:20px}}
#conc-section .conc-tab{{padding:9px 16px;font-size:10px;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;letter-spacing:.06em;text-transform:uppercase;font-family:var(--font-mono);background:none;border-top:none;border-left:none;border-right:none;white-space:nowrap}}
#conc-section .conc-tab.active{{color:var(--accent);border-bottom-color:var(--accent)}}
#conc-section .tab-panel{{display:none;padding:0 4px;flex-direction:column}}
#conc-section .tab-panel.active{{display:flex}}
#conc-section .scroll{{overflow:visible!important;height:auto!important}}

/* ── Criterio selecao ── */
#mega-section-cs .header{{display:none!important}}
#mega-section-cs .tabs{{background:var(--surface);border-bottom:1px solid var(--border);padding:0 20px}}
#mega-section-cs .tab{{font-size:10px}}
#mega-section-cs .panel{{padding:24px 28px}}
#mega-section-cs #cs-tooltip{{position:fixed;display:none;background:var(--surface2);border:1px solid var(--accent);border-radius:8px;padding:12px 14px;font-size:11px;line-height:1.7;pointer-events:none;z-index:9999;max-width:310px}}
/* Dark overrides for "Por Cluster" panel inside Produtoras */
#pr-tab-clusters .card{{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important;border-radius:8px!important}}
#pr-tab-clusters .card div{{color:var(--text)!important}}
#pr-tab-clusters .search-input{{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}}
#pr-tab-clusters .search-count{{color:var(--muted)!important}}
#pr-tab-clusters table thead th{{color:var(--muted)!important;background:var(--bg)!important}}
#pr-tab-clusters table tbody td{{color:var(--text)!important;border-bottom-color:var(--border)!important}}

/* ── Produtoras dark overrides ── */
#mega-section-pr{{
  overflow:hidden!important;
  flex-direction:column;
  height:100%;
  min-height:0;
}}
#mega-section-pr .hdr{{display:none!important}}
#mega-section-pr .tab-bar{{background:var(--surface);border-bottom:1px solid var(--border);flex-shrink:0}}
#mega-section-pr .tab-btn{{color:var(--muted);background:none;border-color:transparent}}
#mega-section-pr .tab-btn.active{{color:var(--accent);border-bottom-color:var(--accent)!important}}
#mega-section-pr .tab-panel{{background:var(--bg)}}
/* Ticket panel inside produtoras */
#pr-ticket-panel{{background:var(--bg)!important;padding:0!important;overflow-y:auto!important;flex-direction:column!important}}
/* Concentração panel inside produtoras */
#pr-conc-panel{{background:var(--bg)!important;padding:0!important;overflow-y:auto!important;flex-direction:column!important}}
#pr-ticket-panel .cmp-tab-panel{{display:block!important}}
#pr-ticket-panel .tab-content{{display:block!important;background:var(--bg)!important;padding:14px 18px!important}}
#pr-ticket-panel .card{{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important}}
#pr-ticket-panel .section-title{{color:var(--accent)!important;border-bottom-color:var(--border)!important}}
#pr-ticket-panel .kpi-card{{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important}}
#pr-ticket-panel .kpi-card .val{{color:var(--accent)!important}}
#pr-ticket-panel table thead th{{color:var(--muted)!important;border-bottom-color:var(--border)!important;background:var(--bg)!important}}
#pr-ticket-panel table tbody td{{border-bottom-color:var(--border)!important;color:var(--text)!important}}
#pr-ticket-panel table tbody tr:hover{{background:var(--surface2)!important}}
#pr-ticket-panel h3,#pr-ticket-panel h4{{color:var(--text)!important}}

/* ── Responsive: small screens ── */
@media (max-width: 900px) {{
  .mega-sidebar{{position:fixed;left:0;top:0;bottom:0;z-index:200;transform:translateX(-100%);width:240px;min-width:240px}}
  .mega-sidebar.mobile-open{{transform:translateX(0);box-shadow:4px 0 24px rgba(0,0,0,.5)}}
  body.sidebar-collapsed .mega-sidebar{{width:240px;min-width:240px;opacity:1;padding:8px 0}}
  .sidebar-toggle{{display:none}}
  .mobile-menu-btn{{display:flex!important}}
  .mega-header-top{{padding:12px 16px 10px}}
  .mega-meto-bar{{padding:0 16px}}
  .mobile-overlay{{display:block!important}}
  .mobile-overlay.active{{opacity:1;pointer-events:auto}}
}}
@media (min-width: 901px) {{
  .mobile-menu-btn{{display:none!important}}
  .mobile-overlay{{display:none!important}}
}}

/* ── UI stabilization: viewport, scroll containers, responsive tabs ── */
html{{height:100%;max-width:100%}}
body{{height:100vh;height:100dvh;min-height:0;overflow:hidden;max-width:100vw}}
.mega-header-titles{{min-width:0}}
.mega-header-titles h1,.mega-header-titles p{{overflow-wrap:anywhere}}
.mega-meto-bar{{overflow-x:auto;overflow-y:hidden;scrollbar-width:thin}}
.meto-item{{flex-shrink:0}}
.mega-layout{{min-height:0;width:100%}}
.mega-content{{min-width:0;min-height:0;overflow-x:hidden;overscroll-behavior:contain;-webkit-overflow-scrolling:touch}}
.mega-panel{{min-width:0}}
#mega-section-pr{{height:100%}}
#mega-section-cmp,#mega-section-cs,#mega-section-cl,#mega-section-div,#mega-section-sp{{min-height:100%}}
.vg-subnav,#mega-section-cs .tabs,#mega-section-pr .tab-bar,#conc-section .conc-subnav{{
  overflow-x:auto;overflow-y:hidden;scrollbar-width:thin;
}}
.vg-subnav::-webkit-scrollbar,#mega-section-cs .tabs::-webkit-scrollbar,
#mega-section-pr .tab-bar::-webkit-scrollbar,#conc-section .conc-subnav::-webkit-scrollbar{{height:5px}}
.vg-subtab,#mega-section-cs .tab,#mega-section-pr .tab-btn,#conc-section .conc-tab{{flex-shrink:0}}
.card,.kpi,.kpi-card,.chart-wrap,.quad-wrap,.comp-card{{min-width:0}}
#mega-section-cmp .card,#mega-section-cs .panel,#pr-ticket-panel .card,#conc-section .card{{overflow-x:auto}}
.grid-2,.kpi-bar,#conc-section .grid2{{min-width:0}}
.plotly-graph-div{{max-width:100%}}
#mega-section-pr .tab-panel.main-tab.active{{overflow-y:auto}}
@media (max-width: 900px) {{
  .mega-header-top{{gap:10px;align-items:flex-start;padding:12px 14px 10px}}
  .mega-header-icon{{width:40px;height:40px;border-radius:10px}}
  .mega-header-titles h1{{font-size:15px;line-height:1.2}}
  .mega-header-titles p{{font-size:10px;line-height:1.45}}
  .mega-header-meta{{display:none}}
  .mega-meto-wrap{{display:none}}
  .mega-content{{width:100%}}
  .vg-subnav{{padding:0 12px}}
  #mega-section-cmp .tab-content,#mega-section-cs .panel{{padding:16px 14px!important}}
  #pr-tab-clusters{{padding:16px 14px!important}}
  .grid-2,.kpi-bar,#conc-section .grid2{{grid-template-columns:1fr!important}}
  .plotly-graph-div{{min-width:680px;min-height:280px}}
  .modal{{width:calc(100vw - 24px);max-height:calc(100dvh - 24px)}}
}}

/* ── CMP CSS (scoped) ── */
{cmp_css}

/* ── CS CSS (scoped - already dark) ── */
{cs_css}

/* ── PR CSS ── */
{pr_css}

/* ── CONC CSS (scoped to #conc-section) ── */
{conc_css_scoped}

/* ── CL/DIV/SP dark overrides + utility classes ── */
{_cl_div_sp_css}
</style>
</head>
<body>

<!-- ── Header ── -->
<div class="mega-header">
  <div class="mega-header-top">
    <!-- Mobile menu button -->
    <button class="mobile-menu-btn" style="display:none;align-items:center;justify-content:center;width:36px;height:36px;background:none;border:1px solid var(--border);border-radius:8px;color:var(--text);cursor:pointer;flex-shrink:0" onclick="toggleMobileSidebar()">
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><line x1="2" y1="4" x2="14" y2="4"/><line x1="2" y1="8" x2="14" y2="8"/><line x1="2" y1="12" x2="14" y2="12"/></svg>
    </button>
    <div class="mega-header-icon">
      <svg viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2.2">
        <rect x="2" y="7" width="20" height="14" rx="2"/>
        <path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/>
        <line x1="12" y1="12" x2="12" y2="15"/>
        <line x1="10.5" y1="13.5" x2="13.5" y2="13.5"/>
      </svg>
    </div>
    <div class="mega-header-titles">
      <h1>FSA Cinema — Painel Integrado de Análise</h1>
      <p>Fundo Setorial do Audiovisual · PRODECINE 2014–2023 · ANCINE · FSA/BRDE · SALIC/MinC · Lumière/CNC</p>
    </div>
    <div class="mega-header-meta">
      {N_OBRAS} obras com estreia em cartaz · PRODECINE 2014–2023<br>
      Bilheteria > 0 · dados deflacionados R$2024
    </div>
  </div>
  <div class="mega-meto-wrap">
    <div class="mega-meto-bar" id="meto-bar">
      <div class="meto-item">
        <span class="meto-label">ROI Doméstico</span>
        <span class="meto-def">
          <strong>Média ponderada</strong> pelo investimento deflacionado (R$2024) —
          receita doméstica proporcional ao FSA ÷ FSA deflacionado · obras com dados de bilheteria
        </span>
      </div>
      <div class="meto-item">
        <span class="meto-label">ROI Internacional</span>
        <span class="meto-def">
          <strong>Média incondicional</strong> (inclui zeros) —
          receita intl. estimada ÷ investimento FSA · cobre bilheteria EU (Lumière/CNC),
          países VOD e janelas complementares
        </span>
      </div>
      <div class="meto-item">
        <span class="meto-label">Deflação</span>
        <span class="meto-def">
          IPCA acumulado · base <strong>dezembro 2024</strong> · valores históricos atualizados para R$2024
        </span>
      </div>
      <div class="meto-item">
        <span class="meto-label">Fontes</span>
        <span class="meto-def">
          ANCINE (SPO/SAM) · BRDE/FSA (editais PRODECINE) · SALIC/MinC (renúncia fiscal) · Lumière/CNC (bilheteria EU)
        </span>
      </div>
    </div>
    <button class="meto-toggle" onclick="toggleMeto()" title="Mostrar/ocultar metodologia">
      <span>Metodologia</span>
      <svg viewBox="0 0 10 6" fill="none" stroke="currentColor" stroke-width="1.5"><polyline points="1,1 5,5 9,1"/></svg>
    </button>
  </div>
</div>

<!-- Mobile overlay -->
<div class="mobile-overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:150;opacity:0;pointer-events:none;transition:opacity .2s" onclick="toggleMobileSidebar()"></div>

<div class="mega-layout">

<!-- ── Sidebar toggle (desktop) ── -->
<button class="sidebar-toggle" onclick="toggleSidebar()" title="Colapsar/expandir sidebar">
  <svg viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.5"><polyline points="8,2 4,6 8,10"/></svg>
</button>

<!-- ── Sidebar navigation ── -->
<nav class="mega-sidebar" id="mega-sidebar">
  <div class="mega-sidebar-label">Painéis</div>
  <button class="mega-tab active" onclick="megaShow('vg')">
    <svg class="tab-icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><rect x="1" y="1" width="14" height="14" rx="2"/><line x1="1" y1="6" x2="15" y2="6"/><line x1="6" y1="6" x2="6" y2="15"/></svg>
    Visão Geral
    <span class="tab-key">1</span>
  </button>
  <button class="mega-tab" onclick="megaShow('cs')">
    <svg class="tab-icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><circle cx="8" cy="8" r="6"/><line x1="8" y1="2" x2="8" y2="8"/><line x1="8" y1="8" x2="12" y2="10"/></svg>
    Categorias das Chamadas
    <span class="tab-key">2</span>
  </button>
  <button class="mega-tab" onclick="megaShow('pr')">
    <svg class="tab-icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><rect x="2" y="4" width="12" height="10" rx="1"/><path d="M5 4V3a3 3 0 0 1 6 0v1"/></svg>
    Produtoras
    <span class="tab-key">3</span>
  </button>
  <button class="mega-tab" onclick="megaShow('conc')">
    <svg class="tab-icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><polyline points="2,14 2,8 6,8 6,5 10,5 10,2 14,2 14,14"/></svg>
    Concentração
    <span class="tab-key">4</span>
  </button>
  <button class="mega-tab" onclick="megaShow('cl')">
    <svg class="tab-icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><circle cx="3.5" cy="8" r="2"/><line x1="5.5" y1="8" x2="10.5" y2="8"/><polyline points="8.5,5.5 11,8 8.5,10.5"/><circle cx="12.5" cy="8" r="2"/></svg>
    Curtas → Longas
    <span class="tab-key">5</span>
  </button>
  <button class="mega-tab" onclick="megaShow('div')">
    <svg class="tab-icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><circle cx="5" cy="5" r="2.5"/><circle cx="11" cy="5" r="2.5"/><circle cx="5" cy="11" r="2.5"/><circle cx="11" cy="11" r="2.5"/></svg>
    Diversidade
    <span class="tab-key">6</span>
  </button>
  <button class="mega-tab" onclick="megaShow('sp')">
    <svg class="tab-icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M8 2L9.5 6h4.5l-3.5 2.5 1.3 4.5L8 10.5 4.2 13l1.3-4.5L2 6h4.5z"/></svg>
    Soft Power
    <span class="tab-key">7</span>
  </button>
</nav>

<!-- ── Content area ── -->
<div class="mega-content" id="mega-content">

<!-- ── SECTION: VISÃO GERAL ── -->
{cmp_section_html}

<!-- ── SECTION: CATEGORIAS DAS CHAMADAS ── -->
{cs_section_html}

<!-- ── SECTION: PRODUTORAS ── -->
{pr_section_html}

<!-- Concentração: handled inside Produtoras tab via prShowConc() -->

<!-- ── SECTION: CURTAS → LONGAS ── -->
<div id="mega-section-cl" class="mega-panel" style="display:none">
{_cl_html}
</div>

{_div_section}

{_sp_section}

</div><!-- /mega-content -->
</div><!-- /mega-layout -->

<!-- Scroll-to-top -->
<button class="scroll-top" id="scroll-top" onclick="document.getElementById('mega-content').scrollTo({{top:0,behavior:'smooth'}})" title="Voltar ao topo">
  <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><polyline points="4,10 8,4 12,10"/></svg>
</button>

<!-- ── MEGA NAVIGATION SCRIPT ── -->
<script>
const MEGA_SECTIONS = {{
  'vg':'cmp', 'cs':'cs', 'pr':'pr', 'conc':'pr', 'cl':'cl', 'div':'div', 'sp':'sp'
}};

function _resizeVisibleCharts(rootId) {{
  var root = rootId ? document.getElementById(rootId) : document;
  if(!root) return;
  setTimeout(function() {{
    if(window.Plotly && Plotly.Plots && Plotly.Plots.resize) {{
      root.querySelectorAll('.plotly-graph-div').forEach(function(div) {{
        if(div.offsetParent !== null && div._fullLayout) Plotly.Plots.resize(div);
      }});
    }}
    window.dispatchEvent(new Event('resize'));
  }}, 80);
}}

function megaShow(id) {{
  // Hide all sections
  ['cmp','cs','pr','cl','div','sp'].forEach(s => {{
    const el = document.getElementById('mega-section-'+s);
    if(el) el.style.display = 'none';
  }});
  // Deactivate all mega tabs
  document.querySelectorAll('.mega-tab').forEach(t => t.classList.remove('active'));

  const section = MEGA_SECTIONS[id];
  const sectionEl = document.getElementById('mega-section-'+section);
  if(sectionEl) sectionEl.style.display = (section === 'pr') ? 'flex' : 'block';

  // For vg: show visao-geral sub-panel by default and init charts
  if(id === 'vg') {{
    document.querySelectorAll('#mega-section-cmp .cmp-tab-panel').forEach(p => p.style.display = 'none');
    var vg = document.getElementById('cmp-panel-visao-geral');
    if(vg) vg.style.display = 'block';
    vgInitCharts('visao-geral');
    setTimeout(vgFixInlineCharts, 80);
    // Update vg-subtab active state
    document.querySelectorAll('.vg-subtab').forEach(t => {{
      t.classList.toggle('active', t.getAttribute('onclick').includes("'visao-geral'"));
    }});
  }}

  // For cs: re-draw canvas charts on first show (they rendered with 0 dimensions
  // while mega-section-cs was display:none)
  if(id === 'cs') {{
    setTimeout(function() {{
      if(typeof drawQuadrant === 'function')  {{ drawQuadrant(); }}
      if(typeof drawRankings === 'function')  {{ drawRankings(); }}
      if(typeof drawTimeline === 'function')  {{ drawTimeline(); }}
    }}, 80);
  }}

  // For pr: activate tab0 first so canvas has real dimensions, then redraw
  if(id === 'pr') {{
    if(typeof prSwitchTab === 'function') prSwitchTab(0);
    setTimeout(function() {{
      if(typeof drawQuad === 'function')     {{ drawQuad(); }}
      if(typeof renderOvRank === 'function') {{ renderOvRank(); }}
    }}, 80);
  }}

  if(id === 'conc') {{
    setTimeout(function() {{
      if(typeof prShowConc === 'function') prShowConc();
    }}, 50);
  }}

  if(id === 'cl') {{
    setTimeout(function() {{
      if(typeof window.clRender === 'function') window.clRender();
      _resizeVisibleCharts('mega-section-cl');
    }}, 80);
  }}

  if(id === 'div') {{
    setTimeout(function() {{
      if(typeof window.divRender === 'function') window.divRender();
    }}, 80);
  }}

  if(id === 'sp') {{
    setTimeout(function() {{
      if(typeof window.spShow === 'function') window.spShow('critica');
      else if(typeof window.spRender === 'function') window.spRender('critica');
    }}, 80);
  }}

  // Activate the correct mega tab
  document.querySelectorAll('.mega-tab').forEach(t => {{
    const fn = t.getAttribute('onclick') || '';
    t.classList.toggle('active', fn.includes("megaShow('" + id + "')"));
  }});

  // Update URL hash (without triggering hashchange)
  var hashMap = {{'vg':'visao-geral','cs':'categorias','pr':'produtoras','conc':'concentracao','cl':'curtas-longas','div':'diversidade','sp':'soft-power'}};
  history.replaceState(null, '', '#' + (hashMap[id] || id));

  // Close mobile sidebar if open
  var sb = document.getElementById('mega-sidebar');
  if(sb && sb.classList.contains('mobile-open')) toggleMobileSidebar();

  // Scroll content to top on section change
  var content = document.getElementById('mega-content');
  if(content) content.scrollTop = 0;
  _resizeVisibleCharts('mega-section-' + section);
}}

function vgShow(subname) {{
  document.querySelectorAll('#mega-section-cmp .cmp-tab-panel').forEach(p => p.style.display = 'none');
  const panel = document.getElementById('cmp-panel-'+subname);
  if(panel) panel.style.display = 'block';
  // Update subtab active state
  document.querySelectorAll('.vg-subtab').forEach(t => {{
    t.classList.toggle('active', (t.getAttribute('onclick') || '').includes("'"+subname+"'"));
  }});
  vgInitCharts(subname);
  // D3 map: ensure it renders even if fetch finished while panel was hidden
  if(subname === 'ret-intl') {{
    setTimeout(function() {{
      if(typeof window.intlMapEnsure === 'function') window.intlMapEnsure();
    }}, 60);
  }}
  _resizeVisibleCharts('cmp-panel-' + subname);
}}

// Lazy Plotly init — dark theme matching panel
// Use _vg_initialized (NOT _initialized) to avoid conflict with the comparativo
// scripts which declare  var _initialized = {{}}  — redeclaring a const causes
// a SyntaxError that silently kills the entire comparativo script block.
var _vg_initialized = {{}};

var _DARK_LAYOUT = {{
  paper_bgcolor: 'rgba(0,0,0,0)',
  plot_bgcolor:  'rgba(14,16,24,0.6)',
  font: {{color:'#e8eaf2', family:'DM Mono, monospace', size:11}}
}};
var _DARK_AXIS = {{gridcolor:'#1e2035',linecolor:'#2a2d42',zerolinecolor:'#2a2d42'}};

function _applyDark(div, data, layout) {{
  var _dl = Object.assign({{}}, layout, _DARK_LAYOUT);
  if(_dl.xaxis)  _dl.xaxis  = Object.assign({{}}, _dl.xaxis,  _DARK_AXIS);
  if(_dl.yaxis)  _dl.yaxis  = Object.assign({{}}, _dl.yaxis,  _DARK_AXIS);
  if(_dl.xaxis2) _dl.xaxis2 = Object.assign({{}}, _dl.xaxis2, {{gridcolor:'#1e2035',linecolor:'#2a2d42'}});
  if(_dl.yaxis2) _dl.yaxis2 = Object.assign({{}}, _dl.yaxis2, {{gridcolor:'#1e2035',linecolor:'#2a2d42'}});
  if(_dl.legend) _dl.legend = Object.assign({{}}, _dl.legend, {{bgcolor:'rgba(0,0,0,0)',font:{{color:'#e8eaf2'}}}});
  Plotly.newPlot(div, data, _dl, {{responsive:true,displaylogo:false}});
  if(typeof _originalData !== 'undefined')
    _originalData[div.id] = JSON.parse(JSON.stringify(data));
}}

function vgInitCharts(subname) {{
  if(_vg_initialized[subname]) return;
  _vg_initialized[subname] = true;
  var tabEl = document.getElementById('cmp-tab-'+subname);
  if(!tabEl) return;
  tabEl.querySelectorAll('div[id]').forEach(function(div) {{
    var key = '__fig_' + div.id;
    if(window[key]) {{
      _applyDark(div, window[key].data, window[key].layout);
    }}
  }});
}}

// Generic Plotly init for panels outside #mega-section-cmp
// (e.g. "Por Categoria" in CS section, "Por Cluster" in PR section)
var _figs_initialized = {{}};
function _initFigsInContainer(containerId) {{
  if(_figs_initialized[containerId]) return;
  _figs_initialized[containerId] = true;
  var el = document.getElementById(containerId);
  if(!el) return;
  el.querySelectorAll('div[id]').forEach(function(div) {{
    var key = '__fig_' + div.id;
    if(window[key]) {{
      _applyDark(div, window[key].data, window[key].layout);
    }}
  }});
}}

// Re-apply dark theme to inline (non-lazy) charts in visão-geral
// These charts are rendered at parse time with light theme
function vgFixInlineCharts() {{
  var tabEl = document.getElementById('cmp-tab-visao-geral');
  if(!tabEl) return;
  tabEl.querySelectorAll('.plotly-graph-div').forEach(function(div) {{
    if(div._fullLayout) {{  // already rendered by inline Plotly.newPlot
      Plotly.relayout(div, {{
        'paper_bgcolor': 'rgba(0,0,0,0)',
        'plot_bgcolor':  'rgba(14,16,24,0.6)',
        'font.color':    '#e8eaf2',
        'font.family':   'DM Mono, monospace'
      }});
      // go.Table: relayout não afeta cells.fill.color (propriedade de dados).
      // Restyle para escurecer células e tornar o texto legível no tema escuro.
      var trace = div.data && div.data[0];
      if(trace && trace.type === 'table') {{
        var nCols = trace.cells.values.length;
        var darkFills = [];
        var origFills = trace.cells.fill ? trace.cells.fill.color : null;
        for(var i = 0; i < nCols; i++) {{
          var nRows = trace.cells.values[i].length;
          if(i === 0 && origFills && origFills[0]) {{
            // Mantém cores dos grupos mas aumenta opacidade para visibilidade
            darkFills.push(origFills[0].map(function(c) {{
              return c.replace(/rgba\((\d+),(\d+),(\d+),[^)]+\)/, 'rgba($1,$2,$3,0.35)');
            }}));
          }} else {{
            var col = [];
            for(var r = 0; r < nRows; r++) col.push('#141620');
            darkFills.push(col);
          }}
        }}
        Plotly.restyle(div, {{'cells.fill.color': [darkFills]}});
      }}
    }}
  }});
}}

// ── Concentracao tab switcher ──
function concShow(id) {{
  document.querySelectorAll('.conc-tab').forEach(t => {{
    const fn = t.getAttribute('onclick') || '';
    t.classList.toggle('active', fn.includes("'"+id+"'"));
  }});
  ['conc-t1','conc-t2','conc-t3'].forEach(i => {{
    const el = document.getElementById(i);
    if(el) el.style.display = (i === id) ? 'flex' : 'none';
  }});
  window.dispatchEvent(new Event('resize'));
  _resizeVisibleCharts(id);
}}

// ── Helper: show a custom PR extra panel (ticket / concentração) ──
function _prShowExtraPanel(panelId, btnId, initFn) {{
  // Deactivate all standard PR tab panels and buttons
  if(typeof prSwitchTab === 'function') prSwitchTab(-1);
  // Deactivate other extra buttons
  ['pr-ticket-btn','pr-conc-btn'].forEach(function(id) {{
    var b = document.getElementById(id);
    if(b) b.classList.remove('active');
  }});
  // Hide other extra panels
  ['pr-ticket-panel','pr-conc-panel'].forEach(function(id) {{
    var p = document.getElementById(id);
    if(p) {{ p.classList.remove('active'); p.style.display = 'none'; }}
  }});
  // Show requested panel
  var panel = document.getElementById(panelId);
  if(panel) {{
    panel.classList.add('active');
    panel.style.display = 'flex';
    panel.style.flexDirection = 'column';
    panel.style.overflowY = 'auto';
    panel.style.flex = '1';
    panel.style.minHeight = '0';
  }}
  var btn = document.getElementById(btnId);
  if(btn) btn.classList.add('active');
  if(typeof initFn === 'function') setTimeout(initFn, 80);
  _resizeVisibleCharts(panelId);
}}

// ── Ticket tab inside Produtoras ──
function prShowTicket() {{
  _prShowExtraPanel('pr-ticket-panel', 'pr-ticket-btn', function() {{
    _initFigsInContainer('pr-ticket-panel');
    var ticketDiv = document.getElementById('cmp-tab-ticket');
    if(ticketDiv) {{
      ticketDiv.querySelectorAll('div[id]').forEach(function(div) {{
        var key = '__fig_' + div.id;
        if(window[key]) _applyDark(div, window[key].data, window[key].layout);
      }});
    }}
  }});
}}

// ── Concentração tab inside Produtoras ──
function prShowConc() {{
  _prShowExtraPanel('pr-conc-panel', 'pr-conc-btn', function() {{
    _initFigsInContainer('pr-conc-panel');
    // Initialize the inner concentration tabs explicitly.
    concShow('conc-t1');
  }});
}}

// ── Sidebar toggle ──
function toggleSidebar() {{
  document.body.classList.toggle('sidebar-collapsed');
  // Trigger Plotly resize after animation
  setTimeout(function() {{ window.dispatchEvent(new Event('resize')); }}, 300);
}}

// ── Methodology bar toggle ──
function toggleMeto() {{
  var bar = document.getElementById('meto-bar');
  if(bar) bar.classList.toggle('collapsed');
}}

// ── Mobile sidebar ──
function toggleMobileSidebar() {{
  var sb = document.getElementById('mega-sidebar');
  var ov = document.querySelector('.mobile-overlay');
  if(!sb) return;
  var isOpen = sb.classList.contains('mobile-open');
  sb.classList.toggle('mobile-open');
  if(ov) {{
    ov.classList.toggle('active', !isOpen);
    ov.style.opacity = isOpen ? '0' : '1';
    ov.style.pointerEvents = isOpen ? 'none' : 'auto';
  }}
}}

// ── Scroll-to-top visibility ──
(function() {{
  var content = document.getElementById('mega-content');
  var btn = document.getElementById('scroll-top');
  if(content && btn) {{
    content.addEventListener('scroll', function() {{
      btn.classList.toggle('visible', content.scrollTop > 300);
    }});
  }}
}})();

// ── Keyboard navigation ──
document.addEventListener('keydown', function(e) {{
  // Don't intercept if user is typing in an input/select
  if(e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;
  if(e.key === '1') {{ megaShow('vg'); e.preventDefault(); }}
  if(e.key === '2') {{ megaShow('cs'); e.preventDefault(); }}
  if(e.key === '3') {{ megaShow('pr'); e.preventDefault(); }}
  if(e.key === '4') {{ megaShow('conc'); e.preventDefault(); }}
  if(e.key === '5') {{ megaShow('cl'); e.preventDefault(); }}
  if(e.key === '6') {{ megaShow('div'); e.preventDefault(); }}
  if(e.key === '7') {{ megaShow('sp'); e.preventDefault(); }}
  if(e.key === '[') {{ toggleSidebar(); e.preventDefault(); }}
}});

// ── URL hash navigation ──
function handleHash() {{
  var h = window.location.hash.replace('#','');
  if(h === 'categorias' || h === 'cs') megaShow('cs');
  else if(h === 'produtoras' || h === 'pr') megaShow('pr');
  else if(h === 'concentracao') megaShow('conc');
  else if(h === 'curtas-longas' || h === 'cl') megaShow('cl');
  else if(h === 'diversidade' || h === 'div') megaShow('div');
  else if(h === 'soft-power' || h === 'sp') megaShow('sp');
  else if(h === 'soft-power-citacao' || h === 'citacao-academica') {{
    megaShow('sp');
    setTimeout(function() {{
      if(typeof window.spShow === 'function') window.spShow('citacao');
    }}, 160);
  }}
  else if(h === 'retorno-domestico' || h === 'financeiro') {{
    megaShow('vg');
    setTimeout(function() {{ vgShow('financeiro'); }}, 120);
  }}
  else if(h === 'ret-intl' || h === 'retorno-internacional') {{
    megaShow('vg');
    setTimeout(function() {{ vgShow('ret-intl'); }}, 120);
  }}
  else megaShow('vg');
}}

// Initialize on load
document.addEventListener('DOMContentLoaded', function() {{
  // Route from hash or default to vg
  if(window.location.hash) handleHash();
  else megaShow('vg');
  // Fix inline charts that rendered during parse with light theme
  setTimeout(vgFixInlineCharts, 150);
}});
window.addEventListener('hashchange', handleHash);
</script>

<!-- ── COMPARATIVO PLOTLY SCRIPTS ── -->
{cmp_scripts_prefixed}
{intl_map_fallback_script}

<!-- ── CONCENTRACAO DATA ── -->
<script>
{conc_data_script}
</script>

<!-- ── CONCENTRACAO LOGIC ── -->
<script>
{conc_logic_script}
</script>

<!-- ── CRITERIO SELECAO SCRIPT ── -->
<script>
{cs_script}
</script>

<!-- ── PRODUTORAS DATA ── -->
<script id="pr-prod-data">
{pr_data_script}
</script>

<!-- ── PRODUTORAS LOGIC ── -->
<script>
{pr_logic_script}
</script>

<!-- ── DIVERSIDADE RUNTIME CHARTS ── -->
{_div_render_script}

<!-- ── SOFT POWER RUNTIME CHARTS ── -->
{_sp_render_script}

</body>
</html>'''

with open(OUT, 'w', encoding='utf-8') as f:
    f.write(mega_html)

size_kb = os.path.getsize(OUT) / 1024
print(f"\nDone! Output: {OUT}")
print(f"File size: {size_kb:.1f} KB")
print(f"Lines: {mega_html.count(chr(10))}")
