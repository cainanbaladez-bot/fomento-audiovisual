"""
shared_styles.py — CSS centralizado para os painéis de dados.

Para alterar cores, fontes, componentes ou layout:
  1. Edite as constantes neste arquivo
  2. Re-execute os scripts que geram HTML (03–07)

Hierarquia:
  THEME_VARS  →  tokens de cor, fonte e spacing
  BASE_CSS    →  reset, body, scrollbar
  TABLE / KPI / CARD / TAB / SEARCH / MODAL / BADGE  →  componentes
  MEGA_*      →  layout do painel integrado (header, sidebar, content)
  *_SECTION   →  CSS específico de cada seção dentro do mega painel
"""

# ── CDN links ──────────────────────────────────────────────────────────────
GOOGLE_FONTS = (
    '<link rel="preconnect" href="https://fonts.googleapis.com">'
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
    '<link href="https://fonts.googleapis.com/css2?family='
    'Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">'
)
PLOTLY_CDN = '<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>'
D3_CDN = '<script src="https://d3js.org/d3.v7.min.js"></script>'
TOPOJSON_CDN = '<script src="https://cdn.jsdelivr.net/npm/topojson@3/dist/topojson.min.js"></script>'

# ── Plotly chart defaults (Python-side) ───────────────────────────────────
PLOTLY_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(14,16,24,0.6)',
    font=dict(color='#e2e8f0', family='Inter, system-ui, sans-serif', size=11),
    margin=dict(l=50, r=20, t=80, b=120),
)
PLOTLY_AXIS = dict(gridcolor='#282d42', linecolor='#343a54', zerolinecolor='#343a54')

# ═══════════════════════════════════════════════════════════════════════════
# CSS — THEME TOKENS
# ═══════════════════════════════════════════════════════════════════════════
THEME_VARS = """:root{
  /* ── Surfaces ── */
  --bg:#0b0d14;
  --surface:#14171f;
  --surface2:#1a1e2c;
  --surface3:#212638;
  --border:#282d42;
  --border-light:#343a54;

  /* ── Accent palette ── */
  --accent:#6c7bf7;
  --accent-dim:rgba(108,123,247,.12);
  --gold:#fbbf24;
  --coral:#f87171;
  --purple:#a78bfa;
  --cyan:#38bdf8;
  --green:#34d399;
  --muted-blue:#5fd1ff;

  /* ── Text ── */
  --text:#e2e8f0;
  --text2:#c1c9d9;
  --muted:#7b849a;
  --dim:#282d42;

  /* ── Typography ── */
  --font-head:'Inter',system-ui,sans-serif;
  --font-mono:'Inter',system-ui,sans-serif;
  --font-ui:'Inter',system-ui,sans-serif;

  /* ── Layout ── */
  --sidebar-w:210px;
  --sidebar-collapsed-w:0px;
  --transition-speed:.25s;
}"""

# ═══════════════════════════════════════════════════════════════════════════
# CSS — BASE RESET
# ═══════════════════════════════════════════════════════════════════════════
BASE_CSS = """*{box-sizing:border-box;margin:0;padding:0}
html{-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale}
body{background:var(--bg);color:var(--text);font-family:var(--font-ui);font-size:13px;min-height:100vh;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:var(--border-light)}
a{color:var(--accent);text-underline-offset:3px}"""

# ═══════════════════════════════════════════════════════════════════════════
# CSS — SHARED COMPONENTS
# ═══════════════════════════════════════════════════════════════════════════

# ── Tables ──
TABLE_CSS = """table{width:100%;border-collapse:collapse}
thead th{position:sticky;top:0;z-index:1;background:var(--surface);
  padding:6px 10px;text-align:left;font-size:9px;letter-spacing:.08em;
  text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);
  white-space:nowrap;cursor:pointer;user-select:none}
thead th:hover{color:var(--text)}
thead th.sa::after{content:' ↑';color:var(--accent)}
thead th.sd::after{content:' ↓';color:var(--accent)}
thead th.r{text-align:right}
tbody tr{border-bottom:1px solid rgba(255,255,255,.035);transition:background .06s}
tbody tr:hover{background:var(--surface2)}
td{padding:6px 10px;vertical-align:middle;white-space:nowrap;font-size:11px;color:var(--text)}
td.r{text-align:right;font-variant-numeric:tabular-nums}
td.dim{color:var(--muted)}"""

# ── KPIs ──
KPI_CSS = """.kpi-bar{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:18px}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:12px 14px;position:relative;overflow:hidden}
.kpi::after{content:'';position:absolute;inset:0;background:radial-gradient(circle at 80% 20%,rgba(108,123,247,.03),transparent 70%);pointer-events:none}
.kpi.warn{border-left:3px solid var(--coral)}
.kpi.ok{border-left:3px solid var(--green)}
.kpi.mid{border-left:3px solid var(--gold)}
.kpi-l{font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:5px}
.kpi-v{font-family:var(--font-head);font-size:22px;font-weight:700;font-style:italic;color:var(--accent);line-height:1}
.kpi-u{font-size:9px;color:var(--muted);margin-left:3px}
.kpi-sub{font-size:9px;color:var(--muted);margin-top:4px}
.kpi-label{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;margin-bottom:8px}
.kpi-val{font-family:var(--font-head);font-size:24px;font-style:italic;color:var(--accent);line-height:1}"""

# ── Cards ──
CARD_CSS = """.card{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:14px 16px;margin-bottom:16px}
.card-t{font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:8px}
.chart-wrap{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px;margin-bottom:16px}
.chart-header{display:flex;align-items:center;gap:16px;margin-bottom:14px;flex-wrap:wrap}
.chart-title{font-family:var(--font-head);font-size:15px;font-weight:400;font-style:italic}"""

# ── Tabs ──
TAB_CSS = """.tabs{display:flex;padding:0 24px;border-bottom:1px solid var(--border);background:var(--surface)}
.tab{padding:12px 20px;font-size:11px;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;
  transition:all .2s;letter-spacing:.06em;text-transform:uppercase;font-family:var(--font-mono);
  background:none;border-top:none;border-left:none;border-right:none}
.tab.active{color:var(--accent);border-bottom-color:var(--accent)}
.tab:hover:not(.active){color:var(--text)}
.tab-bar{display:flex;gap:0;padding:8px 18px 0;flex-shrink:0;border-bottom:1px solid var(--border)}
.tab-btn{padding:6px 14px;font-family:var(--font-mono);font-size:10px;letter-spacing:.1em;text-transform:uppercase;
  background:none;color:var(--muted);border:none;border-bottom:2px solid transparent;cursor:pointer;
  transition:color .15s,border-color .15s;margin-bottom:-1px;white-space:nowrap}
.tab-btn:hover{color:var(--text)}
.tab-btn.active{color:var(--accent);border-bottom-color:var(--accent)}
.tab-panel{display:none;flex:1;flex-direction:column;min-height:0;overflow:hidden}
.tab-panel.active{display:flex}
.panel{display:none;padding:24px 28px;animation:fadeIn .2s ease}
.panel.active{display:block}
@keyframes fadeIn{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}"""

# ── Search ──
SEARCH_CSS = """.search-wrap{display:flex;align-items:center;gap:10px;position:relative;flex-shrink:0}
.search-input{flex:1;max-width:420px;padding:7px 12px;border:1px solid var(--border);border-radius:6px;
  font-family:var(--font-mono);font-size:11px;background:var(--surface2);color:var(--text);
  transition:border-color .15s,box-shadow .15s}
.search-input:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 2px rgba(108,123,247,.15)}
.search-input::placeholder{color:var(--muted)}
.search-count{font-size:10px;color:var(--muted)}
.search-in{background:var(--surface2);color:var(--text);border:1px solid var(--border);
  font-family:var(--font-mono);font-size:10px;padding:4px 28px 4px 8px;
  border-radius:4px;outline:none;width:180px;transition:border-color .15s,width .15s}
.search-in:focus{border-color:var(--accent);width:240px}
.search-in::placeholder{color:var(--muted)}
.search-clr{position:absolute;right:6px;background:none;border:none;color:var(--muted);
  cursor:pointer;font-size:13px;padding:0;display:none}
.search-clr.show{display:block}
.search-clr:hover{color:var(--text)}"""

# ── Modals ──
MODAL_CSS = """.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.8);z-index:200;align-items:center;justify-content:center}
.modal-overlay.open{display:flex}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:14px;
  width:min(800px,94vw);max-height:86vh;overflow:hidden;display:flex;flex-direction:column}
.modal-head{padding:20px 24px;border-bottom:1px solid var(--border);display:flex;align-items:flex-start;gap:16px;
  background:linear-gradient(135deg,var(--surface) 0%,var(--surface2) 100%)}
.modal-head h2{font-family:var(--font-head);font-size:18px;font-style:italic;margin-bottom:6px}
.modal-meta{display:flex;gap:14px;flex-wrap:wrap}
.modal-kpi{font-size:11px;color:var(--muted)}
.modal-kpi span{color:var(--accent);font-weight:500}
.modal-close{width:30px;height:30px;border-radius:50%;border:1px solid var(--border);background:transparent;
  color:var(--muted);cursor:pointer;font-size:14px;display:flex;align-items:center;justify-content:center;
  flex-shrink:0;margin-left:auto;transition:background .15s}
.modal-close:hover{background:var(--surface2)}
.modal-body{padding:20px 24px;overflow-y:auto;flex:1}"""

# ── Badges ──
BADGE_CSS = """.badge{font-size:8px;letter-spacing:.08em;text-transform:uppercase;padding:2px 7px;border-radius:4px;color:#fff;white-space:nowrap}
.tag{display:inline-block;padding:2px 8px;border-radius:4px;font-size:9px;font-weight:600;letter-spacing:.04em}"""

# ── Controls ──
CONTROLS_CSS = """.ctrl-group{display:flex;gap:6px;margin-left:auto}
.ctrl-btn{padding:5px 12px;border-radius:6px;border:1px solid var(--border);background:transparent;
  color:var(--muted);cursor:pointer;font-family:var(--font-mono);font-size:10px;transition:all .2s}
.ctrl-btn.active,.ctrl-btn:hover{background:var(--accent);color:#000;border-color:var(--accent);font-weight:600}
.tog-btn{padding:5px 14px;font-size:11px;cursor:pointer;border-radius:4px;
  border:1px solid var(--border);background:var(--surface2);color:var(--muted);transition:all .15s}
.tog-btn.active{background:var(--accent);color:#000;border-color:var(--accent)}
.tog-btn:hover:not(.active){background:var(--surface);color:var(--text)}
.toggle-bar{display:flex;gap:6px;margin-bottom:10px}
select{background:var(--surface2);color:var(--text);border:1px solid var(--border);
  font-family:var(--font-mono);font-size:10px;padding:5px 20px 5px 8px;border-radius:4px;
  cursor:pointer;outline:none;appearance:none;-webkit-appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='8' height='5'%3E%3Cpath d='M0 0l4 5 4-5z' fill='%237b849a'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 6px center}
.filter-bar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:14px}
.filter-bar select,.filter-bar input{background:var(--surface2);border:1px solid var(--border);
  color:var(--text);padding:7px 12px;border-radius:6px;font-family:var(--font-mono);font-size:11px}
.filter-bar input{width:200px}"""

# ── Grids ──
GRID_CSS = """.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px}
.grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:16px}
.comp-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
@media(max-width:768px){.grid-2,.grid2,.grid3,.comp-grid{grid-template-columns:1fr}}"""

# ── Alerts/Info boxes ──
ALERT_CSS = """.alert{background:rgba(248,113,113,.08);border:1px solid rgba(248,113,113,.25);border-radius:6px;
  padding:10px 14px;margin-bottom:14px;font-size:10px;color:var(--text);line-height:1.7}
.alert b{color:var(--coral)}
.info{background:rgba(56,189,248,.06);border:1px solid rgba(56,189,248,.2);border-radius:6px;
  padding:10px 14px;margin-bottom:14px;font-size:10px;color:var(--text);line-height:1.7}
.info b{color:var(--cyan)}"""

# ── Inline bars ──
BAR_CSS = """.bar-inline{display:flex;align-items:center;gap:8px}
.bar-bg{flex:1;height:3px;background:var(--surface3);border-radius:2px;max-width:80px}
.bar-fill{height:3px;border-radius:2px}
.mod-bars{display:flex;flex-direction:column;gap:10px;margin-bottom:16px}
.mod-bar-row{display:flex;align-items:center;gap:14px}
.mod-bar-label{min-width:220px;font-size:11px;color:var(--text)}
.mod-bar-track{flex:1;height:22px;background:var(--surface2);border-radius:4px;position:relative;overflow:hidden}
.mod-bar-fill{height:100%;border-radius:4px;display:flex;align-items:center;padding-left:8px;font-size:10px;font-weight:500;color:#000}
.mod-bar-val{min-width:70px;font-size:11px;color:var(--muted);text-align:right}"""

# ── Comparison cards ──
COMP_CSS = """.comp-card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px}
.comp-card h3{font-family:var(--font-head);font-size:14px;font-style:italic;margin-bottom:12px}
.comp-row{display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid var(--dim);font-size:11px}
.comp-row:last-child{border-bottom:none}
.comp-row .label{color:var(--muted)}
.comp-row .val{font-weight:500}"""

# ── Legenda ──
LEGENDA_CSS = """.legenda{display:flex;gap:18px;flex-wrap:wrap;margin-bottom:16px;
  background:var(--surface);padding:12px 18px;border-radius:8px;border:1px solid var(--border)}
.legenda-item{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--text)}
.legenda-item span{color:var(--muted)}
.legenda-dot{width:12px;height:12px;border-radius:50%;flex-shrink:0}"""

# ── Section title ──
SECTION_TITLE_CSS = """.section-title{font-family:var(--font-head);font-size:13px;font-weight:400;font-style:italic;
  letter-spacing:-.01em;color:var(--accent);margin-bottom:12px;padding-bottom:6px;border-bottom:1px solid var(--border)}"""

# ── Methodology cards ──
METO_CSS = """.meto-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.meto-card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px}
.meto-card h3{font-family:var(--font-head);font-size:13px;font-style:italic;margin-bottom:10px;color:var(--accent)}
.meto-row{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid var(--border);font-size:11px}
.meto-row:last-child{border-bottom:none}
.meto-row span:last-child{color:var(--accent);font-weight:500}
.meto-note{grid-column:1/-1;background:var(--surface2);border:1px solid var(--dim);border-radius:8px;
  padding:14px;font-size:11px;line-height:1.8;color:var(--muted)}
.meto-note strong{color:var(--text)}
@media(max-width:768px){.meto-grid{grid-template-columns:1fr}}"""

# ── Cat pills ──
PILLS_CSS = """.cat-pills{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap}
.cat-pill{padding:7px 16px;border-radius:20px;border:1px solid var(--border);background:var(--surface);
  cursor:pointer;font-family:var(--font-mono);font-size:11px;color:var(--muted);transition:all .2s}
.cat-pill.active{border-color:var(--accent);color:var(--accent);background:var(--accent-dim)}
.cat-pill:hover:not(.active){color:var(--text);border-color:var(--muted)}"""

# ── Tooltip ──
TOOLTIP_CSS = """#tooltip{position:fixed;display:none;background:var(--surface2);border:1px solid var(--accent);
  border-radius:8px;padding:12px 14px;font-size:11px;line-height:1.7;pointer-events:none;z-index:999;
  max-width:310px;box-shadow:0 8px 32px rgba(0,0,0,.6)}
#tooltip strong{font-family:var(--font-head);font-size:14px;font-style:italic;color:var(--accent);display:block;margin-bottom:4px}"""


# ═══════════════════════════════════════════════════════════════════════════
# All components combined (for standalone panels)
# ═══════════════════════════════════════════════════════════════════════════
ALL_COMPONENTS = '\n'.join([
    TABLE_CSS, KPI_CSS, CARD_CSS, TAB_CSS, SEARCH_CSS, MODAL_CSS, BADGE_CSS,
    CONTROLS_CSS, GRID_CSS, ALERT_CSS, BAR_CSS, COMP_CSS, LEGENDA_CSS,
    SECTION_TITLE_CSS, METO_CSS, PILLS_CSS, TOOLTIP_CSS,
])

def standalone_css(*extra):
    """Full CSS string for a standalone dark-themed panel."""
    parts = [THEME_VARS, BASE_CSS, ALL_COMPONENTS] + list(extra)
    return '\n'.join(parts)


# ═══════════════════════════════════════════════════════════════════════════
# CSS — MEGA PANEL LAYOUT (used only by script 07)
# ═══════════════════════════════════════════════════════════════════════════

MEGA_HEADER_CSS = """/* ── Mega header ── */
.mega-header{
  padding:0;border-bottom:1px solid var(--border);
  background:linear-gradient(180deg,#0f1119 0%,var(--bg) 100%);
  display:flex;flex-direction:column;flex-shrink:0;
}
.mega-header-top{display:flex;align-items:center;gap:20px;padding:16px 32px 12px}
.mega-header-icon{
  width:44px;height:44px;border-radius:12px;background:var(--accent);
  display:flex;align-items:center;justify-content:center;flex-shrink:0;
  box-shadow:0 0 20px rgba(108,123,247,.15);
}
.mega-header-icon svg{width:20px;height:20px}
.mega-header-titles{flex:1;min-width:0}
.mega-header-titles h1{
  font-family:var(--font-head);font-size:21px;font-weight:400;
  letter-spacing:-.3px;line-height:1.15;overflow-wrap:anywhere;
}
.mega-header-titles p{font-size:11px;color:var(--muted);margin-top:3px;letter-spacing:.02em;overflow-wrap:anywhere}
.mega-header-meta{text-align:right;font-size:10px;color:var(--muted);line-height:1.7}
.audit-link{
  display:inline-flex;align-items:center;gap:6px;flex-shrink:0;
  padding:6px 14px;border-radius:20px;border:1px solid rgba(108,123,247,.35);
  background:rgba(108,123,247,.1);color:var(--accent);
  font-size:11px;font-weight:600;letter-spacing:.04em;text-decoration:none;
  white-space:nowrap;transition:all .2s;
}
.audit-link:hover{background:rgba(108,123,247,.2);border-color:var(--accent);color:#fff;text-decoration:none}
@media(max-width:900px){.audit-link{display:none}}"""

MEGA_METO_CSS = """/* ── Methodology ribbon ── */
.mega-meto-bar{
  display:flex;align-items:center;gap:0;padding:0 32px;
  border-top:1px solid var(--border);background:rgba(255,255,255,.02);
  overflow:hidden;max-height:60px;transition:max-height var(--transition-speed) ease;
  overflow-x:auto;overflow-y:hidden;scrollbar-width:thin;
}
.mega-meto-bar.collapsed{max-height:0;border-top-color:transparent}
.meto-toggle{
  position:absolute;right:32px;top:0;height:100%;
  display:flex;align-items:center;background:none;border:none;
  color:var(--muted);cursor:pointer;font-family:var(--font-mono);font-size:9px;
  letter-spacing:.08em;text-transform:uppercase;gap:4px;padding:0 4px;transition:color .15s;
}
.meto-toggle:hover{color:var(--text)}
.meto-toggle svg{width:10px;height:10px;transition:transform var(--transition-speed)}
.mega-meto-bar.collapsed ~ .meto-toggle svg{transform:rotate(180deg)}
.mega-meto-wrap{position:relative;border-top:1px solid var(--border)}
.mega-meto-wrap .meto-toggle{top:0;height:32px}
.meto-item{
  display:flex;align-items:baseline;gap:7px;padding:8px 18px 8px 0;
  border-right:1px solid var(--border);margin-right:18px;white-space:nowrap;font-size:10px;flex-shrink:0;
}
.meto-item:last-child{border-right:none;margin-right:0}
.meto-label{font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--accent);font-weight:600}
.meto-def{color:var(--muted);line-height:1.5}
.meto-def strong{color:var(--text)}"""

MEGA_SIDEBAR_CSS = """/* ── Sidebar navigation ── */
body{display:flex;flex-direction:column;height:100vh;overflow:hidden}
.mega-layout{display:flex;flex:1;overflow:hidden;min-height:0;width:100%}

.mega-sidebar{
  width:var(--sidebar-w);min-width:var(--sidebar-w);
  background:var(--surface);border-right:1px solid var(--border);
  display:flex;flex-direction:column;overflow-y:auto;padding:8px 0;
  transition:width var(--transition-speed) ease,min-width var(--transition-speed) ease,
             opacity var(--transition-speed) ease,transform var(--transition-speed) ease;
}
body.sidebar-collapsed .mega-sidebar{width:0;min-width:0;opacity:0;overflow:hidden;padding:0;border-right:none}
.mega-sidebar-label{padding:14px 18px 6px;font-size:9px;color:var(--muted);letter-spacing:.1em;text-transform:uppercase}

.mega-tab{
  display:flex;align-items:center;gap:10px;width:100%;text-align:left;
  padding:10px 18px;font-size:10.5px;color:var(--muted);cursor:pointer;
  border-left:2px solid transparent;transition:all .15s;
  letter-spacing:.04em;text-transform:uppercase;font-family:var(--font-mono);
  white-space:normal;background:none;border-top:none;border-bottom:none;border-right:none;line-height:1.4;
}
.mega-tab .tab-icon{width:16px;height:16px;flex-shrink:0;opacity:.5;transition:opacity .15s}
.mega-tab.active .tab-icon{opacity:1}
.mega-tab .tab-key{
  margin-left:auto;font-size:9px;color:var(--dim);
  border:1px solid var(--border);border-radius:3px;padding:1px 5px;
  font-family:var(--font-mono);line-height:1.4;
}
.mega-tab.active{color:var(--accent);border-left-color:var(--accent);background:var(--accent-dim)}
.mega-tab:hover:not(.active){color:var(--text);background:rgba(255,255,255,.03)}
.mega-tab:hover .tab-key{color:var(--muted)}

/* Sidebar toggle (desktop) */
.sidebar-toggle{
  position:fixed;left:var(--sidebar-w);top:50%;transform:translateY(-50%);
  z-index:50;width:18px;height:38px;background:var(--surface);
  border:1px solid var(--border);border-left:none;border-radius:0 6px 6px 0;
  cursor:pointer;display:flex;align-items:center;justify-content:center;
  color:var(--muted);transition:left var(--transition-speed) ease,color .15s,background .15s;
}
.sidebar-toggle:hover{color:var(--accent);background:var(--surface2)}
.sidebar-toggle svg{width:12px;height:12px;transition:transform var(--transition-speed)}
body.sidebar-collapsed .sidebar-toggle{left:0}
body.sidebar-collapsed .sidebar-toggle svg{transform:rotate(180deg)}"""

MEGA_CONTENT_CSS = """/* ── Main content area ── */
.mega-content{flex:1;overflow-y:auto;position:relative;min-width:0;min-height:0;
  overflow-x:hidden;overscroll-behavior:contain;-webkit-overflow-scrolling:touch}
.mega-panel{display:none;min-height:100%;min-width:0}
.mega-panel.active{display:block}

/* Scroll-to-top */
.scroll-top{
  position:fixed;bottom:24px;right:24px;z-index:100;width:38px;height:38px;border-radius:50%;
  background:var(--surface2);border:1px solid var(--border);color:var(--muted);cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  opacity:0;pointer-events:none;transition:opacity .2s,background .15s,color .15s,transform .15s;
  box-shadow:0 4px 12px rgba(0,0,0,.3);
}
.scroll-top.visible{opacity:1;pointer-events:auto}
.scroll-top:hover{background:var(--accent);color:#000;transform:scale(1.08)}
.scroll-top svg{width:16px;height:16px}"""

# ── Visão Geral sub-navigation ──
VG_SUBNAV_CSS = """.vg-subnav{
  display:flex;padding:0 28px;background:var(--surface);
  border-bottom:1px solid var(--border);position:sticky;top:0;z-index:10;
  overflow-x:auto;overflow-y:hidden;scrollbar-width:thin;
}
.vg-subtab{
  padding:11px 18px;font-size:11px;color:var(--muted);cursor:pointer;
  border-bottom:2px solid transparent;transition:all .2s;
  letter-spacing:.06em;text-transform:uppercase;font-family:var(--font-mono);
  background:none;border-top:none;border-left:none;border-right:none;
  margin-bottom:-1px;white-space:nowrap;flex-shrink:0;
}
.vg-subtab.active{color:var(--accent);border-bottom-color:var(--accent)}
.vg-subtab:hover:not(.active){color:var(--text)}"""

# ═══════════════════════════════════════════════════════════════════════════
# CSS — SECTION OVERRIDES INSIDE MEGA PANEL
# ═══════════════════════════════════════════════════════════════════════════

# ── CMP (Visão Geral) — dark-theme overrides for comparativo content ──
CMP_SECTION_CSS = """/* ── Visão Geral section ── */
#mega-section-cmp{background:var(--bg)}
#mega-section-cmp .tab-content{display:block!important;background:var(--bg)!important;padding:24px 28px!important}
#mega-section-cmp .topbar{display:none!important}
#mega-section-cmp .tabs{display:none!important}
#mega-section-cmp .card{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important;border-radius:8px!important}
#mega-section-cmp .section-title{color:var(--accent)!important;border-bottom-color:var(--border)!important}
#mega-section-cmp .kpi-card{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important}
#mega-section-cmp .kpi-card .val{color:var(--accent)!important}
#mega-section-cmp .kpi-card .label{color:var(--muted)!important}
#mega-section-cmp .kpi-card .sub{color:var(--muted)!important}
#mega-section-cmp table thead th{color:var(--muted)!important;border-bottom-color:var(--border)!important;background:var(--bg)!important}
#mega-section-cmp table tbody td{border-bottom-color:var(--border)!important;color:var(--text)!important}
#mega-section-cmp table tbody tr:hover{background:var(--surface2)!important}
#mega-section-cmp h3,#mega-section-cmp h4{color:var(--text)!important}
#mega-section-cmp .insight-box,#mega-section-cmp .alert-box{background:var(--surface2)!important;border-color:var(--border)!important;color:var(--text)!important}
#mega-section-cmp select{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}
#mega-section-cmp input{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}
#mega-section-cmp .legenda{background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:8px}
#mega-section-cmp .legenda-item{color:var(--text)!important}
#mega-section-cmp .legenda-item span{color:var(--muted)!important}
#mega-section-cmp .chart-wrap{background:var(--surface)!important;border:1px solid var(--border)!important;border-radius:10px!important}
#mega-section-cmp .chart-title{color:var(--text)!important;font-family:var(--font-head)!important}
#mega-section-cmp .ctrl-btn{border-color:var(--border)!important;color:var(--muted)!important;background:transparent!important}
#mega-section-cmp .ctrl-btn.active,#mega-section-cmp .ctrl-btn:hover{background:var(--accent)!important;color:#000!important;border-color:var(--accent)!important;font-weight:600!important}
#mega-section-cmp .tog-btn{border-color:var(--border)!important;color:var(--muted)!important;background:var(--surface2)!important}
#mega-section-cmp .tog-btn.active{background:var(--accent)!important;color:#000!important;border-color:var(--accent)!important}
#mega-section-cmp .tog-btn:hover:not(.active){background:var(--surface)!important;color:var(--text)!important}
#mega-section-cmp .search-input{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}
#mega-section-cmp .search-input:focus{border-color:var(--accent)!important;box-shadow:0 0 0 2px rgba(108,123,247,.15)!important}
#mega-section-cmp .search-count{color:var(--muted)!important}
#mega-section-cmp .mod-bar-track{background:var(--surface2)!important}
#mega-section-cmp .mod-bar-fill{color:#000!important}
/* cluster cards */
#mega-section-cmp .cl-card{background:var(--surface)!important;border:1px solid var(--border)!important}
#mega-section-cmp .cl-card:hover{background:var(--surface2)!important}
#mega-section-cmp .cl-name{color:var(--text)!important}
#mega-section-cmp .cl-n{color:var(--text)!important}
#mega-section-cmp .cl-desc,#mega-section-cmp .cl-stat span,#mega-section-cmp .cl-top,#mega-section-cmp .cl-pct{color:var(--muted)!important}
#mega-section-cmp .cl-stat b{color:var(--text)!important}
#mega-section-cmp .cl-top b{color:var(--accent)!important}
#mega-section-cmp .ov-rank-head{border-bottom-color:var(--border)!important}
#mega-section-cmp .bc-nm{color:var(--text)!important}
#mega-section-cmp .bc-val{color:var(--text)!important}
#mega-section-cmp .bc-rank,#mega-section-cmp .bc-sub{color:var(--muted)!important}
#mega-section-cmp .bc-bar-wrap{background:var(--surface2)!important}
#mega-section-cmp .quad-wrap{background:var(--surface)!important;border:1px solid var(--border)!important}
#mega-section-cmp .filter-bar select,#mega-section-cmp .filter-bar input{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}
#mega-section-cmp .bar-bg{background:rgba(30,32,53,.8)!important}
#mega-section-cmp .comp-card{background:var(--surface)!important;border:1px solid var(--border)!important}
#mega-section-cmp .comp-row .label{color:var(--muted)!important}
#mega-section-cmp .comp-row .val{color:var(--text)!important}
#mega-section-cmp .quad-ctrl .lbl{color:var(--muted)!important}
#mega-section-cmp .obras-drawer{background:var(--surface)!important;border:1px solid var(--border)!important}
#mega-section-cmp .obras-drawer-title{color:var(--text)!important}
#mega-section-cmp .cl-panel{background:var(--surface2)!important;border:1px solid var(--border)!important}
#mega-section-cmp .cat-pill{background:var(--surface)!important;border-color:var(--border)!important;color:var(--muted)!important}
#mega-section-cmp .cat-pill.active{background:var(--accent-dim)!important;border-color:var(--accent)!important;color:var(--accent)!important}

/* Retorno Internacional: override inline light-theme styles */
#cmp-panel-ret-intl *{color:var(--text)}
#cmp-panel-ret-intl [style*="background:#fff"],
#cmp-panel-ret-intl [style*="background: #fff"]{background:var(--bg)!important;color:var(--text)!important}
#cmp-panel-ret-intl [style*="background:#f7f8fb"]{background:var(--surface)!important;border-color:var(--border)!important}
#cmp-panel-ret-intl [style*="background:#f0f4f8"]{background:var(--surface)!important;border-color:var(--border)!important}
#cmp-panel-ret-intl [style*="background:#eef0f4"]{background:var(--surface2)!important}
#cmp-panel-ret-intl [style*="background:rgba(255,255,255"]{background:var(--surface)!important;border-color:var(--border)!important}
#cmp-panel-ret-intl [style*="color:#222"]{color:var(--text)!important}
#cmp-panel-ret-intl [style*="color:#333"]{color:var(--text)!important}
#cmp-panel-ret-intl [style*="color:#444"]{color:var(--text)!important}
#cmp-panel-ret-intl [style*="color:#555"]{color:var(--muted)!important}
#cmp-panel-ret-intl [style*="color:#666"]{color:var(--muted)!important}
#cmp-panel-ret-intl [style*="color:#888"]{color:var(--muted)!important}
#cmp-panel-ret-intl [style*="border:1px solid #dde0e8"]{border-color:var(--border)!important}
#cmp-panel-ret-intl [style*="border:1px solid #ccd0da"]{border-color:var(--border)!important}
#cmp-panel-ret-intl [style*="border:1px solid #ddd"]{border-color:var(--border)!important}
#cmp-panel-ret-intl [style*="border:1px solid #5B6BB5"]{border-color:var(--accent)!important}
#cmp-panel-ret-intl #intl-tooltip{background:var(--surface2)!important;border-color:var(--accent)!important;color:var(--text)!important}"""

# ── CS (Critério de Seleção) — minimal overrides ──
CS_SECTION_CSS = """/* ── Critério de Seleção ── */
#mega-section-cs .header{display:none!important}
#mega-section-cs .tabs{background:var(--surface);border-bottom:1px solid var(--border);padding:0 20px}
#mega-section-cs .tab{font-size:10px}
#mega-section-cs .panel{padding:24px 28px}
#mega-section-cs #cs-tooltip{position:fixed;display:none;background:var(--surface2);border:1px solid var(--accent);
  border-radius:8px;padding:12px 14px;font-size:11px;line-height:1.7;pointer-events:none;z-index:9999;max-width:310px}"""

# ── PR (Produtoras) — overrides ──
PR_SECTION_CSS = """/* ── Produtoras ── */
#mega-section-pr{overflow:hidden!important;flex-direction:column;height:100%;min-height:0}
#mega-section-pr .hdr{display:none!important}
#mega-section-pr .tab-bar{background:var(--surface);border-bottom:1px solid var(--border);flex-shrink:0}
#mega-section-pr .tab-btn{color:var(--muted);background:none;border-color:transparent}
#mega-section-pr .tab-btn.active{color:var(--accent);border-bottom-color:var(--accent)!important}
#mega-section-pr .tab-panel{background:var(--bg)}
#mega-section-pr .tab-panel.main-tab.active{overflow-y:auto}

/* Por Cluster panel */
#pr-tab-clusters .card{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important;border-radius:8px!important}
#pr-tab-clusters .card div{color:var(--text)!important}
#pr-tab-clusters .search-input{background:var(--surface2)!important;color:var(--text)!important;border-color:var(--border)!important}
#pr-tab-clusters .search-count{color:var(--muted)!important}
#pr-tab-clusters table thead th{color:var(--muted)!important;background:var(--bg)!important}
#pr-tab-clusters table tbody td{color:var(--text)!important;border-bottom-color:var(--border)!important}

/* Ticket panel */
#pr-ticket-panel{background:var(--bg)!important;padding:0!important;overflow-y:auto!important;flex-direction:column!important}
#pr-ticket-panel .cmp-tab-panel{display:block!important}
#pr-ticket-panel .tab-content{display:block!important;background:var(--bg)!important;padding:14px 18px!important}
#pr-ticket-panel .card{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important}
#pr-ticket-panel .section-title{color:var(--accent)!important;border-bottom-color:var(--border)!important}
#pr-ticket-panel .kpi-card{background:var(--surface)!important;box-shadow:none!important;border:1px solid var(--border)!important;color:var(--text)!important}
#pr-ticket-panel .kpi-card .val{color:var(--accent)!important}
#pr-ticket-panel table thead th{color:var(--muted)!important;border-bottom-color:var(--border)!important;background:var(--bg)!important}
#pr-ticket-panel table tbody td{border-bottom-color:var(--border)!important;color:var(--text)!important}
#pr-ticket-panel table tbody tr:hover{background:var(--surface2)!important}
#pr-ticket-panel h3,#pr-ticket-panel h4{color:var(--text)!important}"""

# ── CONC (Concentração) — overrides ──
CONC_SECTION_CSS = """/* ── Concentração ── */
#conc-section .conc-subnav{display:flex;gap:0;border-bottom:1px solid var(--border);margin-bottom:20px}
#conc-section .conc-tab{padding:9px 16px;font-size:10px;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;
  letter-spacing:.06em;text-transform:uppercase;font-family:var(--font-mono);background:none;border-top:none;border-left:none;border-right:none;white-space:nowrap}
#conc-section .conc-tab.active{color:var(--accent);border-bottom-color:var(--accent)}
#conc-section .tab-panel{display:none;padding:0 4px;flex-direction:column}
#conc-section .tab-panel.active{display:flex}
#conc-section .scroll{overflow:visible!important;height:auto!important}

/* Standalone Concentração section */
#mega-section-conc,#mega-section-cl{background:var(--bg);color:var(--text)}
#mega-section-conc .card,#mega-section-cl .card{background:var(--surface)!important;border:1px solid var(--border)!important;color:var(--text)!important;box-shadow:none!important;min-width:0!important;overflow:hidden!important}
#mega-section-conc .card-t,#mega-section-cl .card-t{color:var(--accent)!important;line-height:1.35!important}
#mega-section-conc .kpi,#mega-section-cl .kpi{background:var(--surface)!important;border:1px solid var(--border)!important}
#mega-section-conc .kpi-v,#mega-section-cl .kpi-v{color:var(--text)!important}
#mega-section-conc .kpi-l,#mega-section-conc .kpi-sub,#mega-section-cl .kpi-l,#mega-section-cl .kpi-sub{color:var(--muted)!important}
#mega-section-conc .alert,#mega-section-cl .alert{background:var(--surface)!important;border-color:var(--border)!important;color:var(--text)!important}
#mega-section-conc .info,#mega-section-cl .info{background:var(--surface2)!important;color:var(--muted)!important;border-color:var(--border)!important}
#mega-section-conc table thead th,#mega-section-cl table thead th{color:var(--muted)!important;border-bottom-color:var(--border)!important;background:var(--bg)!important}
#mega-section-conc table tbody td,#mega-section-cl table tbody td{border-bottom-color:var(--border)!important;color:var(--text)!important}
#mega-section-conc .grid2,#mega-section-cl .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
#mega-section-conc .scroll,#mega-section-cl .scroll{overflow-y:auto}
#mega-section-conc .conc-tab,#mega-section-conc .conc-subnav{margin:0;padding:0}
#mega-section-conc .conc-tab.active{color:var(--accent)!important;border-bottom-color:var(--accent)!important}
#mega-section-conc{overflow-x:hidden!important}
#mega-section-conc .conc-subnav{
  display:flex!important;gap:8px!important;align-items:center!important;
  background:transparent!important;border-bottom:1px solid var(--border)!important;
  margin:4px 0 18px!important;padding:0 0 10px!important;overflow-x:auto!important;
}
#mega-section-conc .conc-tab{
  appearance:none!important;background:var(--surface)!important;border:1px solid var(--border)!important;
  border-radius:3px!important;color:var(--muted)!important;padding:8px 13px!important;
  font-family:var(--font-mono)!important;font-size:9px!important;letter-spacing:.08em!important;
  text-transform:uppercase!important;white-space:nowrap!important;cursor:pointer!important;
  border-bottom:1px solid var(--border)!important;
}
#mega-section-conc .conc-tab:hover{color:var(--text)!important;border-color:var(--muted)!important}
#mega-section-conc .conc-tab.active{color:var(--accent)!important;border-color:var(--accent)!important;background:rgba(108,123,247,.06)!important}
#mega-section-conc .tab-panel{min-width:0!important}
#mega-section-conc .scroll{overflow-x:hidden!important;max-width:100%!important}
#mega-section-conc .grid2{display:grid!important;grid-template-columns:minmax(0,1.15fr) minmax(0,.85fr)!important;gap:14px!important}
#mega-section-conc .js-plotly-plot,#mega-section-conc .plot-container,#mega-section-conc .svg-container{max-width:100%!important}
@media(max-width:900px){#mega-section-conc .grid2{grid-template-columns:1fr!important}}"""

# ── Diversidade / Soft Power ──
DIV_SP_SECTION_CSS = """/* ── Diversidade + Soft Power ── */
#mega-section-div,#mega-section-sp{background:var(--bg);color:var(--text)}"""

# ═══════════════════════════════════════════════════════════════════════════
# CSS — RESPONSIVE
# ═══════════════════════════════════════════════════════════════════════════
MEGA_RESPONSIVE_CSS = """/* ── Responsive ── */
@media (max-width: 900px) {
  .mega-sidebar{position:fixed;left:0;top:0;bottom:0;z-index:200;transform:translateX(-100%);width:240px;min-width:240px}
  .mega-sidebar.mobile-open{transform:translateX(0);box-shadow:4px 0 24px rgba(0,0,0,.5)}
  body.sidebar-collapsed .mega-sidebar{width:240px;min-width:240px;opacity:1;padding:8px 0}
  .sidebar-toggle{display:none}
  .mobile-menu-btn{display:flex!important}
  .mega-header-top{padding:12px 16px 10px}
  .mega-meto-bar{padding:0 16px}
  .mobile-overlay{display:block!important}
  .mobile-overlay.active{opacity:1;pointer-events:auto}
}
@media (min-width: 901px) {
  .mobile-menu-btn{display:none!important}
  .mobile-overlay{display:none!important}
}

/* ── UI stabilization ── */
html{height:100%;max-width:100%}
body{height:100vh;height:100dvh;min-height:0;overflow:hidden;max-width:100vw}
.mega-content{width:auto}
#mega-section-pr{height:100%}
#mega-section-cmp,#mega-section-cs,#mega-section-cl,#mega-section-div,#mega-section-sp{min-height:100%}
.vg-subnav,#mega-section-cs .tabs,#mega-section-pr .tab-bar,#conc-section .conc-subnav{
  overflow-x:auto;overflow-y:hidden;scrollbar-width:thin;
}
.vg-subnav::-webkit-scrollbar,#mega-section-cs .tabs::-webkit-scrollbar,
#mega-section-pr .tab-bar::-webkit-scrollbar,#conc-section .conc-subnav::-webkit-scrollbar{height:5px}
.vg-subtab,#mega-section-cs .tab,#mega-section-pr .tab-btn,#conc-section .conc-tab{flex-shrink:0}
.card,.kpi,.kpi-card,.chart-wrap,.quad-wrap,.comp-card{min-width:0}
#mega-section-cmp .card,#mega-section-cs .panel,#pr-ticket-panel .card,#conc-section .card{overflow-x:auto}
.grid-2,.kpi-bar,#conc-section .grid2{min-width:0}
.plotly-graph-div{max-width:100%}

@media (max-width: 900px) {
  .mega-header-top{gap:10px;align-items:flex-start;padding:12px 14px 10px}
  .mega-header-icon{width:40px;height:40px;border-radius:10px}
  .mega-header-titles h1{font-size:15px;line-height:1.2}
  .mega-header-titles p{font-size:10px;line-height:1.45}
  .mega-header-meta{display:none}
  .mega-meto-wrap{display:none}
  .mega-content{width:100%}
  .vg-subnav{padding:0 12px}
  #mega-section-cmp .tab-content,#mega-section-cs .panel{padding:16px 14px!important}
  #pr-tab-clusters{padding:16px 14px!important}
  .grid-2,.kpi-bar,#conc-section .grid2{grid-template-columns:1fr!important}
  .plotly-graph-div{min-width:680px;min-height:280px}
  .modal{width:calc(100vw - 24px);max-height:calc(100dvh - 24px)}
}"""


# ═══════════════════════════════════════════════════════════════════════════
# ASSEMBLY HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def mega_css(cmp_css='', cs_css='', pr_css='', conc_css='', extra=''):
    """Full CSS for the mega panel (script 07).

    Parameters are section-specific CSS extracted from sub-panel HTML.
    They are now minimal since :root/body/globals are stripped and variable
    names are canonical — no normalization needed.
    """
    sections = [
        '/* ── BASE ── */',
        THEME_VARS, BASE_CSS,
        '/* ── COMPONENTS ── */',
        ALL_COMPONENTS,
        '/* ── MEGA LAYOUT ── */',
        MEGA_HEADER_CSS, MEGA_METO_CSS, MEGA_SIDEBAR_CSS,
        MEGA_CONTENT_CSS, VG_SUBNAV_CSS,
        '/* ── SECTION OVERRIDES ── */',
        CMP_SECTION_CSS, CS_SECTION_CSS, PR_SECTION_CSS,
        CONC_SECTION_CSS, DIV_SP_SECTION_CSS,
        '/* ── SUB-PANEL CSS ── */',
        f'/* CMP */\n{cmp_css}' if cmp_css else '',
        f'/* CS */\n{cs_css}' if cs_css else '',
        f'/* PR */\n{pr_css}' if pr_css else '',
        f'/* CONC */\n{conc_css}' if conc_css else '',
        '/* ── EXTRA ── */',
        extra,
        '/* ── RESPONSIVE ── */',
        MEGA_RESPONSIVE_CSS,
    ]
    return '\n'.join(s for s in sections if s)
