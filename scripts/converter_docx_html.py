"""
converter_docx_html.py
======================
Converte um arquivo .docx em HTML estilizado.
Se o arquivo convertido for o artigo de opinião, aplica automaticamente
links cruzados para as seções correspondentes no artigo de análise de dados.

Uso:
    python scripts/converter_docx_html.py output_final/meu_arquivo.docx

Saída: mesmo diretório do .docx, mesmo nome com extensão .html

Requisito: pandoc instalado (https://pandoc.org)
"""

import re
import subprocess
import sys
import pathlib

# ─────────────────────────────────────────────────────────────────────────────
# Links cruzados: artigo de opinião → artigo de análise de dados
# Adicione ou edite entradas aqui para ajustar os links.
# ─────────────────────────────────────────────────────────────────────────────
ANALYSIS_DOC = "Análise de Dados - Fomento Audiovisual Brasileiro.html"
PANEL_DOC    = "painel_fomento_audiovisual.html"

# ── Legendas/títulos de gráfico (parágrafo inteiro vira link) ────────────────
CROSSLINKS = [
    ("Investimento vs. Retorno por Grupo de financiamento",
     ANALYSIS_DOC + "#investimento-vs.-retorno-por-grupo"),
    ("Ranking de ROI Doméstico por categoria de chamada",
     ANALYSIS_DOC + "#roi-doméstico-médio-por-categoria"),
    ("Ranking de ROI Internacional por categoria de chamada",
     ANALYSIS_DOC + "#síntese-comparativa-ranking-por-métrica"),
    ("Vocação Comercial vs. Alcance Internacional por categoria",
     ANALYSIS_DOC + "#vocação-comercial-alcance-internacional-por-categoria"),
    ("Produtoras ativas vs. ticket médio anual por produtora",
     ANALYSIS_DOC + "#ticket-médio-por-tier-e-viabilidade-operacional"),
    ("Curva de Lorenz",
     ANALYSIS_DOC + "#distribuição-do-investimento-curva-de-lorenz"),
    ("Quatro clusters de produtoras",
     ANALYSIS_DOC + "#retorno-por-cluster-múltiplas-métricas"),
]

# ── Dados no corpo do texto (só o trecho estatístico vira link) ──────────────
# Formato: (frase exata no texto, URL de destino)
TEXT_CROSSLINKS = [
    # → Análise de dados
    ("ROI doméstico agregado de 1,3x",
     ANALYSIS_DOC + "#roi-doméstico-médio-por-categoria"),
    ("ROI internacional 56% superior",
     ANALYSIS_DOC + "#vocação-comercial-alcance-internacional-por-categoria"),
    ("coeficiente de Gini da distribuição do investimento FSA por produtora é 0,61",
     ANALYSIS_DOC + "#distribuição-do-investimento-curva-de-lorenz"),
    ("Das 412 obras com investimento direto do FSA no período, 171",
     ANALYSIS_DOC + "#capital-fsa-sem-retorno-mensurável"),
    ("3,5\u00d7 de probabilidade adicional",
     ANALYSIS_DOC + "#probabilidade-de-transição-curta-longa-em-festival"),
    ("O cluster de Retorno Internacional",
     ANALYSIS_DOC + "#distribuição-do-roi-internacional-por-cluster"),
    # → Painel interativo
    ("mobilizou R$ 4,1 bilhões em investimento total",
     PANEL_DOC),
    ("27 obras responsáveis por 75% da renda",
     PANEL_DOC),
    ("cresceu 5,5 vezes",
     PANEL_DOC),
]


def _apply_text_crosslinks(html: str) -> str:
    """Envolve frases-chave com dados no corpo do texto com links."""
    _style = (
        "color:inherit;text-decoration:none;"
        "border-bottom:1.5px solid rgba(26,79,138,.3);"
        "padding-bottom:1px"
    )
    count = 0
    for phrase, href in TEXT_CROSSLINKS:
        escaped = re.escape(phrase)
        new_html = re.sub(
            escaped,
            f'<a href="{href}" target="_blank" style="{_style}">{phrase}</a>',
            html,
            count=1
        )
        if new_html != html:
            count += 1
        html = new_html
    return html, count


def _apply_crosslinks(html: str) -> str:
    """Envolve parágrafos de legenda com links para o artigo de análise."""
    _link_style = (
        "color:var(--muted);font-style:italic;font-size:13.5px;"
        "text-decoration:none;border-bottom:1px dashed var(--rule);"
        "padding-bottom:1px;transition:color .15s"
    )
    _sup_style = (
        "font-size:10px;margin-left:5px;opacity:0.55;font-style:normal"
    )
    for snippet, anchor in CROSSLINKS:
        # Escapa o snippet para uso em regex (lida com parênteses, pontos etc.)
        escaped = re.escape(snippet)
        pattern = re.compile(
            r'(<p>)(' + escaped + r'[^<]*)(\s*</p>)',
            re.IGNORECASE
        )
        href = ANALYSIS_DOC + anchor
        def _make_link(m, href=href):
            return (
                m.group(1)
                + f'<a href="{href}" target="_blank" style="{_link_style}">'
                + m.group(2)
                + f'<sup style="{_sup_style}">↗</sup>'
                + '</a>'
                + m.group(3)
            )
        html = pattern.sub(_make_link, html)
    return html

CSS = """
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Lora:ital,wght@0,400;0,600;1,400&display=swap');

  :root {
    --bg:      #fafaf8;
    --surface: #f2f1ee;
    --border:  #dddbd6;
    --text:    #1a1917;
    --muted:   #6b6860;
    --accent:  #1a4f8a;
    --link:    #1a4f8a;
    --rule:    #c8c5bf;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Lora', Georgia, serif;
    font-size: 16px;
    line-height: 1.8;
    background: var(--bg);
    color: var(--text);
    padding: 56px 24px 80px;
  }
  .page {
    max-width: 760px;
    margin: 0 auto;
  }
  h1 {
    font-family: 'Inter', system-ui, sans-serif;
    font-size: 2rem;
    font-weight: 600;
    color: var(--accent);
    margin-bottom: 10px;
    line-height: 1.25;
    letter-spacing: -0.02em;
  }
  h2 {
    font-family: 'Inter', system-ui, sans-serif;
    font-size: 1.2rem;
    font-weight: 600;
    color: var(--text);
    margin: 42px 0 14px;
    padding-bottom: 8px;
    border-bottom: 1px solid var(--rule);
    letter-spacing: -0.01em;
  }
  h3 {
    font-family: 'Inter', system-ui, sans-serif;
    font-size: 1rem;
    font-weight: 600;
    color: var(--text);
    margin: 28px 0 8px;
  }
  h4, h5, h6 {
    font-family: 'Inter', system-ui, sans-serif;
    font-size: 0.9rem;
    font-weight: 500;
    color: var(--muted);
    margin: 18px 0 6px;
    text-transform: uppercase;
    letter-spacing: 0.04em;
  }
  p  { margin-bottom: 16px; }
  ul, ol { margin: 10px 0 16px 24px; }
  li { margin-bottom: 6px; }
  a  { color: var(--link); text-decoration: underline; text-underline-offset: 3px; }
  a:hover { color: #0d3261; }
  blockquote {
    border-left: 3px solid var(--accent);
    margin: 20px 0;
    padding: 10px 20px;
    background: var(--surface);
    border-radius: 0 6px 6px 0;
    color: var(--muted);
    font-style: italic;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    margin: 20px 0;
    font-family: 'Inter', system-ui, sans-serif;
    font-size: 13px;
  }
  th {
    background: var(--surface);
    color: var(--muted);
    font-weight: 600;
    text-align: left;
    padding: 9px 13px;
    border-bottom: 2px solid var(--border);
    text-transform: uppercase;
    font-size: 11px;
    letter-spacing: 0.05em;
  }
  td {
    padding: 8px 13px;
    border-bottom: 1px solid var(--border);
  }
  tr:hover td { background: var(--surface); }
  code {
    font-family: 'Cascadia Code', 'Fira Code', monospace;
    font-size: 0.85em;
    background: var(--surface);
    padding: 2px 6px;
    border-radius: 4px;
    color: var(--accent);
  }
  pre { background: var(--surface); padding: 16px; border-radius: 8px; overflow-x: auto; }
  pre code { background: none; padding: 0; }
  hr { border: none; border-top: 1px solid var(--rule); margin: 32px 0; }
  img { max-width: 100%; border-radius: 4px; display: block; margin: 24px auto; }
  figure { margin: 24px 0; text-align: center; }
  figcaption { font-size: 13px; color: var(--muted); margin-top: 8px; font-style: italic; }
</style>
"""

TEMPLATE = """\
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
{css}
</head>
<body>
<div class="page">
$body$
</div>
</body>
</html>
"""


def convert(docx_path: str) -> None:
    src = pathlib.Path(docx_path).resolve()
    if not src.exists():
        print(f"Erro: arquivo não encontrado: {src}")
        sys.exit(1)
    if src.suffix.lower() != ".docx":
        print(f"Erro: esperado um arquivo .docx, recebido: {src.name}")
        sys.exit(1)

    dst = src.with_suffix(".html")
    title = src.stem.replace("_", " ").replace("-", " ")

    # Grava template temporário
    import tempfile, os
    tpl_content = TEMPLATE.format(title=title, css=CSS)
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False,
                                    encoding="utf-8") as tpl:
        tpl.write(tpl_content)
        tpl_path = tpl.name

    try:
        result = subprocess.run(
            ["pandoc", str(src), "-o", str(dst),
             "--template", tpl_path,
             "--standalone",
             "--embed-resources",
             "--wrap=none"],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print("Erro no pandoc:")
            print(result.stderr)
            sys.exit(1)

        # Aplica links cruzados se algum snippet bater no HTML gerado
        html_out = dst.read_text(encoding="utf-8")
        html_linked = _apply_crosslinks(html_out)
        html_linked, n_text = _apply_text_crosslinks(html_linked)
        if html_linked != html_out:
            dst.write_text(html_linked, encoding="utf-8")
            n_cap = sum(1 for s, _ in CROSSLINKS if s in html_linked)
            print(f"  {n_cap} legenda(s) + {n_text} trecho(s) de texto linkados")

        print(f"Convertido: {dst}")
    finally:
        os.unlink(tpl_path)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python scripts/converter_docx_html.py <arquivo.docx>")
        sys.exit(1)
    convert(sys.argv[1])
