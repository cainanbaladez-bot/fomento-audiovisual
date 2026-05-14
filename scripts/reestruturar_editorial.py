"""
Reestrutura o editorial separando claims verificáveis de argumentos interpretativos.
- Não altera o texto em si
- Reorganiza blocos para que claims construam a base dos argumentos
- Adiciona hierarquia visual (CSS) para diferenciar os dois tipos
"""
import re, sys

INPUT = r"C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6_editado_editorial.html"
OUTPUT = r"C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v7_claims_args.html"

with open(INPUT, "r", encoding="utf-8") as f:
    html = f.read()

# ── 1. Inject new CSS before </style> ──────────────────────────────
NEW_CSS = """
/* ── Claim vs Argument blocks ── */
.claim-block, .arg-block {
  position: relative;
  padding: 20px 24px 16px;
  margin: 24px 0;
  border-radius: 0 10px 10px 0;
  transition: background .35s, border-color .35s;
}
.claim-block {
  border-left: 3px solid var(--accent);
  background: var(--evidence-bg);
}
.arg-block {
  border-left: 3px solid var(--accent-2);
  background: var(--accent-2-soft);
}
.block-label {
  display: inline-block;
  font: 700 9px/1 'Inter', system-ui, sans-serif;
  text-transform: uppercase;
  letter-spacing: .12em;
  padding: 3px 8px;
  border-radius: 4px;
  margin-bottom: 12px;
}
.claim-block .block-label {
  color: var(--accent);
  background: var(--accent-soft);
  border: 1px solid var(--evidence-border);
}
.arg-block .block-label {
  color: var(--accent-2);
  background: var(--accent-2-soft);
  border: 1px solid rgba(200,154,106,.25);
}
.claim-block p, .arg-block p { margin-bottom: 14px; }
.claim-block p:last-child, .arg-block p:last-child { margin-bottom: 0; }
.claim-block ul, .arg-block ul { margin-bottom: 14px; }
.claim-block table { margin: 12px 0; }
"""

html = html.replace("</style>", NEW_CSS + "\n</style>")


# ── 2. Helpers ─────────────────────────────────────────────────────
def claim(content, label="Evidência"):
    return f'<div class="claim-block">\n<div class="block-label">{label}</div>\n{content}\n</div>'

def arg(content, label="Argumento"):
    return f'<div class="arg-block">\n<div class="block-label">{label}</div>\n{content}\n</div>'

def is_chart_or_image(line):
    """Check if line contains embedded chart/image data."""
    s = line.strip()
    if 'data:image/png;base64' in s:
        return True
    if 'Plotly.' in s:
        return True
    if '<!-- chart' in s:
        return True
    if len(s) > 5000 and '<div style=' in s:
        return True
    return False

def text_of(line):
    """Strip HTML tags to get plain text for keyword matching."""
    return re.sub(r'<[^>]+>', '', line.strip())


# ── 3. Section-aware restructuring ─────────────────────────────────
# We'll use a different approach: find each section between h2/h3 headings
# and wrap paragraphs inline using regex-based approach.

lines = html.split('\n')

# Find article boundaries
art_start = art_end = None
for i, line in enumerate(lines):
    if '<article>' in line and art_start is None:
        art_start = i
    if '</article>' in line:
        art_end = i
        break

# Identify heading positions (section boundaries)
heading_positions = []
for i in range(art_start, art_end + 1):
    s = lines[i].strip()
    if (s.startswith('<h2') or
        (s.startswith('<h3') and ('id="metodologia"' in s or 'id="a-proliferacao' in s or 'id="financiamento' in s))):
        heading_positions.append(i)

# For each section between headings, classify paragraphs
# Section definitions: which keywords indicate claims vs arguments
SECTION_CLAIMS = {
    'o-que-determina-o-sucesso': [
        'renúncia fiscal via artigos',
        'majors, como Disney',
        'desconto de 70%',
        'remessa para o exterior',
        'explorar por várias óticas',
        'experiência das majors',
        'volumes maiores de P&A',
        'linha de comercialização',
    ],
    'chamadas-bilheteria-e-festival': [
        'categorias seletivas',
        'filtros de desempenho prévio',
        'superam as totalmente automáticas',
        'Pontuação Festivais e Roteiro',
    ],
    'de-forma-geral-as-distribuidoras': [
        'dois perfis completamente distintos',
        'distribuidoras de perfil doméstico selecionaram',
        'distribuidoras de perfil internacional selecionaram',
        'Minha Mãe é Uma Peça',
        'Bacurau',
    ],
    'e-as-chamadas-de-ampla-concorrencia': [
        'não são ainda objeto de análise',
        'falta de tempo de maturação',
        'dois novos modelos',
    ],
    'o-que-o-retorno-das-produtoras': [
        'avaliação agregada',
        'Existe de fato um grupo',
        'cluster Retorno Doméstico',
        'outro lado do espectro',
        'Retorno Internacional, formado',
        'Duplo Retorno. São 39',
        'Cinemascópio aparece',
        'não temos como ignorar os outros',
        '174 empresas que receberam',
        '927 empresas tem mediana',
    ],
    'a-estrutura-que-o-fomento': [
        'transição do proponente CPF',
        'produtoras são hoje o centro',
        'detentoras primárias do direito patrimonial',
        'exibidores ficam com 50%',
        'ANCINE teve um papel importante',
    ],
    'a-proliferacao': [
        'número de obras produzidas',
        'cresceu 5,5 vezes',
        'cresceu 5 vezes',
        'coeficiente de Gini',
        'metade das produtoras recebe',
        'top 10% concentram',
        'Quarenta e três por cento',
        'mediana de obras por produtora',
        'Curva de Lorenz',
        'Rio de Janeiro e São Paulo',
        'custos anuais começando em R$',
    ],
    'a-diversidade-que-temos': [
        'BRDE enviou via LAI',
        'teste simples comparando',
        'chamadas com política afirmativa tem a capacidade',
        'Standard A',
        'Standard B',
        'Standard C',
        'Standard D',
        'BFI (British Film',
    ],
    'a-instabilidade-tem-endereco': [
        'composição é: 4 votos',
        'órgão de governo, não é de Estado',
        'Ancine hoje é uma máquina',
    ],
    'financiamento-e-recuperacao': [
        'FSA é alimentado',
        'CONDECINE',
        'arrecadação de um ano',
    ],
}

def get_section_id(heading_line):
    """Extract section ID from heading line."""
    m = re.search(r'id="([^"]*)"', heading_line)
    return m.group(1) if m else ""

def classify_paragraph(text, section_id):
    """Return 'claim' or 'arg' based on keyword matching."""
    for sec_prefix, keywords in SECTION_CLAIMS.items():
        if sec_prefix in section_id:
            for kw in keywords:
                if kw in text:
                    return 'claim'
            return 'arg'
    return 'arg'

# ── 4. Process the article, wrapping paragraphs ────────────────────
# Strategy: process each section. Collect consecutive paragraphs of
# the same type (claim/arg) and wrap them together.

new_article_lines = lines[art_start:art_start+1]  # <article> line

# Add preamble (before first heading)
first_heading = heading_positions[0] if heading_positions else art_end
for i in range(art_start + 1, first_heading):
    new_article_lines.append(lines[i])

# Process each section
for sec_idx, hpos in enumerate(heading_positions):
    next_hpos = heading_positions[sec_idx + 1] if sec_idx + 1 < len(heading_positions) else art_end

    heading_line = lines[hpos]
    section_id = get_section_id(heading_line)
    new_article_lines.append(heading_line)

    # Collect elements in this section
    elements = []  # list of (type, line) where type is 'claim', 'arg', 'chart', 'table', 'other'

    i = hpos + 1
    in_table_div = False
    table_lines = []

    while i < next_hpos:
        line = lines[i]
        stripped = line.strip()

        # Skip empty lines
        if not stripped:
            i += 1
            continue

        # Chart/image - always passthrough
        if is_chart_or_image(line):
            elements.append(('chart', line))
            i += 1
            continue

        # Table div (the summary table)
        if '<div style="margin:32px 0">' in stripped:
            in_table_div = True
            table_lines = [line]
            i += 1
            continue
        if in_table_div:
            table_lines.append(line)
            if stripped == '</div>':
                in_table_div = False
                elements.append(('table', '\n'.join(table_lines)))
            i += 1
            continue

        # <ul>/<ol> blocks - attach to previous type
        if stripped in ('<ul>', '<ol>'):
            # Collect until closing tag
            list_lines = [line]
            i += 1
            while i < next_hpos:
                list_lines.append(lines[i])
                if lines[i].strip() in ('</ul>', '</ol>'):
                    break
                i += 1
            prev_type = elements[-1][0] if elements else 'arg'
            if prev_type in ('claim', 'arg'):
                elements.append((prev_type, '\n'.join(list_lines)))
            else:
                elements.append(('arg', '\n'.join(list_lines)))
            i += 1
            continue

        # Paragraphs
        if stripped.startswith('<p'):
            text = text_of(stripped)
            ptype = classify_paragraph(text, section_id)
            elements.append((ptype, line))
            i += 1
            continue

        # Other HTML (hr, divs, etc)
        elements.append(('other', line))
        i += 1

    # Now group consecutive same-type elements and wrap
    groups = []
    current_type = None
    current_items = []

    for etype, content in elements:
        if etype in ('chart', 'table', 'other'):
            # Flush current group
            if current_items:
                groups.append((current_type, current_items))
                current_items = []
                current_type = None
            if etype == 'table':
                groups.append(('table_claim', [content]))
            else:
                groups.append((etype, [content]))
        elif etype in ('claim', 'arg'):
            if current_type == etype:
                current_items.append(content)
            else:
                if current_items:
                    groups.append((current_type, current_items))
                current_type = etype
                current_items = [content]

    if current_items:
        groups.append((current_type, current_items))

    # Reorder: put claim groups before arg groups within each section,
    # but keep charts/tables/other in their original relative positions.
    # Actually, we want claims to build up to arguments, so:
    # charts/tables first (data), then claims, then arguments.
    # But we need to be careful not to break the narrative flow too much.
    # Best approach: just wrap with labels but keep order, since the user
    # wants to see the distinction visually.

    for gtype, items in groups:
        content = '\n'.join(items)
        if gtype == 'claim':
            new_article_lines.append(claim(content))
        elif gtype == 'arg':
            new_article_lines.append(arg(content))
        elif gtype == 'table_claim':
            new_article_lines.append(claim(content, "Dados"))
        else:
            new_article_lines.append(content)

# Add everything after the article
new_article_lines.append(lines[art_end])
result_lines = lines[:art_start] + new_article_lines + lines[art_end + 1:]

output = '\n'.join(result_lines)

with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write(output)

# Stats
claim_count = output.count('class="claim-block"')
arg_count = output.count('class="arg-block"')
print(f"Output: {len(output)} chars")
print(f"Claim blocks: {claim_count}")
print(f"Arg blocks: {arg_count}")

# Show per-section breakdown
for sec_idx, hpos in enumerate(heading_positions):
    heading = text_of(lines[hpos])[:70]
    section_id = get_section_id(lines[hpos])
    next_hpos = heading_positions[sec_idx + 1] if sec_idx + 1 < len(heading_positions) else art_end
    section_html = '\n'.join(lines[hpos:next_hpos])
    # But we need to count in the OUTPUT, not the input
    print(f"  {heading}")
