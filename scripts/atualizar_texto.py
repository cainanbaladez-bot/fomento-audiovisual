#!/usr/bin/env python3
"""
atualizar_texto.py
==================
Converte o docx do texto de politica em HTML e publica em docs/
usando o template fixo (dark/light, topbar, modal de dados, TOC).

Uso:
    python scripts/atualizar_texto.py

Requisito: pandoc instalado (https://pandoc.org)

Fluxo:
  1. Converte o docx via pandoc (extrai body HTML)
  2. Corrige hierarquia de headings (tudo H2, conforme docx)
  3. Estiliza epigrafe
  4. Adiciona data-panel nos titulos com dados
  5. Reconstroi TOC no template
  6. Monta HTML final com template fixo e publica em docs/
"""
import re, subprocess, sys, os, tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCX = os.path.join(BASE, 'output_final',
                    'Uma pol\u00edtica de fomento baseada em evid\u00eancias_v6.docx')
DOCS_OUT = os.path.join(BASE, 'docs', 'politica.html')

# Template fixo (dark/light, topbar, modal, TOC)
SHELL_BEFORE = os.path.join(BASE, 'templates', 'politica_shell_before.html')
SHELL_AFTER  = os.path.join(BASE, 'templates', 'politica_shell_after.html')

# ══════════════════════════════════════════════════════════
# Config: heading id -> panel hash (None = no data link)
# ══════════════════════════════════════════════════════════
PANEL_MAP = {
    'o-que-determina-o-sucesso-em-cinema-e-o-investimento-das-maj': '#retorno-domestico',
    'chamadas-bilheteria-e-festival': '#categorias',
    'de-forma-geral-as-distribuidoras-selecionam-melhor-do-que-as': '#categorias',
    'mas-os-dados-revelam-dois-perfis-completamente-distintos': '#categorias',
    'o-que-o-retorno-das-produtoras-revela': '#produtoras',
    'a-estrutura-que-o-fomento-desconstruiu': '#concentracao',
    'a-proliferacao-das-produtoras-vs-sustentabilidade': '#concentracao',
    'a-diversidade-que-temos': '#diversidade',
    'suat-internacional-curtas': '#curtas-longas',
}

# Epigraph: quotes to wrap
EPIGRAPH_QUOTES = [
    {
        'match': 'Ningu\u00e9m quer entrar em um mercado',
        'cite': '\u2014 Rodrigo Teixeira',
        'url': None,
    },
    {
        'match': 'Muitos dos profissionais mais experientes',
        'cite': '\u2014 Ana Paula Sousa',
        'cite_source': 'IndieWire, 2026',
        'url': 'https://www.indiewire.com/features/interviews/brazilian-film-industry-thrives-oscars-cannes-1235163028/',
    },
    {
        'match': 'Marte Um',
        'cite': '\u2014 Gabriel Martins',
        'cite_source': 'G1, 2026',
        'url': None,
    },
]


def step1_pandoc():
    """Convert docx to raw HTML body via pandoc."""
    print('[1/6] Convertendo docx via pandoc...')
    if not os.path.exists(DOCX):
        print(f'  ERRO: {DOCX} nao encontrado')
        sys.exit(1)

    with tempfile.NamedTemporaryFile('w', suffix='.html', delete=False, encoding='utf-8') as f:
        tmp = f.name

    try:
        r = subprocess.run(
            ['pandoc', DOCX, '-o', tmp, '--wrap=none', '--embed-resources', '--standalone'],
            capture_output=True, text=True
        )
        if r.returncode != 0:
            print(f'  ERRO pandoc: {r.stderr}')
            sys.exit(1)

        with open(tmp, 'r', encoding='utf-8') as f:
            raw = f.read()

        m = re.search(r'<body[^>]*>(.*)</body>', raw, re.DOTALL)
        body = m.group(1).strip() if m else raw
        print(f'  Body extraido: {len(body):,} chars')
        return body
    finally:
        os.unlink(tmp)


def _slugify(text):
    """Convert text to URL-friendly slug."""
    import unicodedata
    text = unicodedata.normalize('NFKD', text)
    text = text.encode('ascii', 'ignore').decode('ascii')
    text = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'[-\s]+', '-', text).strip('-')[:60]


def step2_fix_headings(body):
    """Convert bold paragraphs and known sections to H2."""
    print('[2/6] Corrigindo hierarquia de headings...')

    body = re.sub(r'<h3(\s)', r'<h2\1', body)
    body = re.sub(r'</h3>', '</h2>', body)

    def bold_to_h2(m):
        inner = m.group(1)
        clean = re.sub(r'<[^>]+>', '', inner).strip()
        slug = _slugify(clean)
        return f'<h2 id="{slug}">{clean}</h2>'

    body, n = re.subn(r'<p><strong>([^<]+)</strong></p>', bold_to_h2, body)
    print(f'  Convertidos {n} paragrafos bold -> H2')

    extra_headings = [
        ('Chamadas que usam desempenho de bilheteria', 'chamadas-bilheteria-e-festival'),
        ('Mas os dados revelam dois perfis', 'mas-os-dados-revelam-dois-perfis-completamente-distintos'),
        ('O que o retorno das produtoras revela', 'o-que-o-retorno-das-produtoras-revela'),
        ('A estrutura que o fomento', 'a-estrutura-que-o-fomento-desconstruiu'),
        ('A instabilidade tem endere', 'a-instabilidade-tem-endereco-o-cgfsa'),
        ('O que o FSA poderia ser', 'o-que-o-fsa-poderia-ser'),
        ('Joint Ventures, as SPEs brasileiras', 'joint-ventures-as-spes-brasileiras'),
        ('Previsibilidade e planejamento', 'previsibilidade-e-planejamento-de-longo-prazo'),
        ('Novas possibilidades de investimento', 'novas-possibilidades-de-investimento'),
        ('Renova\u00e7\u00e3o depende do desempenho', 'renovacao-depende-do-desempenho'),
        ('Seletivos', 'seletivos'),
        ('SUAT Internacional Curtas', 'suat-internacional-curtas'),
        ('Ampla concorr\u00eancia a partir do desenvolvimento', 'ampla-concorrencia-a-partir-do-desenvolvimento'),
        ('Uma distribui\u00e7\u00e3o poss\u00edvel', 'uma-distribuicao-possivel'),
    ]

    for search_text, slug in extra_headings:
        idx = body.find(search_text)
        if idx < 0:
            continue
        before = body[max(0, idx-20):idx]
        if '<h2' in before:
            continue
        p_start = body.rfind('<p>', 0, idx)
        if p_start < 0 or idx - p_start > 50:
            continue
        p_end = body.find('</p>', idx)
        if p_end < 0:
            continue
        p_end += 4
        inner = re.sub(r'<[^>]+>', '', body[p_start:p_end]).strip()
        h2 = f'<h2 id="{slug}">{inner}</h2>'
        body = body[:p_start] + h2 + body[p_end:]
        print(f'  Extra H2: {slug}')

    h2_count = len(re.findall(r'<h2[^>]*>', body))
    print(f'  Total H2: {h2_count}')
    return body


def step3_epigraph(body):
    """Style the opening quotes as an epigraph block."""
    print('[3/6] Estilizando epigrafe...')

    pattern = (
        r'<p>\u201c[^<]*Ningu\u00e9m quer entrar[^<]*</p>'
        r'\s*<p>\u2014\s*Rodrigo Teixeira</p>'
        r'\s*<p>\u201c[^<]*profissionais mais experientes[^<]*</p>'
        r'\s*<p>\u2014\s*Ana Paula Sousa[^<]*</p>'
        r'\s*<p>\u201c[^<]*Marte Um[^<]*</p>'
        r'\s*<p>\u2014\s*Gabriel Martins[^<]*</p>'
        r'\s*<p>\u2014{2,3}</p>'
    )

    m = re.search(pattern, body, re.DOTALL)
    if not m:
        idx1 = body.find('\u201cNingu\u00e9m quer entrar')
        idx2 = body.find('\u2014\u2014\u2014</p>', idx1 if idx1 > 0 else 0)
        if idx1 > 0 and idx2 > 0:
            p_start = body.rfind('<p>', 0, idx1)
            p_end = idx2 + len('\u2014\u2014\u2014</p>')
            old_block = body[p_start:p_end]
        else:
            print('  AVISO: epigrafe nao encontrada')
            return body
    else:
        old_block = m.group(0)

    quote_paras = re.findall(r'<p>(\u201c[^<]+)</p>', old_block)

    bqs = []
    for i, q_text in enumerate(quote_paras):
        qinfo = EPIGRAPH_QUOTES[i] if i < len(EPIGRAPH_QUOTES) else {}
        cite_html = qinfo.get('cite', '')
        if qinfo.get('url'):
            cite_html += f' (<a href="{qinfo["url"]}" target="_blank" rel="noopener">{qinfo.get("cite_source", "")}</a>)'
        elif qinfo.get('cite_source'):
            cite_html += f' ({qinfo["cite_source"]})'
        bqs.append(f'<blockquote>\n<p>{q_text}</p>\n<cite>{cite_html}</cite>\n</blockquote>')

    new_block = '<div class="epigraph">\n' + '\n'.join(bqs) + '\n</div>'
    body = body.replace(old_block, new_block, 1)
    print(f'  Epigrafe formatada ({len(quote_paras)} citacoes)')
    return body


def step4_data_panels(body):
    """Add data-panel attributes to H2 headings that link to data."""
    print('[4/6] Adicionando data-panel aos topicos...')

    body = re.sub(r'(<h2[^>]*?) data-panel="[^"]*"', r'\1', body)

    count = 0
    for hid, panel in PANEL_MAP.items():
        old = f'<h2 id="{hid}"'
        if old in body:
            new = f'<h2 id="{hid}" data-panel="{panel}"'
            body = body.replace(old, new, 1)
            count += 1

    print(f'  {count} topicos com link para dados')
    return body


def step5_rebuild_toc(shell_before, body):
    """Rebuild TOC links in the shell from actual H2 ids in the body."""
    print('[5/6] Reconstruindo TOC...')

    h2s = re.findall(r'<h2[^>]*id="([^"]+)"[^>]*>(.*?)</h2>', body, re.DOTALL)
    toc_links = []
    for hid, inner in h2s:
        clean = re.sub(r'<[^>]+>', '', inner).strip()
        short = clean[:50] + ('...' if len(clean) > 50 else '')
        toc_links.append(f'  <a href="#{hid}">{short}</a>')

    toc_html = '\n'.join(toc_links)

    # Replace TOC content between toc-label and toc-sep
    toc_pattern = r'(<div class="toc-label">Sum[^<]*</div>\s*)\n.*?\n(\s*<div class="toc-sep")'
    m = re.search(toc_pattern, shell_before, re.DOTALL)
    if m:
        shell_before = shell_before[:m.end(1)] + '\n' + toc_html + '\n' + shell_before[m.start(2):]
        print(f'  TOC reconstruido com {len(h2s)} itens')
    else:
        print('  AVISO: padrao TOC nao encontrado no template')

    return shell_before


def step6_assemble_and_publish(body, shell_before):
    """Assemble final HTML from template + body and save to docs/."""
    print('[6/6] Montando e publicando...')

    with open(SHELL_AFTER, 'r', encoding='utf-8') as f:
        shell_after = f.read()

    html = shell_before + '\n' + body + '\n' + shell_after

    os.makedirs(os.path.dirname(DOCS_OUT), exist_ok=True)
    with open(DOCS_OUT, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'  Salvo: {DOCS_OUT}')
    print(f'  Tamanho: {len(html):,} chars')


def main():
    print('='*56)
    print('  Atualizando texto: docx -> docs/politica.html')
    print('='*56)
    print()

    with open(SHELL_BEFORE, 'r', encoding='utf-8') as f:
        shell_before = f.read()

    body = step1_pandoc()
    body = step2_fix_headings(body)
    body = step3_epigraph(body)
    body = step4_data_panels(body)
    shell_before = step5_rebuild_toc(shell_before, body)
    step6_assemble_and_publish(body, shell_before)

    print()
    print('Pronto! Faca commit e push no GitHub Desktop.')


if __name__ == '__main__':
    main()
