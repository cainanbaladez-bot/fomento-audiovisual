"""
Reestrutura editorial v6 → v7 baseado nos comentários do autor.

Estrutura final:
  Intro + Metodologia (mantém)
  Contexto: Renúncia fiscal e majors (pré-claim)
  Claim 1: Critérios de seleção determinam o perfil de retorno
  Claim 2: Distribuidoras como proponentes superam produtoras
  [Colapsável] Nota: Ampla concorrência
  Claim 3: Produtoras revelam 5 perfis distintos de retorno
  Premissa interpretativa: A estrutura que o fomento (des)construiu
  Claim 4: Proliferação reduziu escala sem reduzir concentração
  Claim 5: Políticas afirmativas reduzem desigualdade entre selecionados
  A instabilidade tem endereço (argumento)
  Uma nota sobre financiamento (reposicionada)
  Propostas: O que o FSA poderia ser
"""
import re, sys

INPUT = r"C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6_editado_editorial_backup.html"
OUTPUT = r"C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências.html"

with open(INPUT, "r", encoding="utf-8") as f:
    html = f.read()

# ── Extract parts ──────────────────────────────────────────────────
lines = html.split('\n')

# Find key line numbers
def find_line(pattern, start=0):
    for i in range(start, len(lines)):
        if pattern in lines[i]:
            return i
    return None

def extract_range(start, end):
    """Extract lines start..end-1 as string."""
    return '\n'.join(lines[start:end])

# Key positions
l_article = find_line('<article>')
a = l_article  # search from article start
l_hero_start = find_line('class="article-hero"', a)
l_epigraph_start = find_line('class="epigraph"', a)
l_sobre = find_line('id="sobre-este-texto"', a)
l_intro_p1 = find_line('Essas três falas na epígrafe', a)
l_metodologia = find_line('Metodologia e indicadores', a)
l_h2_sucesso = find_line('<h2 id="o-que-determina-o-sucesso', a)
l_h2_chamadas = find_line('<h2 id="chamadas-bilheteria-e-festival', a)
l_h2_distrib = find_line('<h2 id="de-forma-geral-as-distribuidoras', a)
l_h2_ampla = find_line('<h2 id="e-as-chamadas-de-ampla-concorrencia', a)
l_h2_produtoras = find_line('<h2 id="o-que-o-retorno-das-produtoras', a)
l_h2_estrutura = find_line('<h2 id="a-estrutura-que-o-fomento', a)
l_h3_proliferacao = find_line('<h3 id="a-proliferacao-das-produtoras', a)
l_h2_diversidade = find_line('<h2 id="a-diversidade-que-temos', a)
l_h2_instabilidade = find_line('<h2 id="a-instabilidade-tem-endereco', a)
l_h3_financiamento = find_line('<h3 id="financiamento-e-recuperacao', a)
l_h2_propostas = find_line('<h2 id="o-que-o-fsa-poderia-ser', a)
l_article_end = find_line('</article>', a)

print("Line positions found:")
for name, val in [
    ('article', l_article), ('hero', l_hero_start), ('sobre', l_sobre),
    ('intro', l_intro_p1), ('metodologia', l_metodologia),
    ('sucesso', l_h2_sucesso), ('chamadas', l_h2_chamadas),
    ('distrib', l_h2_distrib), ('ampla', l_h2_ampla),
    ('produtoras', l_h2_produtoras), ('estrutura', l_h2_estrutura),
    ('proliferacao', l_h3_proliferacao), ('diversidade', l_h2_diversidade),
    ('instabilidade', l_h2_instabilidade), ('financiamento', l_h3_financiamento),
    ('propostas', l_h2_propostas), ('article_end', l_article_end)
]:
    print(f"  {name}: {val}")

# ── Extract each section's content ─────────────────────────────────
head_and_css = '\n'.join(lines[:l_article])
hero_to_sobre = extract_range(l_hero_start, l_sobre + 4)  # includes closing </div>
intro_paragraphs = extract_range(l_intro_p1, l_metodologia)
metodologia = extract_range(l_metodologia, l_h2_sucesso)
sec_sucesso = extract_range(l_h2_sucesso, l_h2_chamadas)
sec_chamadas = extract_range(l_h2_chamadas, l_h2_distrib)
sec_distrib = extract_range(l_h2_distrib, l_h2_ampla)
sec_ampla = extract_range(l_h2_ampla, l_h2_produtoras)
sec_produtoras = extract_range(l_h2_produtoras, l_h2_estrutura)
sec_estrutura = extract_range(l_h2_estrutura, l_h3_proliferacao)
sec_proliferacao = extract_range(l_h3_proliferacao, l_h2_diversidade)
sec_diversidade = extract_range(l_h2_diversidade, l_h2_instabilidade)
sec_instabilidade = extract_range(l_h2_instabilidade, l_h3_financiamento)
sec_financiamento = extract_range(l_h3_financiamento, l_h2_propostas)
sec_propostas = extract_range(l_h2_propostas, l_article_end)
footer = '\n'.join(lines[l_article_end:])

# ── New CSS ────────────────────────────────────────────────────────
NEW_CSS = """
/* ── Claim sections ── */
.claim-section {
  margin: 56px 0 48px;
  scroll-margin-top: 64px;
}
.claim-header {
  display: flex;
  align-items: flex-start;
  gap: 16px;
  padding: 28px 0 14px;
  border-top: 1px solid var(--rule);
}
.claim-num {
  display: flex;
  align-items: center;
  justify-content: center;
  min-width: 36px;
  height: 36px;
  background: var(--accent);
  color: #fff;
  border-radius: 50%;
  font: 800 15px/1 'Inter', system-ui, sans-serif;
  flex-shrink: 0;
  margin-top: 2px;
}
.claim-header h2 {
  margin: 0;
  padding: 0;
  border: none;
  font-size: 20px;
}
.claim-header h2::before { display: none; }

/* ── Evidence summary box ── */
.evidence-box {
  margin: 24px 0 32px;
  border: 1px solid var(--evidence-border);
  border-radius: 12px;
  overflow: hidden;
  font-family: 'Inter', system-ui, sans-serif;
  font-size: 13px;
  line-height: 1.6;
  background: var(--evidence-bg);
  transition: background .35s, border-color .35s;
}
.eb-grid {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 0;
}
@media(max-width:700px) {
  .eb-grid { grid-template-columns: 1fr; }
}
.eb-col {
  padding: 18px 20px;
  border-right: 1px solid var(--evidence-border);
}
.eb-col:last-child { border-right: none; }
@media(max-width:700px) {
  .eb-col { border-right: none; border-bottom: 1px solid var(--evidence-border); }
  .eb-col:last-child { border-bottom: none; }
}
.eb-label {
  font: 700 9px/1 'Inter', system-ui, sans-serif;
  text-transform: uppercase;
  letter-spacing: .12em;
  margin-bottom: 10px;
  display: flex;
  align-items: center;
  gap: 6px;
}
.eb-col:nth-child(1) .eb-label { color: var(--accent); }
.eb-col:nth-child(2) .eb-label { color: var(--accent-2); }
.eb-col:nth-child(3) .eb-label { color: var(--green, #5a9a6a); }
.eb-col ul {
  margin: 0;
  padding: 0 0 0 16px;
  color: var(--ink);
}
.eb-col li {
  margin: 5px 0;
  font-size: 12.5px;
}
.eb-col p {
  margin: 0;
  color: var(--ink);
  font-size: 12.5px;
}

/* ── Context / pre-claim section ── */
.context-section {
  margin: 48px 0;
  padding: 28px 28px 24px;
  background: var(--paper-2);
  border: 1px solid var(--rule);
  border-radius: 14px;
  transition: background .35s, border-color .35s;
}
.context-section h2 {
  margin: 0 0 16px;
  padding: 0;
  border: none;
  font-size: 19px;
  color: var(--accent-2);
}
.context-section h2::before {
  background: var(--accent-2);
}

/* ── Opinion / argument section ── */
.opinion-section {
  margin: 48px 0;
  padding: 28px 28px 24px;
  border-left: 3px solid var(--accent-2);
  background: var(--accent-2-soft);
  border-radius: 0 14px 14px 0;
  transition: background .35s;
}
.opinion-label {
  display: inline-block;
  font: 700 9px/1 'Inter', system-ui, sans-serif;
  text-transform: uppercase;
  letter-spacing: .12em;
  padding: 3px 10px;
  border-radius: 4px;
  color: var(--accent-2);
  background: rgba(200,154,106,.15);
  border: 1px solid rgba(200,154,106,.25);
  margin-bottom: 16px;
}
.opinion-section h2 {
  margin: 0 0 16px;
  padding: 0;
  border: none;
  font-size: 19px;
}
.opinion-section h2::before { display: none; }
.opinion-section h3 {
  margin: 24px 0 12px;
}

/* ── Collapsible minor section ── */
details.collapsible {
  margin: 32px 0;
  border: 1px solid var(--rule);
  border-radius: 12px;
  overflow: hidden;
  transition: border-color .35s;
}
details.collapsible summary {
  padding: 16px 24px;
  font: 600 15px/1.4 'Inter', system-ui, sans-serif;
  color: var(--ink-strong);
  cursor: pointer;
  background: var(--paper-2);
  list-style: none;
  display: flex;
  align-items: center;
  gap: 10px;
  transition: background .35s, color .35s;
}
details.collapsible summary::-webkit-details-marker { display: none; }
details.collapsible summary::before {
  content: '▸';
  font-size: 12px;
  color: var(--accent);
  transition: transform .2s;
}
details.collapsible[open] summary::before {
  transform: rotate(90deg);
}
details.collapsible summary:hover {
  background: var(--accent-soft);
  color: var(--accent);
}
details.collapsible .collapsible-body {
  padding: 20px 24px;
}

/* ── Aside note (for RLP/retorno and financiamento) ── */
.aside-note {
  margin: 36px 0;
  padding: 20px 24px;
  background: var(--paper-2);
  border: 1px solid var(--rule);
  border-radius: 10px;
  font-size: 16px;
  line-height: 1.75;
  transition: background .35s;
}
.aside-note h3 {
  font-size: 15px;
  margin: 0 0 12px;
  color: var(--muted);
}
.aside-note h3::before { display: none; }

/* ── Green var fallback ── */
:root, [data-theme="dark"] { --green: #5a9a6a; }
[data-theme="light"] { --green: #2d7a3e; }
"""

# Inject CSS
head_and_css = head_and_css.replace("</style>", NEW_CSS + "\n</style>")

# ── Build evidence boxes ───────────────────────────────────────────
def evidence_box(evidence_items, limitation_items, conclusion_text):
    ev = '\n'.join(f'<li>{e}</li>' for e in evidence_items)
    lm = '\n'.join(f'<li>{l}</li>' for l in limitation_items)
    return f'''<div class="evidence-box">
<div class="eb-grid">
<div class="eb-col">
<div class="eb-label">&#x1F4CA; Evidência</div>
<ul>{ev}</ul>
</div>
<div class="eb-col">
<div class="eb-label">&#x26A0; Limitações</div>
<ul>{lm}</ul>
</div>
<div class="eb-col">
<div class="eb-label">&#x2192; Conclusão</div>
<p>{conclusion_text}</p>
</div>
</div>
</div>'''

# ── Build claim header ─────────────────────────────────────────────
def claim_header(num, title, html_id, data_panel=None):
    dp = f' data-panel="{data_panel}"' if data_panel else ''
    return f'''<div class="claim-section" id="{html_id}">
<div class="claim-header">
<span class="claim-num">{num}</span>
<h2{dp}>{title}</h2>
</div>'''

# ── Evidence boxes for each claim ──────────────────────────────────

ebox_claim1 = evidence_box(
    evidence_items=[
        'Categorias seletivas (filtro de desempenho + banca de leitura humana) superam as totalmente automáticas em ambas as dimensões de retorno.',
        'Critérios de bilheteria selecionam obras com ROI doméstico superior; critérios de festivais selecionam obras com ROI internacional superior.',
        'A separabilidade dos dados confirma que a metodologia do edital interfere na seleção e na trajetória dos filmes.',
        'A categoria "Pontuação Festivais e Roteiro" confirma capacidade de selecionar com viés à internacionalização.',
    ],
    limitation_items=[
        'A categoria "Automático Festivais" tem amostra pequena e destinação dos recursos difícil de rastrear integralmente.',
        'As categorias seletivas receberam volumes maiores de recursos, o que pode inflar o resultado comparativo.',
        'Número limitado de anos e mudanças frequentes nas regras dificultam isolar o efeito puro do critério de seleção.',
    ],
    conclusion_text='A metodologia de decisão de investimento importa: dados objetivos de desempenho pretérito, quando combinados com leitura humana qualificada, produzem resultados mensuravelmente superiores em ambas as dimensões.'
)

ebox_claim2 = evidence_box(
    evidence_items=[
        'Editais seletivos com distribuidora como proponente apresentam o resultado mais expressivo entre todas as categorias analisadas.',
        'Dois perfis distintos de distribuidoras: domésticas (Paris, Downtown, WMIX) com alta bilheteria e internacionais (Vitrine, Reserva Nacional, Elo, Embaúba) com alto ROI internacional.',
        'Produtoras como proponentes também obtiveram resultados relevantes (O Menino e o Mundo, As Boas Maneiras, Mato Seco Em Chamas).',
    ],
    limitation_items=[
        'Essa categoria recebeu o maior volume de recursos, o que pode inflar a comparação com categorias menos capitalizadas.',
        'Sequências de grandes sucessos (Minha Mãe é Uma Peça 2 e 3, etc.) migram para "Renúncia Pura", inflando o resultado daquele grupo e desidratando a comparação.',
        'Necessária investigação mais detalhada para isolar o efeito da distribuidora do efeito do volume.',
    ],
    conclusion_text='A experiência das distribuidoras na curadoria e comercialização de obras agrega valor ao processo de seleção, mas o efeito do volume de recursos precisa ser ponderado antes de uma conclusão definitiva.'
)

ebox_claim3 = evidence_box(
    evidence_items=[
        '<strong>Duplo Retorno</strong> (39 produtoras): 619 obras, R$ 2,2 bi de investimento para R$ 2,5 bi de renda estimada, presença em 33 países.',
        '<strong>Retorno Doméstico</strong>: empresas voltadas ao mercado interno, com alta bilheteria e internacionalização próxima de zero.',
        '<strong>Retorno Internacional</strong>: empresas sem resultado expressivo em bilheteria, mas que romperam a barreira da exportação com seleções em festivais e distribuição europeia.',
        '<strong>Baixo Retorno</strong> (174 empresas > R$ 5 mi): resultado esperado dentro do modelo de negócio do audiovisual.',
        '<strong>Pequeno Porte</strong> (927 empresas): mediana de menos de R$ 1 mi em 10 anos — ticket tão baixo que impede retorno significativo.',
    ],
    limitation_items=[
        'A classificação por cluster depende dos limiares escolhidos (R$ 2,5 mi de receita, ROI Intl > 0) — pequenas variações nos cortes produzem mudanças na distribuição.',
        'O ROI Internacional é um índice composto, não uma razão financeira direta — não mede receita internacional em valores monetários.',
        'Não é possível atribuir causalidade: produtoras "Duplo Retorno" podem ter características pré-existentes que explicam o desempenho.',
    ],
    conclusion_text='O FSA produziu talento e resultado mensurável em todas as faixas, mas não conseguiu responder com assertividade e tempestividade para que produtoras emergentes alcancem escala. As únicas que chegaram à segunda marcha já eram maiores antes do FSA.'
)

ebox_claim4 = evidence_box(
    evidence_items=[
        'Entre 2010 e 2023 o número de obras/ano cresceu 5,5× (72 → 397) e o de produtoras ativas/ano cresceu 5× (60 → 303).',
        'Coeficiente de Gini da distribuição do FSA por produtora = 0,634. Metade das produtoras recebe 11% do investimento; top 10% concentram 41%.',
        'RJ + SP concentram 60% dos recursos do FSA.',
        '43% das produtoras fizeram apenas 1 obra em uma década; mediana = 2 obras por produtora.',
        'Custo fixo estimado de uma produtora independente com equipe mínima: a partir de R$ 400 mil/ano.',
    ],
    limitation_items=[
        'A contagem de produtoras inclui CNPJs criados especificamente para contornar regras de concentração — a proliferação pode ser parcialmente artificial.',
        'As regras de regionalização (teto RJ+SP) estão próximas do máximo permitido pela legislação — o nível de concentração poderia ser ainda maior sem elas.',
        'O cálculo de custo fixo mínimo (R$ 400 mil) é uma estimativa; estruturas de custo variam significativamente.',
    ],
    conclusion_text='O fomento público induziu democratização no acesso, mas produziu fragmentação em vez de escala. Mais produtoras e mais obras com menos dinheiro por projeto resultam em menor capacidade de acumulação e inserção internacional.'
)

ebox_claim5 = evidence_box(
    evidence_items=[
        'Comparação entre chamadas com e sem políticas afirmativas mostra redução da desigualdade entre raça e gênero dos selecionados.',
        'Dados declaratórios do BRDE (via LAI) permitem comparar percentual de inscritos e selecionados por raça/gênero.',
    ],
    limitation_items=[
        'Análise preliminar — o autor planeja revisão mais aprofundada dos dados do BRDE.',
        'Dados são declaratórios e referem-se ao nível da empresa (sócios), não necessariamente da equipe ou do conteúdo.',
        'O modelo atual de cotas atribui a responsabilidade da diversidade apenas aos selecionados via cota, sem corresponsabilização dos demais.',
    ],
    conclusion_text='Políticas afirmativas por cota funcionam para reduzir a desigualdade de acesso, mas um modelo de corresponsabilização horizontal (como o BFI Diversity Standards) poderia ampliar o alcance da diversidade para todos os elos e camadas da produção.'
)


# ── Reassemble the article ─────────────────────────────────────────

# Remove the old h2 headings from extracted sections (we'll add new ones)
def strip_heading(section, tag='h2'):
    """Remove first h2/h3 from section."""
    return re.sub(rf'<{tag}[^>]*>.*?</{tag}>', '', section, count=1).strip()

# Remove h2::before from claim headers in CSS
# (already handled by .claim-header h2::before { display: none })

# The "two paragraphs about RLP" from comment 6
# These are the last 2 paragraphs of sec_estrutura
# Let's extract them
estrutura_lines = sec_estrutura.split('\n')
# Find the RLP paragraphs (lines containing "RLP" or "Ministério do Planejamento")
rlp_paras = []
other_estrutura_lines = []
found_rlp = False
for line in estrutura_lines:
    if 'Renda Líquida do produtor' in line or 'Ministério do Planejamento' in line:
        found_rlp = True
        rlp_paras.append(line)
    elif found_rlp and line.strip().startswith('<p>'):
        rlp_paras.append(line)
        found_rlp = False
    else:
        other_estrutura_lines.append(line)

rlp_content = '\n'.join(rlp_paras)
sec_estrutura_clean = '\n'.join(other_estrutura_lines)

# Build the article
article = f'''<article>

{hero_to_sobre}
{intro_paragraphs}
{metodologia}

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- PRÉ-CLAIM: Contexto — Renúncia fiscal e o papel das majors    -->
<!-- ═══════════════════════════════════════════════════════════════ -->
<div class="context-section" id="contexto-renuncia">
{sec_sucesso}
</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- CLAIM 1                                                        -->
<!-- ═══════════════════════════════════════════════════════════════ -->
{claim_header(1,
  'A metodologia de seleção determina o perfil de retorno: critérios de bilheteria favorecem desempenho doméstico e critérios de festivais favorecem desempenho internacional',
  'claim-criterios-selecao',
  '#categorias'
)}
{ebox_claim1}
{strip_heading(sec_chamadas)}
</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- CLAIM 2                                                        -->
<!-- ═══════════════════════════════════════════════════════════════ -->
{claim_header(2,
  'Distribuidoras como proponentes selecionam obras com melhor desempenho que produtoras',
  'claim-distribuidoras',
  '#categorias'
)}
{ebox_claim2}
{strip_heading(sec_distrib)}
</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- COLAPSÁVEL: Ampla concorrência                                 -->
<!-- ═══════════════════════════════════════════════════════════════ -->
<details class="collapsible" id="nota-ampla-concorrencia">
<summary>E as chamadas de ampla concorrência?</summary>
<div class="collapsible-body">
{strip_heading(sec_ampla)}
</div>
</details>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- CLAIM 3                                                        -->
<!-- ═══════════════════════════════════════════════════════════════ -->
{claim_header(3,
  'O perfil de retorno das produtoras revela cinco padrões distintos de desempenho — do duplo retorno à insuficiência estrutural',
  'claim-produtoras',
  '#produtoras'
)}
{ebox_claim3}
{strip_heading(sec_produtoras)}
</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- PREMISSA INTERPRETATIVA                                        -->
<!-- ═══════════════════════════════════════════════════════════════ -->
<div class="opinion-section" id="premissa-estrutura">
<div class="opinion-label">Premissa interpretativa</div>
{sec_estrutura_clean}

<div class="aside-note" id="nota-retorno-fsa">
<h3>Uma nota sobre a métrica de retorno do FSA</h3>
{rlp_content}
</div>

</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- CLAIM 4                                                        -->
<!-- ═══════════════════════════════════════════════════════════════ -->
{claim_header(4,
  'A proliferação de produtoras reduziu escala sem reduzir concentração',
  'claim-proliferacao',
  '#concentracao'
)}
{ebox_claim4}
{strip_heading(sec_proliferacao, 'h3')}
</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- CLAIM 5                                                        -->
<!-- ═══════════════════════════════════════════════════════════════ -->
{claim_header(5,
  'Políticas afirmativas por cota reduzem a desigualdade entre selecionados',
  'claim-diversidade',
  '#diversidade'
)}
{ebox_claim5}
{strip_heading(sec_diversidade)}
</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- ARGUMENTO: Instabilidade                                       -->
<!-- ═══════════════════════════════════════════════════════════════ -->
<div class="opinion-section" id="argumento-instabilidade">
<div class="opinion-label">Argumento</div>
{sec_instabilidade}
</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- NOTA: Financiamento                                            -->
<!-- ═══════════════════════════════════════════════════════════════ -->
<div class="aside-note" id="nota-financiamento">
{sec_financiamento}
</div>

<!-- ═══════════════════════════════════════════════════════════════ -->
<!-- PROPOSTAS                                                      -->
<!-- ═══════════════════════════════════════════════════════════════ -->
{sec_propostas}
</article>'''

# ── Update the TOC drawer ──────────────────────────────────────────
# Find and replace the TOC links
old_toc_start = head_and_css.find('<nav class="toc-drawer"')
old_toc_end = head_and_css.find('</nav>', old_toc_start) + len('</nav>')
if old_toc_start > 0:
    old_toc = head_and_css[old_toc_start:old_toc_end]
    new_toc = '''<nav class="toc-drawer" id="tocD">
  <div class="toc-label">Sumário</div>
  <a href="#sobre-este-texto">Sobre este texto</a>
  <a href="#contexto-renuncia">Contexto: Renúncia fiscal e majors</a>
  <a href="#claim-criterios-selecao">1. Critérios de seleção e retorno</a>
  <a href="#claim-distribuidoras">2. Distribuidoras vs produtoras</a>
  <a href="#nota-ampla-concorrencia">Nota: Ampla concorrência</a>
  <a href="#claim-produtoras">3. Perfil de retorno das produtoras</a>
  <a href="#premissa-estrutura">Premissa: A estrutura do fomento</a>
  <a href="#claim-proliferacao">4. Proliferação vs sustentabilidade</a>
  <a href="#claim-diversidade">5. Diversidade</a>
  <a href="#argumento-instabilidade">A instabilidade tem endereço</a>
  <a href="#nota-financiamento">Financiamento do FSA</a>
  <a href="#o-que-o-fsa-poderia-ser">O que o FSA poderia ser</a>

  <div class="toc-sep"></div>
  <a class="toc-ext" href="analise.html" target="_blank" rel="noopener">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6"/><path d="M10 14L21 3"/><path d="M18 13v6a2 2 0 01-2 2H5a2 2 0 01-2-2V8a2 2 0 012-2h6"/></svg> Análise Técnica
  </a>
  <a class="toc-ext" href="painel.html" target="_blank" rel="noopener">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg> Painel de Dados
  </a>
</nav>'''
    head_and_css = head_and_css[:old_toc_start] + new_toc + head_and_css[old_toc_end:]

# ── Assemble final HTML ────────────────────────────────────────────
final = head_and_css + '\n<div class="reading-area">\n' + article + '\n</div>\n' + footer

with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write(final)

print(f"\nOutput: {OUTPUT}")
print(f"Size: {len(final)} chars")
