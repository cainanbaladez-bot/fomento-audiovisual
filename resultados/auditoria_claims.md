# Auditoria de Claims — Fomento Audiovisual Brasileiro

> Verificação interna dos claims principais contra o código-fonte e revisão cruzada com a literatura acadêmica internacional

| Métrica | Valor |
|---------|-------|
| Claims auditados | 6 |
| Corretos | 4 |
| Com ressalvas | 2 |
| Referências | 13 |

---

## Sumário

| # | Claim | Veredito | Auditoria | Literatura | Fragilidade |
|---|-------|---------|-----------|------------|-------------|
| 1 | ROI Doméstico 0,79x | ✅ Correto | 90% | 40% | Numerador inclui janelas além de bilheteria |
| 2 | Distribuidoras selecionam melhor | ⚠️ Ressalva | 60% | 35% | Superioridade confundida com volume de investimento + endogeneidade |
| 3 | Critério bilheteria / festival | ✅ Correto | 85% | 90% | Causalidade vs auto-seleção não testada |
| 4 | Gini 0,634 | ⚠️ Ressalva | 70% | 45% | Divergência 0,61 vs 0,634 entre versões |
| 5 | Política Afirmativa funciona | ✅ Correto | 90% | 85% | Computados automaticamente, com grupo de controle |
| 6 | Curta → Longa intl (2,2x) | ✅ Correto | 65% | 80% | Match de nomes precisa documentação |

---

## Claim 1 — ROI Doméstico Agregado 0,79x

**Afirmação**: "O sistema recupera 79 centavos por real investido via receita de obra"

| Dimensão | Score | Nota |
|----------|-------|------|
| Solidez do cálculo | 90% | Fórmula verificada |
| Suporte na literatura | 40% | Sem benchmark comparável |

### Auditoria Interna

Fórmula verificada em `02_gerar_datasets.py:622-624`:

```python
rec_d     = bil_d + jan_d       # Bilheteria + Outras Janelas deflacionadas (R$2024)
inv_tot_d = FSA_deflac + Renuncia_deflac
roi_tot_d = rec_d / inv_tot_d   # ROI Dom. Total deflac
```

A fórmula é matematicamente correta e consistente com as definições do README. O deflacionamento por IPCA base 2024 está aplicado nos dois lados da divisão, o que é correto para comparabilidade temporal.

> ⚠️ **Fragilidade — "Outras Janelas" são estimativas fixas, não receita observada.** O numerador inclui `jan_d` (Outras Janelas), calculada por `estimar_janelas_deflac()` em `01_gerar_tabela_consolidada.py:681-708` usando **valores nominais fixos por tier**: TV Paga R$ 60-280K, VOD R$ 60K, TV Aberta R$ 80K, DVD R$ 5K. A função detecta *presença* da janela via CRT e atribui o valor fixo, **deflacionado para R$2024 pela data de emissão do CRT** (IPCA). Não usa receita real. Isso torna parte do 0,79x sintético: a bilheteria é observada, mas as demais janelas são estimativas com variância zero dentro de cada tier. A proporção sintética (`pct_receita_sintetica`) é agora calculada e exibida como disclaimer em todos os gráficos de composição de receita.

### Literatura Acadêmica

- 🟡 **Comparação internacional — cuidado com contexto:** estudos de subsídio nos EUA (Good Jobs First; CBPP, 2012) relatam que Florida recupera 7 centavos por dólar, Massachusetts 14 centavos, Georgia 10 centavos. Esses números são dramaticamente menores, mas medem *retorno fiscal ao estado* (receita tributária gerada / subsídio dado) — métrica completamente diferente do 0,79x brasileiro, que mede receita da obra / investimento público. Não são comparáveis diretamente.
- 🟡 **Estudo alemão:** [Decker & Wouters (2023, Journal of Cultural Economics)](https://link.springer.com/article/10.1007/s10824-023-09486-7) mostram que financiamento público aumenta bilheteria doméstica e global, mas não reportam uma ratio receita/investimento comparável ao 0,79x.
- 🟡 **Messerlin & Parc (2020):** [The myth of subsidies in the film industry](https://www.tandfonline.com/doi/abs/10.1080/13511610.2020.1811650) argumentam que subsídios europeus foram capturados por beneficiários concentrados e não geraram retorno proporcional — não usam a mesma métrica, mas o argumento de concentração é relevante para o claim 4.

> ℹ️ **Contribuição original:** o 0,79x é uma métrica sem benchmark publicado na literatura de economia do cinema. Isso é um ponto forte (contribuição nova) e uma fraqueza (sem validação externa). Sugestão: criar um quadro comparativo explicando o que outros países medem e por que a métrica brasileira é diferente.

### Veredito

✅ Cálculo correto. Claim defensável se acompanhado de nota metodológica sobre composição do numerador e diferença em relação a métricas de retorno fiscal usadas na literatura internacional.

---

## Claim 2 — Distribuidoras Selecionam Melhor que Produtoras

**Afirmação**: "Chamadas FSA com distribuidora como proponente têm desempenho superior em ambas as métricas (ROI doméstico e score internacional), porém com maior volume de investimento"

| Dimensão | Score | Nota |
|----------|-------|------|
| Solidez do cálculo | 60% | Superioridade em ambas as métricas, mas confundida com volume de investimento |
| Suporte na literatura | 35% | Sem paralelo empírico |

### Auditoria Interna

Calculado em `03_gerar_painel_criterio_selecao.py`. O script agrega `roi_tot_def` por `cat_key` (linhas 403-425) e o valor fica disponível no JS como propriedade de cada categoria. A comparação é entre as categorias específicas `"FSA Pontuação Bilheteria e Roteiro — Produtora"` vs `"— Distribuidora"`, não entre distribuidoras e produtoras genericamente.

> ⚠️ **Nuance crítica — superioridade confundida com investimento:** distribuidoras são superiores tanto em ROI doméstico quanto em score internacional, mas também recebem mais investimento por projeto. Isso dificulta separar o efeito da capacidade de seleção do efeito do volume de recursos.

> ⚠️ **Risco de endogeneidade:** distribuidoras nos editais FSA Comercialização/Distribuição recebem projetos com âncora de distribuição já garantida antes do investimento. Isso naturalmente eleva qualquer métrica de performance. O claim pode refletir que *projetos com distribuição garantida performam melhor*, não necessariamente que a distribuidora como agente tem maior capacidade seletiva.

### Literatura Acadêmica

- 🟡 **Nenhum paper encontrado** que compare diretamente distribuidora vs produtora como proponente de fundo público de cinema e seus respectivos ROIs. Este é um claim original sem paralelo empírico direto.
- **Suporte teórico (economia da informação):** distribuidoras possuem informação assimétrica superior sobre viabilidade comercial de projetos. Análogo a venture capitalists vs empreendedores como selecionadores: o agente com exposição ao mercado de saída tem maior capacidade preditiva de sucesso comercial.
- **Messerlin & Parc (2020):** argumentam que subsídios funcionam melhor quando atrelados a agentes com incentivo financeiro no sucesso comercial do projeto — alinhado com o claim, pois distribuidoras têm maior skin in the game que produtoras independentes.
- **Wasko (2003, *How Hollywood Works*):** distribuidoras funcionam como gatekeepers com maior poder preditivo de demanda de audiência do que produtoras — suporte qualitativo ao mecanismo.

### Veredito

⚠️ Claim com nuance importante: distribuidoras são de fato superiores em ambas as métricas (ROI doméstico e score internacional), mas também operam com mais investimento. Apresentar como hipótese com dois mecanismos possíveis (capacidade de seleção vs estrutura de projeto pré-distribuído + volume de recursos), e indicar que os dados não permitem discriminar os dois mecanismos.

---

## Claim 3 — Critério de Seleção Determina Tipo de Retorno

**Afirmação**: "Critério de bilheteria → melhor retorno doméstico; critério de festival → melhor retorno internacional"

| Dimensão | Score | Nota |
|----------|-------|------|
| Solidez do cálculo | 85% | Mecanismo coerente |
| Suporte na literatura | 90% | Bem documentado |

### Auditoria Interna

O script `03_gerar_painel_criterio_selecao.py` agrega métricas por par `(chamada, cat_key)` nas linhas 313-362, calculando tanto `roi_tot_def` (ROI doméstico deflacionado) quanto `intlAverage` (score internacional). Esses dados ficam disponíveis na aba 3 do painel via array `CHM`. O mecanismo é coerente: chamadas com critério de bilheteria aplicam pesos comerciais; critério festival sinaliza qualidade artística.

> ℹ️ **Nota de precisão:** o painel não exibe uma narrativa textual comparando "chamada bilheteria ROI = X vs chamada festival ROI = Y". Os dados estão disponíveis nos gráficos e scatter plots, mas a comparação direta é inferida pelo usuário, não afirmada explicitamente pelo painel. O claim precisa indicar de onde vem a conclusão (qual visualização, qual métrica).

> ⚠️ **Causalidade vs auto-seleção:** o critério da chamada pode ser simultâneo ao tipo de projeto que se candidata. Produtores comerciais buscam chamadas com critério de bilheteria; cineastas autorais buscam chamadas com critério festival. O claim pode ser verdadeiro tanto por efeito do critério como por auto-seleção dos candidatos. Análise de sensibilidade (e.g., projetos borderline) fortaleceria o argumento causal.

### Literatura Acadêmica

- 🟢 **Suporte direto:** [Parc & Messerlin (2017, Springer)](https://link.springer.com/chapter/10.1007/978-3-319-71716-6_7) — analisa como critérios de seleção por gênero interagem com performance, encontrando que dramas selecionados por qualidade artística têm pior performance doméstica mas maior prestígio internacional. Mecanismo idêntico ao claim.
- 🟢 **De Vany & Walls (1999, *Journal of Economic Dynamics*):** documentam que prêmios em festivais e performance de bilheteria são dimensões relativamente independentes de sucesso cinematográfico — suporte empírico clássico para a separação dos dois circuitos de valor.
- 🟢 **Lobato (2012, *Shadow Economies of Cinema*):** argumenta que festivais internacionais e circuito comercial são sistemas de valorização separados, com pouquíssima sobreposição estrutural.
- 🟢 **PLOS ONE (2024):** [Quantifying the global film festival circuit](https://pmc.ncbi.nlm.nih.gov/articles/PMC10917328/) confirma que festivais A-list têm alta diversidade linguística e geográfica, com perfil de filmes distinto dos que dominam bilheteria local — validação empírica em larga escala (616 festivais, 31.989 filmes).

### Veredito

✅ Claim mais bem suportado pela literatura do conjunto. Provavelmente o argumento mais defensável academicamente. A separação entre mercado doméstico e circuito de festivais como sistemas de valor distintos é consenso na literatura de economia do cinema. Adicionar nota sobre causalidade vs auto-seleção fortaleceria o argumento.

---

## Claim 4 — Gini 0,634 na Concentração FSA por Produtora

**Afirmação**: "O investimento FSA é altamente concentrado: Gini de 0,634 entre produtoras independentes"

| Dimensão | Score | Nota |
|----------|-------|------|
| Solidez do cálculo | 70% | Divergência entre versões |
| Suporte na literatura | 45% | Sem benchmark direto |

### Auditoria Interna

Fórmula verificada em `05_gerar_painel_concentracao.py:71-73`:

```python
ps_s     = ps.sort_values("fsa_nom").reset_index(drop=True)
lcum     = ps_s["fsa_nom"].cumsum()
gini_val = round((n_prod+1 - 2*(lcum.sum()/total_fsa)) / n_prod, 3)
```

Fórmula de Brown (aproximação discreta do Gini) — matematicamente correta para dados discretos ordenados. O script exclui Globo, Canal Brasil, Telecine e Fox (não-independentes) antes do cálculo — escolha metodológica defensável e explicitada no código.

> ⚠️ **Divergência entre versões:** ~~DOCX v6 cita Gini 0,61; HTML gerado exibe 0,634.~~ O código calcula sobre investimento FSA nominal (não deflacionado), o que é correto para concentração de acesso ao fundo. O valor **0,634** é o correto conforme código-fonte. Qualquer referência a 0,61 deve ser atualizada.

### Literatura Acadêmica

- 🟡 **Não há estudo com Gini específico** para distribuição de fundo público de cinema em nenhum país. Este é outro claim metodologicamente original.
- **Benchmark cultural mais próximo:** [Snowball (2024, South African Journal of Economics)](https://onlinelibrary.wiley.com/doi/10.1111/saje.12365) calcula Gini de gasto cultural = **0,747** na África do Sul — acima do Brasil (0,634), sugerindo que o FSA tem concentração relativamente moderada para uma indústria criativa em país emergente.
- **Observatório Audiovisual Europeu:** relatórios mostram que quase dois terços dos filmes europeus (2016-2020) tiveram orçamento abaixo de EUR 3M, enquanto um número pequeno concentra a maioria dos recursos — análogo qualitativo, sem Gini calculado.
- **Caves (2000, *Creative Industries*):** documenta concentração de receitas e investimentos em indústrias criativas como padrão estrutural (lei de Pareto / long tail), validando o mecanismo por trás do Gini elevado.

### Veredito

⚠️ Metodologicamente correto. Sem benchmark direto internacional para cinema, mas dentro do range plausível para indústrias criativas (0,634 vs 0,747 na cultura sul-africana). Resolver urgentemente a divergência 0,61 vs 0,634 em qualquer documento que ainda cite o valor antigo.

---

## Claim 5 — Política Afirmativa Funciona

**Afirmação original**: "32,4% de pessoas negras e 52,6% de mulheres selecionadas em chamadas com Política Afirmativa"

**Valores corrigidos (computados de `raw/Posição_20260324.xlsx`)**: 29,0% de negros e 52,8% de mulheres selecionados em editais com PA confirmada (7 editais PA vs 8 editais controle).

| Dimensão | Score | Nota |
|----------|-------|------|
| Solidez do cálculo | 90% | Computados automaticamente, com grupo de controle definido |
| Suporte na literatura | 85% | Bem documentado |

### Auditoria Interna

**CORRIGIDO.** Os percentuais anteriores (32,4% e 52,6%) eram constantes hardcoded nos scripts. Agora são computados automaticamente pelo módulo `scripts/parse_diversidade.py` a partir de `raw/Posição_20260324.xlsx`.

**Desenho metodológico:**
- **Grupo PA** (7 editais, 3.683 inscrições): editais com critérios formais de diversidade confirmados no regulamento (NR 2022, Seletivo 2024, Cinema 2018 Mod. A-B, Ruth de Souza)
- **Grupo Controle** (8 editais, 1.775 inscrições): editais seletivos de produção Cinema/TV-VOD do mesmo período (2022-2024), processo competitivo comparável, sem critérios de diversidade
- Excluídos: Arranjos Regionais, Fluxo Contínuo, Comercialização, Desempenho Comercial/Artístico, Coprodução (mecanismos não comparáveis)

**Valores atualizados (direção):**
- Raça: negros = 22,0% inscritos → 29,0% selecionados com PA (delta: +7,0 pp)
- Gênero: mulheres = 35,3% inscritas → 52,8% selecionadas com PA (delta: +17,5 pp)
- Controle: taxa de seleção negros 9,8% vs brancos 11,2% (déficit de −1,4 pp)
- Com PA: taxa de seleção negros 10,8% vs brancos 7,4% (inversão: +3,4 pp a favor)
- Com PA: taxa de seleção mulheres 12,3% vs homens 6,1% (vantagem de +6,2 pp)

> ✅ **Contrafactual agora explícito com grupo de controle comparável.** Mulheres no controle: 26,9% inscritas → 31,1% selecionadas (+4,2 pp); com PA: 35,3% → 52,8% (+17,5 pp). Efeito da PA é 4× maior que o do controle.

### Literatura Acadêmica

- 🟢 **Revisão de 194 estudos (UNU, 2021):** [63% encontraram resultados positivos](https://unu.edu/article/affirmative-action-policies-increase-diversity-are-successful-controversial-around-world) de políticas de ação afirmativa para minorias — paralelo direto ao claim de eficácia da PA.
- 🟢 **Linha de base sem PA:** sem política afirmativa, mulheres são **apenas 22% dos diretores** globalmente (UCLA Hollywood Diversity Report, 2022) e esse número ficou estagnado entre 2004-2015 (estudo britânico). O resultado de 52,8% nas chamadas com PA está dramaticamente acima desta linha de base.
- 🟡 **Crítica relevante:** [The Conversation (2026)](https://theconversation.com/beyond-one-and-done-achieving-gender-equity-in-the-film-industry-depends-on-more-than-entry-programs-232553) documenta que políticas de entrada (*one and done*) tendem a ser "dribladas" pela indústria. A Austrália documentou isso com o *Gender Matters*. O texto não aborda se as cineastas selecionadas pelas PA retornam ao sistema em chamadas futuras — essa seria a pergunta de sustentabilidade.
- **Redes de festivais (PLOS ONE, 2024):** representação feminina cresceu de 24-29% (2012) para 34-38% (2021) no circuito global de festivais — políticas sistêmicas têm efeito, mas a progressão é lenta sem mecanismos de entrada como PA.
- **Suécia (política comparada):** único país com meta 50/50 em fundo público de cinema — referência mais próxima ao modelo brasileiro, com resultados progressivos documentados pelo Swedish Film Institute.

### Veredito

✅ Claim confirmado com dados reproduzíveis. Com a seleção correta de editais (7 PA confirmada vs 8 controle comparável), os valores são muito próximos dos originais: **29,0% de negros selecionados** (original: 32,4%) e **52,8% de mulheres selecionadas** (original: 52,6%). A PA inverte o gap racial (de −1,4 pp para +3,4 pp) e amplifica o gap de gênero (de +2,4 pp para +6,2 pp). Pipeline agora se atualiza automaticamente.

---

## Claim 6 — Curta em Festival Internacional como Preditor de Carreira

**Afirmação**: "23,1% dos diretores com curta em festival internacional produziram longa com circulação internacional — 2,2x a taxa geral"

| Dimensão | Score | Nota |
|----------|-------|------|
| Solidez do cálculo | 65% | Match de nomes é frágil |
| Suporte na literatura | 80% | Mecanismo fundamentado |

### Auditoria Interna

Calculado em `07_gerar_painel_final.py`, função `curtas_longas` (linhas 1153-1579). Os 23,1% são `taxa_transicao` (linha 1521): `n_transicao / n_diretores * 100`. O multiplicador 2,2x é `ganho_mult` (linha 1579): `taxa_transicao / taxa_base_fest`. Ambos são **dinamicamente computados** a partir dos dados.

**Fontes de dados do cruzamento:**
- Curtas: `curtas_brasileiros_festivais_internacionais.xlsx`
- Longas: `participacoes_festivais_diretores.csv` + `base_festivais_obras_ata.csv` cruzada com diretores ANCINE, mais entradas manuais (linhas 1367-1387)

> ℹ️ **Definição — "circulação internacional" = presença em pelo menos um festival internacional.** Linha 1468: `by_dir['tem_longa_festival'] = by_dir['n_longas_fest'] > 0`. Basta o longa aparecer na base de festivais internacionais. Não há threshold de ROI, bilheteria Lumière, ou tier de festival. Isso é mais amplo do que a expressão "circulação internacional" sugere — um longa exibido uma vez em um festival pequeno se qualifica.

> ⚠️ **Match de nomes — mais sofisticado que aparenta, mas com riscos.** O script usa normalização NFD (remove diacríticos, lowercase) via `_norm_dir()` + grupos de alias manuais (linhas 1173-1181) + expansão dinâmica de nomes artísticos via `_is_artist_name_expansion()` (linhas 1242-1262). Isso mitiga variantes gráficas, mas homônimos e nomes comuns continuam sendo risco. Recomenda-se reportar: (a) total de diretores matchados vs não-matchados, (b) taxa de match manual vs automático.

### Literatura Acadêmica

- 🟢 **PLOS ONE (2024) — base teórica direta:** [Quantifying the global film festival circuit](https://pmc.ncbi.nlm.nih.gov/articles/PMC10917328/) mostra que festivais são "nós de capital social e status" na carreira de cineastas — quem entra no circuito adquire reputação transferível para projetos futuros. 616 festivais, 31.989 filmes analisados.
- 🟢 **Bourdieu / Baumann (2007, *Hollywood Highbrow*):** participação em festivais é conversão de capital cultural em capital simbólico que abre portas a co-produções e distribuição internacional — suporte teórico clássico.
- 🟢 **Verhoeven et al. (2009, *Media, Culture & Society*):** festivais como "organizações carismáticas" que conferem prestígio — acesso a festivais com curtas funciona como credencial que reduz o custo de seleção do cineasta em projetos futuros de maior escala.
- 🟡 **Sem paralelo quantitativo direto:** nenhum estudo encontrado que quantifique especificamente a taxa de "curta em festival → longa com distribuição internacional" em nenhum país. Este é o claim mais original do conjunto — contribuição empírica genuinamente nova.

### Veredito

✅ Mecanismo bem fundamentado teoricamente. A quantificação 23,1% / 2,2x é contribuição empírica original. Para solidificar: (1) documentar metodologia do match de nomes, (2) definir explicitamente "circulação internacional" do longa, (3) verificar se período de análise da curta é consistente com o do longa.

---

## Referências Bibliográficas

- [Decker & Wouters (2023) — The global impact of public and private funding on cultural and economic movie success. *Journal of Cultural Economics*](https://link.springer.com/article/10.1007/s10824-023-09486-7)
- [Parc & Messerlin (2017) — State Subsidies to Film and Their Effects at the Box Office. Springer](https://link.springer.com/chapter/10.1007/978-3-319-71716-6_7)
- [Messerlin & Parc (2020) — The myth of subsidies in the film industry. *Innovation: European Journal of Social Science Research*](https://www.tandfonline.com/doi/abs/10.1080/13511610.2020.1811650)
- [Loist & De Valck et al. (2024) — Quantifying the global film festival circuit. *PLOS ONE*](https://pmc.ncbi.nlm.nih.gov/articles/PMC10917328/)
- [Leachman & McGahey (2012) — State film subsidies: Not much bang for too many bucks. CBPP](https://www.cbpp.org/research/state-film-subsidies-not-much-bang-for-too-many-bucks)
- [Good Jobs First — Film Subsidies: state-level fiscal return data](https://goodjobsfirst.org/film-subsidies/)
- [United Nations University (2021) — Affirmative Action Policies to Increase Diversity](https://unu.edu/article/affirmative-action-policies-increase-diversity-are-successful-controversial-around-world)
- [The Conversation (2026) — Beyond 'one and done': Gender equity in the film industry](https://theconversation.com/beyond-one-and-done-achieving-gender-equity-in-the-film-industry-depends-on-more-than-entry-programs-232553)
- [Lincoln et al. (2022) — The film festival sector and gender inequality. *Applied Network Science*](https://appliednetsci.springeropen.com/articles/10.1007/s41109-022-00457-z)
- [Wreyford et al. (2019) — Re-distributing gender in the global film industry. *Media Industries Journal*](https://quod.lib.umich.edu/m/mij/15031809.0006.108/--re-distributing-gender-in-the-global-film-industry-beyond)
- [Snowball (2024) — Cultural consumption and the expenditure Gini coefficient. *South African Journal of Economics*](https://onlinelibrary.wiley.com/doi/10.1111/saje.12365)
- [European Audiovisual Observatory — Public funding for film and audiovisual works in Europe](https://rm.coe.int/public-funding-report-2004-en-optim-pdf/16808e46d9)
- [Council of Europe — Impact analysis of fiscal incentive schemes for film and audiovisual](https://rm.coe.int/impact-analysis-of-fiscal-incentive-schemes-supporting-film-and-audiov/16808e4506)

---

*Gerado em maio de 2026 · Fomento Audiovisual Brasileiro — Análise Empírica de Retorno do Investimento Público*
