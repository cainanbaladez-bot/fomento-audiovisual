# Análise de Dados do Fomento Público ao Audiovisual Brasileiro

_FSA Cinema — Painel Integrado de Análise · valores deflacionados para R$ 2024_

---

## Metodologia

**Pergunta central:** editais FSA que usaram histórico de festivais internacionais como critério de seleção produziram obras com maior alcance internacional do que editais que usaram critérios comerciais (bilheteria, market share)? E como os mecanismos automáticos e de renúncia fiscal se comparam a ambos em retorno doméstico?

**Universo de obras:** PRODECINE 2014–2023 · ANCINE (SPO/SAM) · FSA/BRDE · SALIC/MinC · Lumière/CNC

**Fontes:** projetos-fsa.csv (BRDE/FSA) · projetos-com-renuncia-fiscal.csv (SALIC/MinC) · por_filme_ano.csv (ANCINE 2014+) · lumiere_search.xlsx (bilheteria EU) · lumiere_vod_search.xlsx · festivais_consolidado.csv (Ata BRDE/FSA 2024 + pesquisa bibliográfica)

### Métricas de retorno

- **ROI Doméstico (média ponderada):** bilheteria deflacionada proporcional ao investimento ÷ investimento deflacionado do mecanismo. Mede a eficiência de cada mecanismo isolado, ponderada pelo capital alocado.
- **ROI Internacional (média incondicional):** score composto 0–100 calculado por obra e depois medido por categoria. Inclui zeros — obras sem presença internacional reduzem a média. Compõe: 70% score de festivais (Oscar, Cannes, Berlim, Veneza, BAFTA, Globo de Ouro) + 20% admissões EU (Lumière/CNC) + 10% VOD internacional.
- **ROI Total deflacionado:** receita total (bilheteria + janelas estimadas) ÷ investimento total (FSA + renúncia), deflacionado em IPCA base dezembro 2024.
- **Deduplicação:** obras com múltiplas chamadas na mesma categoria aparecem uma única vez para métricas de obra-nível (bilheteria, festivais, VOD). O ROI usa todas as entradas para refletir o investimento real de cada chamada.

### Limitações

- Obras com ano de produção > 2023 excluídas — ciclo de vida incompleto.
- Bilheteria pré-2014 estimada por título (menor precisão).
- Festivais cobre apenas as principais premiações internacionais.
- Renúncia fiscal exclui obras sem FSA confirmado para evitar superestimação do ROI.
- Categorias com menos de 5 obras excluídas do painel por insuficiência amostral.
- A análise mede associação observada, não causalidade.

---

## 1. Visão Geral do Sistema

### 1.1 Mecanismos de financiamento

O sistema de fomento ao audiovisual brasileiro opera por dois mecanismos principais: o **FSA (Fundo Setorial do Audiovisual)**, que é investimento direto via editais do BRDE, e a **renúncia fiscal** (Lei do Audiovisual, Art. 3/3-A/39), em que o Estado abre mão de receita tributária para que empresas apliquem impostos devidos em projetos audiovisuais. Diferente do FSA, a renúncia não tem seleção ANCINE/BRDE — a aprovação é administrativa.

No recorte PRODECINE 2014–2023 com bilheteria registrada (661 obras), os quatro grupos de financiamento mostram perfis radicalmente distintos:

| Grupo | N obras | Inv. Total | Bilheteria | ROI Dom. (pond.) | ROI Intl. (média) | % com Intl. |
|---|---:|---:|---:|---:|---:|---:|
| Renúncia Pura | 270 | R$ 0,68 bi | R$ 1,47 bi | 2,22x | 1,3 | 3,0% |
| FSA + Renúncia — Ren. Maj. | 91 | R$ 0,59 bi | R$ 0,28 bi | 0,50x | 2,1 | 6,6% |
| FSA + Renúncia — FSA Maj. | 112 | R$ 0,55 bi | R$ 0,22 bi | 0,43x | 4,2 | 13,4% |
| FSA Puro | 188 | R$ 0,33 bi | R$ 0,01 bi | 0,08x | 2,9 | 6,9% |

_Valores deflacionados para R$ 2024. ROI Dom. = bilheteria proporcional ao mecanismo ÷ investimento. ROI Intl. = score composto 0–100._

**Leitura dos dados:** há uma troca estrutural entre retorno doméstico e alcance internacional. A Renúncia Pura financia obras de maior apelo comercial (ROI 2,22x) mas com baixo alcance internacional (apenas 3% das obras têm presença externa). O FSA Puro, ao contrário, seleciona projetos com maior presença internacional (6,9%, ROI Intl 2,9) mas retorno doméstico irrisório (0,08x). Os grupos mistos ocupam posições intermediárias, com o grupo de maioria FSA se destacando em internacionalização (ROI Intl 4,2, 13,4% com presença externa).

### 1.2 Retorno doméstico

A renúncia fiscal sustenta o retorno doméstico do sistema. Esse resultado reflete o processo de seleção: projetos que captam via renúncia precisam de patrocinadores privados, o que funciona como filtro de viabilidade de mercado antes mesmo do início da produção. O FSA puro, sem esse filtro, financia projetos de menor tração comercial.

Nos grupos mistos, o aporte de renúncia como maioria (Ren. Maj.) se associa a ROI doméstico mais elevado (0,50x) do que o aporte com maioria FSA (0,43x) — sugerindo que o componente de renúncia carrega correlação positiva com retorno de bilheteria mesmo dentro dos grupos mistos.

### 1.3 Retorno internacional

O padrão se inverte para internacionalização. O FSA Puro (2,9) e o grupo de maioria FSA (4,2) superam a Renúncia Pura (1,3) no score internacional. Isso é consistente com a hipótese de que os editais FSA — especialmente os seletivos — favorecem obras com ambi ção artística e potencial de circulação em festivais, métricas mais fracamente correlacionadas com bilheteria doméstica.

A proporção de obras com presença internacional é mais alta no grupo de maioria FSA (13,4%) do que em qualquer outro grupo — incluindo o FSA Puro (6,9%) — o que sugere que a combinação de critérios FSA com algum componente de renúncia seleciona obras com perfil mais equilibrado.

---

## 2. Categorias das Chamadas FSA

O FSA opera por distintas modalidades de edital, com critérios de seleção diferentes. A análise das 8 categorias identificadas mostra como o critério de entrada molda tanto o retorno doméstico quanto o alcance internacional das obras financiadas.

| Categoria | Obras | Inv. Total (R$mi) | Bilheteria (R$mi) | Rec. Total (R$mi) | ROI Dom. Total | ROI Intl. Médio | % c/ Festival |
|---|---:|---:|---:|---:|---:|---:|---:|
| FSA Complementação | 44 | 163,5 | 92,2 | 97,4 | 0,56x | 1,13 | 23,8% |
| FSA Pontuação Bilheteria — Distribuidora | 101 | 567,9 | 269,4 | 283,5 | 0,47x | 3,91 | 25,7% |
| FSA Comercialização / Distribuição | 21 | 54,1 | 13,0 | 15,3 | 0,24x | 3,30 | 46,2% |
| FSA Pontuação Bilheteria — Produtora | 102 | 450,8 | 102,6 | 117,0 | 0,23x | 1,07 | 12,5% |
| FSA Automático Bilheteria | 34 | 139,4 | 29,2 | 32,2 | 0,21x | 1,18 | 13,3% |
| FSA Automático Festivais | 21 | 21,9 | 2,6 | 3,7 | 0,12x | 7,61 | 64,3% |
| FSA Coprodução Internacional | 28 | 18,1 | 0,9 | 1,8 | 0,05x | 4,27 | 43,5% |
| FSA Pontuação Festivais e Roteiro | 75 | 191,7 | 3,8 | 9,8 | 0,02x | 3,65 | 33,2% |

_Valores deflacionados para R$ 2024. ROI Dom. Total = receita total ÷ investimento total (FSA + renúncia)._

### Padrões por critério de seleção

**Critério comercial (bilheteria):** as duas maiores categorias em volume de investimento — Distribuidora (R$ 567,9mi) e Produtora (R$ 450,8mi) — selecionam projetos por histórico de desempenho comercial. A modalidade via Distribuidora entrega ROI doméstico mais alto (0,47x vs. 0,23x) e ROI internacional muito superior (3,91 vs. 1,07). A diferença aponta para uma vantagem estrutural quando a distribuidora lidera a candidatura: ela assume a responsabilidade pela comercialização, traz expertise de lançamento e tem incentivo direto no sucesso da obra.

**Critério artístico (festivais):** a modalidade Pontuação Festivais e Roteiro (PRODECINE 03–05) pontua o histórico artístico internacional da produtora na Fase 1. O resultado é paradoxal: o menor ROI doméstico do sistema (0,02x) combinado com relevância internacional moderada (3,65). A explicação provável é que o critério artístico seleciona obras de linguagem mais autoral, com menor apelo de mercado mas com outras formas de valor cultural.

**Critério automático por festivais (SUAT):** o FSA Automático Festivais, que recompensa obras já reconhecidas em festivais nacionais e internacionais, registra o maior ROI Internacional médio do sistema (7,61) e a maior proporção de obras com passagem por festivais (64,3%). É a categoria mais eficiente para internacionalização, mas com volume de investimento muito menor (R$ 21,9mi).

**Complementação e Comercialização:** são categorias de etapa final — investem em obras já produzidas. A Complementação aparece como a categoria de melhor ROI total (0,56x), sugerindo captura de projetos maduros ou melhor estruturados na etapa de distribuição. A Comercialização/Distribuição tem ROI doméstico moderado (0,24x) mas alta proporção de obras com festivais (46,2%), reflexo de um portfólio de obras de nicho com circulação especializada.

**Coprodução Internacional:** com apenas 28 obras e menor investimento (R$ 18,1mi), apresenta ROI doméstico muito baixo (0,05x) mas alta presença em festivais (43,5%) e ROI internacional relevante (4,27). O exigido parceiro estrangeiro parece garantir acesso a circuitos internacionais, embora não se traduza em bilheteria doméstica.

### Síntese: qual metodologia priorizar por objetivo?

| Objetivo de política | Categoria recomendada |
|---|---|
| Maximizar retorno doméstico | FSA Complementação / Distribuidora |
| Maximizar alcance internacional | FSA Automático Festivais |
| Equilíbrio doméstico + internacional | FSA Pontuação Bilheteria — Distribuidora |
| Internacionalização via coprodução | FSA Coprodução Internacional |

---

## 3. Produtoras Independentes

O recorte de produtoras usa 929 entidades da base consolidada, sendo 751 classificadas como brasileiras independentes. A análise de clusters agrupa as produtoras por perfil de retorno combinando ROI doméstico, ROI internacional e volume de investimento.

### 3.1 Distribuição por cluster

| Cluster | Produtoras | Obras | FSA (R$mi) | Inv. Total (R$mi) | Receita (R$mi) | ROI Dom. | ROI Intl. Médio |
|---|---:|---:|---:|---:|---:|---:|---:|
| Duplo Retorno | 39 | 619 | 483,9 | 2.355,5 | 2.710,4 | 1,15x | 24,43 |
| Retorno Doméstico | 51 | 461 | 413,8 | 1.974,2 | 2.477,6 | 1,25x | 3,72 |
| Retorno Internacional | 52 | 233 | 156,7 | 445,6 | 48,1 | 0,11x | 22,97 |
| Fomento Baixo Retorno | 174 | 1.291 | 1.070,6 | 3.143,2 | 353,2 | 0,11x | 1,78 |
| Pequeno Porte | 927 | 1.808 | 908,2 | 1.685,9 | 203,0 | 0,12x | 0,47 |

_Valores deflacionados para R$ 2024._

**Critérios de classificação:**

- **Duplo Retorno:** receita total ≥ R$ 2,5M e ROI Internacional > 0 — combina retorno doméstico forte com presença externa.
- **Retorno Doméstico:** receita total > R$ 2,5M e ROI Internacional = 0 — foco no mercado doméstico sem circulação internacional.
- **Retorno Internacional:** ROI Internacional > 0 e receita total < R$ 2,5M — alcance externo sem retorno doméstico forte.
- **Fomento Baixo Retorno:** investimento FSA > R$ 5M e ROI Internacional = 0 — alto capital público alocado com baixo retorno em ambas as dimensões.
- **Pequeno Porte:** categoria residual — produtoras com bilheteria ≤ R$ 500K.

### 3.2 Leitura dos clusters

O cluster **Duplo Retorno** concentra 39 produtoras — uma minoria do universo — que obtêm ROI doméstico agregado de 1,15x e ROI Internacional máximo médio de 24,43. Essas produtoras acumularam um portfólio capaz de competir em mercados doméstico e externo simultaneamente.

O cluster **Retorno Doméstico** reúne 51 produtoras e entrega o maior ROI doméstico agregado (1,25x), acima do Duplo Retorno, mas com ROI Internacional inferior ao dos clusters internacionalizados. São produtoras especializadas em obras de amplo apelo doméstico — filmes de comédia, franquias, animação familiar — que dependem menos dos circuitos internacionais de festivais.

O cluster **Retorno Internacional** (52 produtoras) sustenta ROI Internacional qualificado alto (ROI Internacional máximo médio de 22,97) com investimento total inferior a R$ 500mi. Esse é o segmento de cinema de autor com forte presença em festivais internacionais mas baixo retorno doméstico.

O **Fomento Baixo Retorno** concentra 174 produtoras e o maior volume de investimento total (R$ 3.143,2mi), com retorno fraco em ambas as dimensões (ROI Dom. 0,11x, ROI Intl. 1,78). Esse cluster representa o risco de alocação de capital público sem mecanismos de seleção adequados.

O **Pequeno Porte** (927 produtoras) concentra a maioria das empresas com investimento e receita baixos — é o segmento estruturalmente mais vulnerável, como detalhado na seção de concentração.

### 3.3 Matriz ROI Doméstico × ROI Internacional

O scatter de produtoras revela quatro quadrantes: (1) alto retorno doméstico e alta internacionalização — o espaço do Duplo Retorno, ocupado por pouquíssimas produtoras; (2) alto doméstico, baixo internacional — Retorno Doméstico; (3) baixo doméstico, alta internacionalização — Retorno Internacional; (4) baixo em ambos — a maior parte do universo (Fomento Baixo Retorno e Pequeno Porte).

---

## 4. Concentração do Investimento

A concentração do capital FSA é um traço estrutural do sistema, que persiste independentemente do recorte escolhido.

### 4.1 Indicadores de concentração

| Métrica | Valor |
|---|---:|
| Gini FSA por produtora (universo total) | **0,634** |
| Gini FSA por obra | **0,480** |
| Top 10 produtoras — % do FSA total | **15,2%** |
| Top 50 produtoras — % do FSA total | **35,8%** |
| Bottom 50% das produtoras — % do FSA | **≈ 4%** |
| Produtoras na zona de risco (< ticket mínimo) | **762 (71%)** |

O coeficiente de Gini de 0,634 por produtora é significativamente superior ao Gini de 0,480 calculado por obra individual. Isso indica que a desigualdade se amplifica quando a unidade de análise passa de projetos para empresas: as grandes produtoras captam mais projetos **e** recebem valores maiores por projeto — dupla vantagem estrutural.

### 4.2 Sustentabilidade operacional

Uma produtora independente brasileira com equipe mínima (2–3 profissionais fixos) incorre em custos anuais de aproximadamente R$ 400k–700k apenas em pessoal e overhead — sem contar o desenvolvimento de projetos. Com ciclos de produção de 18–24 meses por longa-metragem, um ticket FSA inferior a R$ 500k/ano torna inviável manter estrutura permanente.

**71% das produtoras** (762 empresas) estão nessa zona de risco: operam de chamada em chamada, sem capacidade de pipeline, formação de equipe ou desenvolvimento continuado. Os 30% de produtoras com ticket acima de R$ 700k/ano concentram **79,7% dos recursos** totais do FSA.

A curva acumulada mostra que 70% das produtoras ficam abaixo de R$ 1M/ano de ticket médio. A inflexão ocorre em torno de R$ 2M/ano, patamar acima do qual se encontram apenas 10% das produtoras — que, por sua vez, concentram a grande maioria dos recursos.

### 4.3 Distribuição por tier

| Tier | N produtoras | % FSA | Ticket med. obs. (R$k/ano) | % abaixo R$ 500k |
|---|---:|---:|---:|---:|
| A — Mega | ≈ 10 | 15,2% | Alto | Baixa |
| B — Grande | ≈ 40 | 20,6% | Médio-alto | Baixa |
| C — Médio | ≈ 100 | 29,4% | Médio | Moderada |
| D — Pequeno | ≈ 200 | 29,1% | Baixo | Alta |
| E — Micro (301+) | 784 | 5,7% | R$ 499k | 50%+ |

O Tier E reúne 784 produtoras com mediana de ticket de R$ 499k/ano — exatamente no limiar de insustentabilidade — e 50% delas abaixo desse patamar. Essas produtoras representam 72% do total de empresas com FSA, mas operam essencialmente como veículos de projeto único. A política que distribui recursos muito dispersamente cria uma ilusão de diversidade setorial ao custo de impedir a formação de capacidade institucional.

---

## 5. Curtas como Porta de Entrada

A análise de curtas-metragens busca identificar se o sucesso em festivais internacionais no formato curta é um preditor de carreira cinematográfica com alcance internacional.

**Base:** 56 curtas brasileiros com seleção ou premiação em festivais internacionais mapeados (dados/curtas_brasileiros_festivais_internacionais.xlsx), período 2004–2025.

### 5.1 Indicadores principais

| Indicador | Valor |
|---|---|
| Curtas mapeados | 56 |
| Seleções em festivais internacionais | 57 |
| Diretores creditados | 65 |
| Premiações | 16 (28,1% das seleções) |
| Diretores com longa internacional posterior | 15 (23,1%) |

### 5.2 Festivais

Os sete festivais internacionais com maior presença de curtas brasileiros no recorte são: Cannes, Berlinale, Annecy, Rotterdam, Locarno, Veneza e Clermont-Ferrand. A concentração em festivais de prestígio máximo é um dado relevante: indica que os curtas que chegam a esses circuitos têm alta visibilidade e geram networking que facilita produções futuras.

### 5.3 Associação curta → longa

23,1% dos diretores com curta em festival internacional identificado neste recorte têm longa-metragem com presença internacional posterior na base do projeto. A taxa é uma associação observada — não implica causalidade, mas é consistente com a hipótese de que os festivais de curta funcionam como mecanismo de credenciamento e network.

A linha do tempo de trajetórias mostra que, quando existe a transição, o intervalo entre o curta em festival e o longa com presença internacional costuma ser de 3 a 7 anos — período compatível com o ciclo de desenvolvimento e produção de um longa independente.

---

## 6. Diversidade e Políticas Afirmativas

### 6.1 Escopo

Análise de 99 editais FSA (BRDE/FSA 2015–2025) com informações sobre raça e gênero declarados por função (Diretor, Roteirista, Produtor). 24 dos 99 editais tinham alguma forma de política afirmativa.

### 6.2 Impacto das políticas afirmativas

| Indicador | Valor |
|---|---|
| Editais com política afirmativa | 24 de 99 |
| Ganho racial com PA | +9,6 pp (negros: 22,8% inscritos → 32,4% selecionados) |
| Ganho de gênero com PA | +15,6 pp (mulheres: 37,0% inscritas → 52,6% selecionadas) |
| Maior impacto isolado | +30 pp (mulheres diretoras TV-VOD NR 2022) |

### 6.3 Achado central

Nos editais **sem política afirmativa**, negros têm taxa de seleção menor que brancos (26,8% vs. 29,9%) — o sistema reproduz o déficit histórico. Nos editais **com política afirmativa**, a relação se inverte: negros passam a ter taxa **63% maior** que brancos (14,8% vs. 9,1%). Para gênero, mulheres atingem quase o **dobro da taxa** dos homens em editais com PA (15,0% vs. 7,9%).

Esses dados indicam que sem mecanismos ativos de equidade, o sistema de seleção tende a reproduzir as desigualdades estruturais do setor. As políticas afirmativas são eficazes — mas cobrem apenas 24% dos editais analisados.

### 6.4 Distribuição global (todos os 99 editais)

Considerando a totalidade dos editais, a distribuição de raça e gênero entre inscritos e selecionados mostra uma sub-representação persistente de pessoas negras e mulheres nas posições de direção — déficit que as PAs corrigem parcialmente nos editais onde são aplicadas, mas que persiste estruturalmente no conjunto.

---

## 7. Soft Power do Cinema Brasileiro

### 7.1 Crítica cinematográfica

**Base:** 1.461 filmes de cinema com FSA e/ou renúncia fiscal, com cobertura de crítica em pelo menos uma fonte (5 fontes de crítica utilizadas), período 2012–2023.

| Indicador | Valor |
|---|---|
| Filmes com crítica | 1.461 |
| Índice médio (2+ fontes) | **3,89 / 5** |
| Cobertura do universo financiado | **53,1%** |
| Filmes com 2+ fontes (consenso) | 780 |

O índice médio de 3,89/5 posiciona o cinema brasileiro financiado por fomento público em patamar favorável de avaliação crítica, considerando que obras sem nenhuma repercussão cultural raramente atraem duas ou mais fontes.

### 7.2 Correlação crítica × ROI Internacional

A análise de dispersão entre índice de crítica e ROI Internacional mostra correlação positiva: filmes com avaliação crítica mais alta tendem a ter maior alcance internacional. Isso é consistente com o papel da crítica especializada como gatekeeping para circulação em festivais e mercados externos.

O coeficiente de correlação é positivo mas moderado — há filmes com alta crítica e baixo alcance internacional (cinema de autor com circulação restrita) e filmes com baixa crítica e presença em alguns mercados (produções comerciais com distribuição internacional de nicho).

### 7.3 Citação acadêmica

A dimensão de citação acadêmica (OpenAlex, h-index por obra) mede o impacto cultural de longo prazo do cinema brasileiro — sua capacidade de gerar reflexão e produção de conhecimento. Obras com maior score de festivais internacionais tendem também a maior citação acadêmica, o que sugere que o reconhecimento artístico e o impacto cultural se alimentam mutuamente.

### 7.4 Top filmes por índice de crítica

Entre os filmes com 2+ fontes de crítica, os dez mais bem avaliados ilustram a amplitude do fomento público: obras de renúncia fiscal pura (Tropa de Elite, Tropa de Elite 2), obras FSA de autor (Noite de Fogo — Automático Festivais) e filmes de nicho (Arábia, Slam: Voz de Levante). O índice máximo registrado é 4,45/5, obtido por dois filmes.

---

## 8. Síntese

### O que os dados mostram

1. **Troca estrutural doméstico/internacional:** a renúncia fiscal financia o retorno doméstico (ROI 2,22x); o FSA seletivo financia a internacionalização. São instrumentos com lógicas diferentes — compará-los por uma única métrica de retorno distorce a análise.

2. **Critério de seleção importa:** entre as categorias FSA, a modalidade via Distribuidora entrega o melhor equilíbrio (ROI Dom. 0,47x + ROI Intl. 3,91). O Automático por Festivais lidera em internacionalização (7,61). A Complementação aparece como linha eficiente por capturar obras maduras.

3. **Concentração é uma característica estrutural:** Gini de 0,634 por produtora, com 71% das empresas abaixo do limiar de sustentabilidade. A dispersão do capital reproduz dependência de cada edital sem formação de capacidade institucional.

4. **Políticas afirmativas funcionam — mas são exceção:** 24 de 99 editais têm PA. Sem PA, o sistema reproduz desigualdades raciais e de gênero. Com PA, os grupos sub-representados alcançam ou superam as taxas dos grupos majoritários.

5. **Curtas como credenciamento:** 23,1% dos diretores de curtas em festivais internacionais têm longa com presença internacional posterior — taxa consistente com o papel de credenciamento dos festivais.

6. **Qualidade crítica é mensurável:** índice médio de 3,89/5 com 53% do universo financiado coberto. A correlação com ROI Internacional confirma que reconhecimento crítico e alcance externo se reforçam.

### O que os dados não mostram

- **Causalidade:** todos os resultados são associações observadas. O FSA não "causa" internacionalização — seleciona projetos com perfil de internacionalização.
- **Receita de janelas:** a estimativa de outras janelas (streaming, TV) é proxy, não observação direta de caixa.
- **Impacto cultural amplo:** festivais e crítica captam parte do valor — impacto social, formação de público e preservação de identidade cultural não estão mensurados.
- **Contrafactual:** o que teria acontecido com as obras sem fomento público é desconhecido.

---

_Documento gerado a partir das bases consolidadas do repositório (resultados/datasets/). Painel interativo correspondente: Análise do Retorno do Fomento Público ao Audiovisual Brasileiro (FSA — Renúncia Fiscal)\_v2.html_
