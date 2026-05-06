# Fomento Audiovisual Brasileiro — Dados e Evidencias

Analise empirica do retorno do investimento publico em audiovisual no Brasil (FSA e renuncia fiscal), com base em dados abertos da ANCINE, Lumiere, IMDB e outras fontes.

**[Leia o texto completo com dados interativos](https://cainanbaladez-bot.github.io/fomento-audiovisual/)**

---

## O que voce encontra aqui

### Tres documentos integrados

| Documento | Descricao | Link direto |
|-----------|-----------|-------------|
| **Uma politica de fomento baseada em evidencias** | Texto argumentativo guiado por dados. Cada topico com dados abre o painel interativo relevante. | [politica.html](https://cainanbaladez-bot.github.io/fomento-audiovisual/politica.html) |
| **Analise de Dados** | Documento tecnico com metodologia, definicoes, ressalvas e resultados por secao. | [analise.html](https://cainanbaladez-bot.github.io/fomento-audiovisual/analise.html) |
| **Painel Interativo** | Dashboard com graficos, tabelas e filtros cobrindo retorno domestico, internacional, categorias FSA, produtoras, concentracao, diversidade e soft power. | [painel.html](https://cainanbaladez-bot.github.io/fomento-audiovisual/painel.html) |

### Navegacao entre os tres

O texto principal (`politica.html`) funciona como porta de entrada. Os titulos de topico que possuem dados associados sao clicaveis e abrem um painel lateral com o dashboard na secao relevante. A barra superior tem links diretos para a analise tecnica e o painel completo, que abrem em nova aba.

---

## Principais achados

- **2.990 obras** na base analisada, com **R$ 2,38 bilhoes** em receita domestica estimada
- **ROI domestico agregado de 0,79x** — o sistema recupera 79 centavos por real investido via bilheteria
- **Distribuidoras selecionam melhor que produtoras** como proponentes nos editais FSA
- **Chamadas com criterio de bilheteria** tem melhor resultado domestico; **criterio de festival** tem melhor resultado internacional
- **39 produtoras de Duplo Retorno** — resultado domestico e internacional simultaneo
- **Gini de 0,634** na distribuicao do investimento FSA por produtora
- **Politica afirmativa funciona**: 32,4% de pessoas negras e 52,6% de mulheres selecionadas em chamadas com PA
- **23,1% dos diretores** com curta em festival internacional produziram longa com circulacao internacional — 2,2x a taxa da base geral

---

## Estrutura do repositorio

```
fomento-audiovisual/
|
|-- docs/                  # GitHub Pages (site publicado)
|   |-- index.html         # Redireciona para politica.html
|   |-- politica.html      # Texto argumentativo + painel modal
|   |-- analise.html       # Documento tecnico
|   |-- painel.html        # Dashboard interativo
|
|-- raw/                   # Dados brutos (ANCINE, Lumiere, IMDB)
|   |-- obras-nao-pub-brasileiras-csv/   # Obras por ano (2002-2026)
|   |-- projetos-fsa.csv                 # Projetos FSA
|   |-- projetos-com-renuncia-fiscal.csv # Renuncia fiscal
|   |-- bilheteria_brasileira_consolidada.xlsx
|   |-- lumiere_search.xlsx              # Lumiere/CNC Europa
|   |-- lumiere_vod_search.xlsx          # VOD Europa
|   +-- ...
|
|-- dados/                 # Datasets processados
|   |-- critica_obras.csv
|   |-- prestigio_diretores.csv
|   |-- participacoes_festivais_diretores.csv
|   |-- perfil_festivais_diretores.csv
|   |-- imdb_enrichment.csv
|   +-- ...
|
|-- tabelas_apoio/         # Deflator IPCA, clusters, preco ingresso
|-- scripts/               # Pipeline Python (01 a 08)
|-- output_final/          # HTMLs originais (nomes com acento)
|-- resultados/            # Secoes estaticas (curtas, diversidade, soft power)
|-- requirements.txt       # Dependencias Python
+-- run_pipeline.sh        # Executa todos os scripts em sequencia
```

---

## Pipeline de dados

Os scripts devem ser executados em ordem. Cada um le a saida do anterior.

| Script | O que faz |
|--------|-----------|
| `01_gerar_tabela_consolidada.py` | Cruza obras, FSA, renuncia fiscal, bilheteria, festivais, IMDB, Lumiere |
| `02_gerar_datasets.py` | Gera datasets analiticos: clusters de produtoras, ROI por categoria, metricas |
| `03_gerar_painel_criterio_selecao.py` | HTML do painel de criterios de selecao (categorias FSA) |
| `04_gerar_painel_produtoras.py` | HTML do painel de produtoras e clusters |
| `05_gerar_painel_concentracao.py` | HTML do painel de concentracao (Gini, Lorenz, tiers) |
| `06_gerar_painel_comparativo.py` | HTML do painel comparativo final |
| `07_gerar_painel_final.py` | Monta o HTML completo do dashboard unificando todos os paineis |
| `08_enriquecer_doc_analise.py` | Gera o documento de analise tecnica |

Para rodar tudo:

```bash
pip install -r requirements.txt
bash run_pipeline.sh
```

---

## Fontes de dados

| Fonte | Conteudo | Acesso |
|-------|----------|--------|
| ANCINE / SAV | Obras, produtores, diretores, agentes economicos | [dados.gov.br](https://dados.gov.br) |
| ANCINE / FSA | Projetos FSA, investimentos, categorias | [fsa.ancine.gov.br](https://fsa.ancine.gov.br) |
| Renuncia Fiscal | Projetos via Art. 1o, 1oA, 3o, 3oA | [ancine.gov.br](https://www.ancine.gov.br) |
| SADIS / Bilheteria | Bilheteria consolidada por filme | [ancine.gov.br](https://www.ancine.gov.br) |
| Lumiere / CNC | Circulacao em salas europeias | [lumiere.obs.coe.int](https://lumiere.obs.coe.int) |
| Lumiere VOD | Disponibilidade em VOD Europa | [lumiere.obs.coe.int](https://lumiere.obs.coe.int) |
| IMDB | Notas, generos, metadata | [imdb.com](https://www.imdb.com) |
| Festivais | Selecoes e premiacoes internacionais | Curadoria propria |
| Critica | Notas AdoroCinema, Metacritic | Coleta propria |
| OpenAlex | Citacoes academicas | [openalex.org](https://openalex.org) |

---

## Definicoes analiticas principais

- **ROI Domestico**: Bilheteria / (FSA + Renuncia)
- **Receita Total Deflacionada**: Bilheteria + outras janelas, corrigida pelo IPCA base 2024
- **Duplo Retorno**: Receita total >= R$ 2,5M E ROI Internacional > 0
- **Retorno Domestico**: Receita total > R$ 2,5M E ROI Internacional = 0
- **Retorno Internacional**: ROI Internacional > 0 E Receita total < R$ 2,5M
- **Fomento Baixo Retorno**: Investimento FSA > R$ 5M E ROI Internacional = 0
- **Pequeno Porte**: Bilheteria <= R$ 500K (residual)

---

## Licenca

Os dados brutos sao publicos (Lei de Acesso a Informacao). O codigo e a analise sao de uso livre para fins de pesquisa e debate sobre politica publica.
