# Peer Review: Uma política de fomento baseada em evidências v6

## Escopo e critérios
- Verifiquei consistência entre o DOCX v6, o HTML gerado pelo pipeline e a auditoria local dos números.
- Avaliei quatro eixos: rastreabilidade dos números, consistência interna, completude editorial e robustez das inferências.
- Foquei nos trechos que sustentam as teses centrais do texto: retorno doméstico, internacionalização, clusters de produtoras, concentração e capacidade de carga do sistema.

## Veredito
- Revisão maior necessária.
- O texto tem uma tese forte e uma base analítica plausível, mas ainda mistura números de versões diferentes, contém placeholders inacabados e usa alguns denominadores/metodologias que não estão explicitados com clareza.

## Principais achados

- [Alta] Há divergência numérica material entre o DOCX v6 e o painel/HTML de referência nos argumentos centrais.
- O DOCX afirma que as categorias com distribuidora como proponente têm ROI internacional de 65% superior, com valores 3,01 vs. 1,82 ([`resultados/_docx_v6_text.txt:37`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L37)).
- O HTML convertido a partir do próprio DOCX já traz 2,87 vs. 1,65, o que implica cerca de 74%, não 65% ([`output_final/Uma política de fomento baseada em evidências_v6.html:205`](C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6.html#L205)).
- Isso não é um arredondamento trivial; altera a magnitude do efeito que sustenta a recomendação de política.

- [Alta] O texto usa contagens de clusters que não batem com a versão já consolidada do painel.
- O DOCX v6 diz que o cluster Retorno Doméstico tem 74 produtoras e que o Duplo Retorno tem 65 produtoras ([`resultados/_docx_v6_text.txt:50`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L50), [`resultados/_docx_v6_text.txt:52`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L52)).
- No HTML de referência, os mesmos blocos aparecem com 54 e 35 produtoras ([`output_final/Uma política de fomento baseada em evidências_v6.html:223`](C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6.html#L223), [`output_final/Uma política de fomento baseada em evidências_v6.html:226`](C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6.html#L226)).
- A conclusão política muda se o universo é 74/65 ou 54/35; isso precisa ser resolvido antes de submissão.

- [Alta] Há pelo menos uma inconsistência metodológica central entre as versões do texto: o Gini da concentração por produtora.
- O DOCX v6 cita Gini 0,61, 11% para metade das produtoras e 47% para o topo 10% ([`resultados/_docx_v6_text.txt:69`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L69)).
- O HTML da mesma versão já embute Gini 0,634 no KPI e também menciona 0,724 em narrativa, ou seja, há conflito interno no próprio artefato convertido ([`output_final/Uma política de fomento baseada em evidências_v6.html:253`](C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6.html#L253), [`output_final/Uma política de fomento baseada em evidências_v6.html:255`](C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6.html#L255)).
- A auditoria local indica que 0,634 é o valor mais consistente com o painel analítico, não 0,61.

- [Alta] Há placeholders editoriais que não podem permanecer em uma versão final.
- O DOCX v6 contém trechos como “xxB”, “xx de São Paulo” e “xx do Rio de Janeiro” no bloco sobre Retorno Doméstico ([`resultados/_docx_v6_text.txt:50`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L50)).
- Isso indica que a seção foi deixada com marcadores de preenchimento e não passou por revisão final.

- [Média] Alguns números são apresentados sem deixar claro o denominador ou a janela analítica.
- O texto fala em “27 obras responsáveis por 75% da renda” e em “94% da renda” quando se soma renúncia e FSA ([`resultados/_docx_v6_text.txt:19`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L19)).
- O painel de referência não traz essa decomposição como KPI explícito, então a afirmação pode estar correta, mas não está suficientemente rastreável para um leitor externo ([`resultados/_auditoria_dados_docx.txt:32`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_auditoria_dados_docx.txt#L32)).
- Para um texto de política pública, a falta de rastreabilidade nessa cifra é um ponto fraco.

- [Média] O argumento “683 beneficiários” versus “1.075 produtoras” sugere que o texto mistura unidades analíticas diferentes sem sinalizar a mudança.
- O DOCX v6 usa 683 beneficiários e 150 empresas de duplo retorno ([`resultados/_docx_v6_text.txt:76`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L76)).
- O HTML consolidado trabalha com 1.075 produtoras e 35 no cluster Duplo Retorno ([`output_final/Uma política de fomento baseada em evidências_v6.html:226`](C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6.html#L226), [`output_final/Uma política de fomento baseada em evidências_v6.html:227`](C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6.html#L227)).
- Sem explicação explícita de universo, o leitor não consegue saber se a mudança é de população, filtro, versão ou método.

- [Média] A seção de conclusões mistura diagnóstico forte com poucas verificações de robustez.
- O texto recomenda contratos plurianuais, joint ventures e um SUAT Internacional com base em diferenças observadas entre categorias ([`resultados/_docx_v6_text.txt:97`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L97), [`resultados/_docx_v6_text.txt:98`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L98), [`resultados/_docx_v6_text.txt:99`](C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt#L99)).
- Falta mostrar sensibilidade das conclusões a limites de corte, a diferentes definições de ROI e à exclusão/inclusão de renúncia fiscal.

## Pontos fortes
- A tese é clara e relevante.
- O texto tenta ligar dados, metodologia e recomendação institucional, o que é raro e valioso.
- A estrutura narrativa é boa: sinergia FSA-renúncia, critérios de seleção, produtores e capacidade de carga do sistema.
- O HTML convertido indica que o pipeline consegue gerar um artefato auditável, o que é uma base boa para um documento mais sólido.

## Recomendação objetiva de revisão
- Corrigir todos os números que mudaram entre versões.
- Eliminar placeholders e qualquer marca de rascunho.
- Explicitar para cada KPI: universo, período, denominador e fórmula.
- Harmonizar o texto com a versão consolidada do painel e congelar um snapshot de dados para evitar drift.
- Incluir uma seção curta de limitações e robustez, especialmente para Gini, clusters e ROI internacional.

## Julgamento final
- Aceitável como rascunho analítico.
- Ainda não está pronto para circulação final sem uma revisão numérica e editorial séria.

## Sources
- `C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_docx_v6_text.txt`
- `C:\Users\INTEL\Desktop\fomento-audiovisual\resultados\_auditoria_dados_docx.txt`
- `C:\Users\INTEL\Desktop\fomento-audiovisual\output_final\Uma política de fomento baseada em evidências_v6.html`
