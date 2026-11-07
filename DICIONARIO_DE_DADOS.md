# Dicionário de dados

## Planilha `Banco de Dados.xlsx` (aba `Planilha1`, formato largo)

Uma linha por oferta e tipo de registro; uma coluna por ano.

| Coluna | Descrição |
|---|---|
| `Nível` | Nível de ensino (sempre Graduação). |
| `Cidade` | *Campus* do IFCE. |
| `Curso` | Nome do curso. |
| `Modalidade (Licenciatura/Bacharelado)` | Modalidade do curso: Bacharelado, Licenciatura ou Tecnologia. |
| `Modalidade(Presencial/EAD)` | Modalidade de ensino: Presencial ou EAD. |
| `Tipo` | Indicador: Ingressantes, Evasão, Matrícula em curso, Matrícula concluída, Vagas ofertadas ou Matrícula trancada. |
| `2016` a `2026` | Valor do indicador no ano ("-" = sem dado). 2026 não é usado (ano em andamento). |
| `CODIGO_AREA_GERAL_CINE`, `AREA_CONHECIMENTO_CINE` | Código e área geral da classificação Cine Brasil. |

As abas `Mapa_CINE` e `Auditoria_CINE` documentam a atribuição da área Cine a cada curso e grau; não são lidas pelo notebook.

**Oferta** = combinação de *campus*, curso, modalidade do curso e modalidade de ensino (139 ofertas).

## Indicadores

| Indicador | Definição (painel *IFCE em Números*) |
|---|---|
| Ingressantes | Matrículas de alunos que ingressaram no curso no ano. |
| Evasão — Matrícula (evadida) | Matrículas de alunos que perderam o vínculo antes da conclusão, no ano. |
| Matrícula (em curso) — matrículas ativas | Matrículas ativas em pelo menos um dia no ano. |
| Matrícula (concluída) — concluintes | Matrículas de alunos que concluíram todos os componentes no ano. |
| Vagas ofertadas | Anuários Estatísticos (2016–2021) e Termos de Adesão ao Sisu (2022–2025). |
| Taxa operacional de evasão | Soma das evasões dividida pela soma das matrículas ativas do conjunto considerado. |
| Taxa de ocupação | Ingressantes divididos pelas vagas ofertadas, por oferta e ano. |

## *Features* (todas calculadas só com anos anteriores ao ano previsto *t*)

| *Feature* | Cálculo | Anos usados |
|---|---|---|
| `ingressantes_lag1`, `ingressantes_lag2` | valor do ano anterior / de dois anos antes | *t*-1 / *t*-2 |
| `ingressantes_ma3` | média dos três anos anteriores | *t*-1 a *t*-3 |
| `ingressantes_cresc_pct` | variação percentual entre *t*-2 e *t*-1 | *t*-1, *t*-2 |
| `evasao_*`, `matriculas_ativas_*`, `concluintes_*`, `vagas_ofertadas_*` | `_lag1` = ano anterior; `_ma3` = média dos três anos anteriores | *t*-1 a *t*-3 |
| `taxa_evasao_hist` | evasões acumuladas / matrículas ativas acumuladas | todos até *t*-1 |
| `taxa_conclusao_hist` | concluintes acumulados / ingressantes acumulados | todos até *t*-1 |
| `taxa_ocupacao_hist` | ingressantes acumulados / vagas acumuladas | todos até *t*-1 |
| `curso_enc` | *target encoding* do curso (suavização 10), ajustado só no treino | treino |
| `ano`, `campus_*`, `mod_curso_*`, `mod_ensino_*`, `area_cine_*` | ano e variáveis categóricas (*one-hot*) | — |

A dependência temporal de cada *feature* é verificada automaticamente na seção 7.1 do notebook.
