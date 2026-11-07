# Previsão de Ingressantes e Análise da Evasão por Curso no IFCE

Material de reprodução do Trabalho de Conclusão de Curso de **Cleiton Barbosa Valente** (Bacharelado em Ciência da
Computação, IFCE *campus* Aracati), orientado pelo Prof. Alexandro Lima Damasceno.

Este repositório permite reproduzir, a partir de uma cópia congelada dos dados, todas as tabelas, figuras e números
do capítulo de Resultados do trabalho.

## Conteúdo

| Arquivo | Para que serve |
|---|---|
| `TCC_Reproducao_Completa.ipynb` | Notebook único com todo o *pipeline*: dados, *features*, modelos, validações, análises e figuras. Está salvo **com as saídas** da execução de referência. |
| `Banco de Dados.xlsx` | Cópia congelada dos dados analisados (única entrada do notebook). |
| `SHA256SUMS` | *Hash* SHA-256 dos arquivos de entrada. |
| `requirements.txt` | Versões das bibliotecas usadas (Python 3.11). `requirements-lock.txt` traz o ambiente completo. |
| `verificar_base.py` | Confere *hash*, estrutura, chaves duplicadas e totais da base. |
| `executar.py` | Executa o notebook do início ao fim, sem intervenção manual. |
| `conferir_resultados.py` | Compara os números da nova execução com os publicados (`resultados_esperados.json`). |
| `PROTOCOLO_COLETA.md` | Como os dados foram coletados e conferidos em cada fonte. |
| `DICIONARIO_DE_DADOS.md` | Estrutura da planilha e definição de cada indicador e *feature*. |
| `figuras/` | Figuras geradas na execução de referência. |
| `scripts/gerar_notebook.py` | Script que gera o notebook (código-fonte das células). |

## Como reproduzir

Requer Python 3.11.

```bash
python3.11 -m venv .venv
source .venv/bin/activate            # no Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m ipykernel install --user --name python3

python verificar_base.py             # 1. valida a base de entrada
python executar.py                   # 2. executa o notebook (5 a 15 minutos, conforme a máquina)
python conferir_resultados.py        # 3. compara com os resultados publicados
```

O passo 3 deve terminar com `RESULTADOS REPRODUZIDOS.`. A comparação padrão é exata. Com outras versões de
scikit-learn ou XGBoost, os resultados de *Random Forest* e XGBoost podem variar na terceira casa decimal; nesse caso,
use `python conferir_resultados.py --tolerancia 0.002`. A Regressão Linear, as referências ingênuas e as análises
descritivas não dependem da versão.

**Google Colab:** abra o notebook, use *Ambiente de execução → Executar tudo* e envie `Banco de Dados.xlsx` quando
solicitado. O Colab usa versões próprias das bibliotecas, então pequenas diferenças são esperadas.

## Configuração experimental

- **Dados:** cursos de graduação do IFCE, 2016 a 2025 (2026 excluído por ser ano em andamento), 139 ofertas em 31 *campi*;
  modelagem com as 113 ofertas que têm ao menos 4 anos de ingressantes (633 linhas).
- **Divisão:** treino 2016–2023 e teste 2024–2025; *walk-forward* com janela expansiva e teste em 2021, 2022, 2023, 2024 e 2025.
- **Modelos:** Regressão Linear (mínimos quadrados); *Random Forest* (300 árvores); XGBoost (300 árvores, profundidade 4,
  taxa de aprendizado 0,05). Semente 42. *Target encoding* do curso com suavização 10, ajustado só no treino.
- **Referências ingênuas:** persistência (ano anterior), média móvel de 3 anos e média histórica da oferta.
- **Incerteza:** *bootstrap* por oferta (2.000 réplicas) para diferenças de MAE.
- **Horizontes:** teste retroativo de previsão recursiva de 1 a 4 anos, com origens de 2019 a 2024.

## Verificação em ambiente limpo

A execução completa foi refeita em um ambiente virtual novo, criado apenas a partir de `requirements.txt`, e todos os
números conferiram exatamente com a execução de referência (`conferir_resultados.py` sem tolerância).

## Dados e licença

Os dados são públicos e agregados (sem informação pessoal), obtidos do painel *IFCE em Números*, dos Anuários Estatísticos
do IFCE, dos Termos de Adesão do IFCE ao Sisu e da classificação Cine Brasil do INEP. O painel pode ser atualizado e
alterar valores históricos; por isso, a referência para reprodução é a cópia congelada identificada em `SHA256SUMS`.
