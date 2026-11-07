# Protocolo de coleta e conferência dos dados

## 1. Indicadores acadêmicos — painel *IFCE em Números*

- **Versão do painel:** "Atualizado em: 07/01/2026 01:24:47".
- **Filtro:** Nível de ensino = Graduação.
- **Indicadores:** Ingressantes, Matrícula (evadida), Matrícula (em curso), Matrícula (concluída) e Matrícula (trancada).
- **Procedimento:** para cada indicador e cada *campus*, foram registrados os valores anuais (2016 a 2026) de cada
  combinação de curso, modalidade do curso (bacharelado, licenciatura ou tecnologia) e modalidade de ensino
  (presencial ou EaD).
- **Conferência:** os valores consolidados foram conferidos manualmente contra o painel antes da modelagem. Para
  Matrícula (trancada), a soma dos valores dos cursos de cada *campus* e ano foi comparada ao total do *campus*
  exibido pelo painel, sem nenhuma divergência.
- **Ofertas homônimas:** quando o painel agrega cursos de mesmo nome com graus diferentes (por exemplo, Educação Física
  em Juazeiro do Norte, licenciatura e bacharelado), os valores foram obtidos separadamente com o filtro de modalidade
  do curso.
- **Ausências:** células sem valor no painel foram registradas como "-" (sem dado).
- **Matrícula (trancada):** coletada em etapa posterior, com o mesmo procedimento, e usada só no teste exploratório.

## 2. Vagas ofertadas

- **2016 a 2021:** tabela "Total de vagas ofertadas/procura por cursos" dos Anuários Estatísticos do IFCE
  (edições 2017 a 2022, cada uma referente ao ano-base anterior), apenas cursos de graduação.
- **2022 a 2025:** Termos de Adesão do IFCE ao Sisu, que registram apenas as vagas ofertadas por esse sistema.
- **Conferência:** amostras dos anos inconsistentes (2017 e 2020) foram conferidas nos próprios Anuários e
  correspondem à fonte (por exemplo, 2.576 vagas de Hotelaria em Fortaleza no ano-base 2017). Esses anos são
  excluídos das análises de ocupação; a auditoria ano a ano está na seção 5.0 do notebook.

## 3. Área do conhecimento

Área geral da classificação Cine Brasil (INEP), atribuída a cada combinação de curso e grau acadêmico e conferida
contra o manual oficial da classificação.

## 4. Congelamento

A planilha consolidada `Banco de Dados.xlsx` é a cópia congelada usada em todas as análises. Seu *hash* SHA-256 está em
`SHA256SUMS` e é verificado por `verificar_base.py` e pelo próprio notebook (seção 1.1).
