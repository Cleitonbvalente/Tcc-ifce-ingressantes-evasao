"""Executa o notebook do início ao fim, sem intervenção manual.

Uso:  python executar.py
Gera TCC_Reproducao_Completa_executado.ipynb e a pasta figuras/.
"""
import time
import nbformat
from nbclient import NotebookClient

ENTRADA = 'TCC_Reproducao_Completa.ipynb'
SAIDA = 'TCC_Reproducao_Completa_executado.ipynb'

inicio = time.time()
nb = nbformat.read(ENTRADA, as_version=4)
# limpa saídas antigas para garantir uma execução nova
for c in nb.cells:
    if c.cell_type == 'code':
        c.outputs, c.execution_count = [], None
NotebookClient(nb, timeout=3600, kernel_name='python3', allow_errors=False).execute()
nbformat.write(nb, SAIDA)
print(f'Notebook executado em {time.time() - inicio:.0f} s -> {SAIDA}')
