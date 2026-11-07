"""Valida o arquivo de entrada antes da execução: hash, estrutura, chaves e totais.

Uso:  python verificar_base.py
"""
import hashlib
import sys
import openpyxl

ARQUIVO = 'Banco de Dados.xlsx'
SHA256 = 'caf62626b22fe77ae2cb968098cf2c7665d48eca34c7e903db961d5dcc9afb52'
TIPOS = {'Ingressantes', 'Evasão', 'Matrícula em curso', 'Matrícula concluída', 'Vagas ofertadas', 'Matrícula trancada'}
# Totais 2016-2025 esperados por tipo de registro (somando todas as ofertas)
TOTAIS = {'Ingressantes': 54866, 'Evasão': 39305, 'Matrícula em curso': 175596}

falhas = []
def checar(nome, cond, detalhe=''):
    print(f'[{"ok" if cond else "FALHA"}] {nome} {detalhe}')
    if not cond:
        falhas.append(nome)

sha = hashlib.sha256(open(ARQUIVO, 'rb').read()).hexdigest()
checar('hash SHA-256', sha == SHA256, sha)

ws = openpyxl.load_workbook(ARQUIVO, data_only=True)['Planilha1']
cab = [ws.cell(row=1, column=c).value for c in range(1, 20)]
linhas = [[ws.cell(row=r, column=c).value for c in range(1, 20)] for r in range(2, ws.max_row + 1)]
linhas = [l for l in linhas if any(v is not None for v in l)]
checar('colunas de anos 2016-2026', [str(c) for c in cab[6:17]] == [str(a) for a in range(2016, 2027)], str(cab[6:17]))
checar('tipos de registro', {l[5] for l in linhas} == TIPOS, str(sorted({l[5] for l in linhas})))
chaves = [tuple(l[1:6]) for l in linhas]
checar('chaves (campus, curso, modalidades, tipo) sem duplicata', len(chaves) == len(set(chaves)), f'{len(chaves)} linhas')
ofertas = {tuple(l[1:5]) for l in linhas}
checar('número de ofertas', len(ofertas) == 139, str(len(ofertas)))
checar('número de campi', len({l[1] for l in linhas}) == 31, str(len({l[1] for l in linhas})))
for tipo, esperado in TOTAIS.items():
    soma = sum(float(v) for l in linhas if l[5] == tipo for v in l[6:16] if v not in (None, '-'))
    checar(f'total 2016-2025 de "{tipo}"', round(soma) == esperado, f'{soma:.0f} (esperado {esperado})')
valores = [v for l in linhas for v in l[6:17] if v not in (None, '-')]
checar('valores numéricos e não negativos', all(isinstance(v, (int, float)) and v >= 0 for v in valores), f'{len(valores)} valores')

print('\nBase válida.' if not falhas else f'\n{len(falhas)} verificação(ões) falharam: {falhas}')
sys.exit(1 if falhas else 0)
