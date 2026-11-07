"""Compara os números de uma nova execução do notebook com os resultados publicados no TCC.

Uso:
    python conferir_resultados.py                       # compara TCC_Reproducao_Completa_executado.ipynb
    python conferir_resultados.py outro.ipynb --tolerancia 0.002
    python conferir_resultados.py --gerar-referencia    # (autor) regrava resultados_esperados.json

Todos os números impressos por cada célula de código são extraídos e comparados, na ordem, com os de
resultados_esperados.json. Por padrão a comparação é exata (tolerância relativa de 1e-9). Com outras versões
de scikit-learn/XGBoost, use --tolerancia (diferença absoluta máxima aceita) para verificar se as diferenças
ficam na terceira casa decimal.
"""
import argparse, hashlib, json, re, sys

NUM = re.compile(r'-?\d+(?:\.\d+)?(?:e[-+]?\d+)?')
IGNORAR = ('Versões utilizadas', 'Figuras geradas', 'Arquivo figuras_tcc.zip')

def extrair(caminho):
    nb = json.load(open(caminho, encoding='utf-8'))
    saida = []
    for c in nb['cells']:
        if c['cell_type'] != 'code':
            continue
        src = ''.join(c['source'])
        texto = ''
        for o in c.get('outputs', []):
            t = o.get('text') or o.get('data', {}).get('text/plain', '')
            texto += ''.join(t) if isinstance(t, list) else t
        if any(m in texto for m in IGNORAR):
            continue
        ident = hashlib.sha1(src.encode()).hexdigest()[:10]
        saida.append({'celula': ident, 'inicio': src.splitlines()[0][:70] if src else '',
                      'numeros': [float(x) for x in NUM.findall(texto)]})
    return saida

ap = argparse.ArgumentParser()
ap.add_argument('notebook', nargs='?', default='TCC_Reproducao_Completa_executado.ipynb')
ap.add_argument('--tolerancia', type=float, default=None)
ap.add_argument('--gerar-referencia', action='store_true')
a = ap.parse_args()

if a.gerar_referencia:
    json.dump(extrair('TCC_Reproducao_Completa.ipynb'), open('resultados_esperados.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('resultados_esperados.json regravado a partir de TCC_Reproducao_Completa.ipynb'); sys.exit(0)

ref = {c['celula']: c for c in json.load(open('resultados_esperados.json', encoding='utf-8'))}
novo = {c['celula']: c for c in extrair(a.notebook)}
total = diverg = 0; maior = 0.0
for k, r in ref.items():
    if k not in novo:
        print(f'[FALTA] célula não encontrada: {r["inicio"]}'); diverg += 1; continue
    x, y = r['numeros'], novo[k]['numeros']
    if len(x) != len(y):
        print(f'[DIFERENTE] {r["inicio"]}: {len(x)} números esperados, {len(y)} obtidos'); diverg += 1; continue
    ruins = []
    for i, (u, v) in enumerate(zip(x, y)):
        total += 1
        if u == v:                      # inclui valores não finitos idênticos
            continue
        d = abs(u - v); maior = max(maior, d) if d == d else maior
        ok = d <= a.tolerancia if a.tolerancia is not None else d <= 1e-9 * max(1.0, abs(u))
        if not ok:
            ruins.append((u, v))
    if ruins:
        diverg += 1
        print(f'[DIFERENTE] {r["inicio"]}: {len(ruins)} valor(es), ex.: esperado {ruins[0][0]} obtido {ruins[0][1]}')
print(f'\n{total} números comparados em {len(ref)} células; maior diferença absoluta: {maior:.3g}')
print('RESULTADOS REPRODUZIDOS.' if diverg == 0 else f'{diverg} célula(s) com diferença.')
sys.exit(1 if diverg else 0)
