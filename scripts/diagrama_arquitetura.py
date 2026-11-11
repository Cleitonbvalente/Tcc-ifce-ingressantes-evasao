import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import sys
saida = sys.argv[1] if len(sys.argv) > 1 else 'Diagrama-de-arquitetura.png'

# Figura desenhada no tamanho de impressão (16 cm de largura) para que as letras saiam com 7,5 a 9 pt.
fig, ax = plt.subplots(figsize=(6.3, 3.55))
ax.set_xlim(0, 14); ax.set_ylim(0.35, 7.25); ax.axis('off')
AZUL, AZUL_CLARO, CINZA, CINZA_CLARO, LARANJA_CLARO = '#2c7fb8', '#e8f1f8', '#595959', '#f2f2f2', '#fdf0e6'
TIT, TXT = 7.6, 6.9

def caixa(x, y, w, h, titulo, texto, cor_borda=AZUL, cor_fundo=AZUL_CLARO, etapa=None, ft=TIT):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.02,rounding_size=0.12', fc=cor_fundo, ec=cor_borda, lw=1.0))
    yy = y + h - 0.16
    if etapa:
        ax.text(x + w / 2, yy, etapa, ha='center', va='top', fontsize=TXT, color=cor_borda, fontweight='bold'); yy -= 0.40
    ax.text(x + w / 2, yy, titulo, ha='center', va='top', fontsize=ft, fontweight='bold', color='#1a1a1a', linespacing=1.1)
    nl = titulo.count('\n') + 1
    ax.text(x + w / 2, yy - 0.36 * nl - 0.14, texto, ha='center', va='top', fontsize=TXT, color='#333333', linespacing=1.25)

def seta(x1, y1, x2, y2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle='-|>', mutation_scale=9, lw=1.0, color=CINZA))

ax.text(1.4, 7.0, 'Fontes de dados', ha='center', fontsize=TIT, fontweight='bold', color=CINZA)
fontes = [('IFCE em Números', 'ingressantes, evasão,\nmatrículas, concluintes'),
          ('Anuários Estatísticos', 'vagas 2016–2021'),
          ('Termos de Adesão', 'Sisu: vagas 2022–2025'),
          ('Cine Brasil (Inep)', 'área do conhecimento')]
for (t, d), y in zip(fontes, [5.45, 4.0, 2.55, 1.1]):
    caixa(0.1, y, 2.6, 1.3, t, d, cor_borda='#8c8c8c', cor_fundo=CINZA_CLARO, ft=7.0)
    seta(2.75, y + 0.65, 3.15, 3.9)

caixa(3.2, 2.65, 2.3, 2.5, 'Coleta e\nintegração', 'planilha única\nconferida com\no painel', etapa='Etapa 1')
seta(5.55, 3.9, 5.95, 3.9)
caixa(6.0, 2.15, 2.6, 3.5, 'Pré-processamento\ne atributos', 'ofertas com\n≥ 4 anos\ndefasagens e\nmédias móveis\n(shift(1))\ntarget encoding\ne one-hot', etapa='Etapas 2 e 3', ft=6.9)
seta(8.65, 4.9, 9.05, 5.35)
caixa(9.1, 4.85, 2.85, 2.1, 'Análise\nexploratória', 'evasão, ocupação\nde vagas', etapa='Etapa 4', cor_borda='#d95f02', cor_fundo=LARANJA_CLARO)
seta(8.65, 3.0, 9.05, 2.7)
caixa(9.1, 1.0, 2.85, 3.5, 'Modelagem e\navaliação', 'RL, RF e XGBoost\nteste 2024–2025\nwalk-forward\nreferências ingênuas\nablação, SHAP', etapa='Etapas 5 e 6')
caixa(12.35, 1.9, 1.6, 3.3, 'Resultados', 'tabelas,\nfiguras,\nprojeção\n2026–2029,\nnotebook', etapa='Etapa 7', cor_borda='#1a1a1a', cor_fundo='white', ft=6.9)
seta(12.0, 5.9, 12.75, 5.25)
seta(12.0, 2.5, 12.35, 2.5)
fig.subplots_adjust(0, 0, 1, 1)
fig.savefig(saida, dpi=300, facecolor='white')
print('ok', saida)
