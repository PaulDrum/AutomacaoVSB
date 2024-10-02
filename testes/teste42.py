import glob, os
import pandas as pd
from datetime import date
from openpyxl.styles import Alignment, Side, Border
from openpyxl import load_workbook
from time import sleep
from matplotlib import pyplot as plt

user = os.path.expanduser('~')

arquivo = glob.glob(user + r'\Downloads\*VALLOURECADM*.xls')

center_align = Alignment(horizontal='center', vertical='center')
box = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

data_hoje = date.today().strftime("%d/%m/%Y")

barreiro_df = pd.read_excel(arquivo[0])

linha = barreiro_df[barreiro_df['Unnamed: 0'] == 'Veículos dentro das Áreas'].index[0]

barreiro_df.drop(barreiro_df.columns[10:17], axis=1, inplace=True)

barreiro_df.drop(barreiro_df.index[:linha+1], axis=0, inplace=True)

barreiro_df.columns = barreiro_df.iloc[0]

barreiro_df = barreiro_df[1:]

barreiro_df = barreiro_df.rename(columns={'Sub-grupo': 'Data', 'Grupo da Área': 'Identificador'})

barreiro_df['Data'] = data_hoje

barreiro_df['Identificador'] = 'Barreiro-Tradimaq'

barreiro_df = barreiro_df[barreiro_df['Categoria'].str.contains('OFICINA', na=False)]

x = barreiro_df['Tempo dentro'].str.split(':').str[0].astype(int)

x = x.sort_values(ascending=False)

y = barreiro_df['Nome do veículo']

fig, ax = plt.subplots(figsize =(16, 9))

ax.barh(y, x)

for s in ['top', 'bottom', 'left', 'right']:
    ax.spines[s].set_visible(False)

ax.grid(b = True, color ='grey',
        linestyle ='-.', linewidth = 0.5,
        alpha = 0.2)

ax.xaxis.set_ticks_position('none')
ax.yaxis.set_ticks_position('none')

ax.xaxis.set_tick_params(pad = 5)
ax.yaxis.set_tick_params(pad = 10)

for i in ax.patches:
    plt.text(i.get_width()+0.5, i.get_y()+0,
             str(round((i.get_width()), 2)),
             fontsize = 8,
             color ='black')

ax.set_title('Equipamentos no Canteiro',
             loc ='left', )

fig.text(0.9, 0.15, 'Shalkz', fontsize = 12,
         color ='grey', ha ='right', va ='bottom',
         alpha = 0.7)

plt.xlabel("Horas")

plt.show()

# os.remove(arquivo)