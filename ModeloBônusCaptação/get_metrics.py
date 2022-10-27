import pandas as pd
import numpy as np
import sys
sys.path.insert (1, r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\Funções')
from classes import Get
from range_capt import range_capt_bonus, range_roa

assessores = Get.assessores()

captacao = Get.captacao_acumulada()
receita = Get.receita_acumulada()

captacao['Captação Externa'] = captacao['Total conta velha'] + captacao['Total conta nova'] + captacao['Total contas perdidas']
captacao['Captação Interna'] = captacao['Total transferências']

del captacao['Total conta velha'] 
del captacao['Total conta nova'] 
del captacao['Total contas perdidas']
del captacao['Total transferências']
del captacao['NET XP']
del captacao['Ticket Médio']
del captacao['Qtd Clientes XP']

del receita['Nome assessor']
del captacao['Time']

receita = receita.groupby(['Código assessor', 'Mes']).sum().reset_index(drop=False)

df = captacao.merge(receita, how='left', on=['Código assessor', 'Mes'])

df['ROA'] = df['Receita'] / df['Net Em M'] * 12

average_df = df[ df['Código assessor'].isin(assessores['Código assessor']) ]

average_df['ROA'].fillna(0, inplace=True)

average_df = average_df[['Nome assessor', 'Captação Líquida', 'ROA']].groupby('Nome assessor').mean()

price = []

for roa_min in range_roa:
    
    average_df[str(roa_min)] = [
        0 if roa<roa_min or capt<0 
        else 1 
        for roa, capt in average_df[['ROA', 'Captação Líquida']].to_numpy()
        ]
        
    price += [
        [
        roa_min, capt,
        sum(12*average_df['Captação Líquida']*average_df[str(roa_min)]*capt),
        sum(average_df[str(roa_min)])
        ] 
        for capt in range_capt_bonus]

price = pd.DataFrame(price, columns = ['Roa Min', 'Bônus Capt', 'Preço', 'Qtd. Beneficiados'])

writer = pd.ExcelWriter('ModeloBônusCaptação\BD\métricas_captação.xlsx' , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')

average_df.to_excel(writer, sheet_name = 'average_df', index=True)
price.to_excel(writer, sheet_name='price', index=False)
df.to_excel(writer, sheet_name='df', index=False)

writer.save()