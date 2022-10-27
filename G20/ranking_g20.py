import pandas as pd
import numpy as np
import sys
sys.path.insert(1, r'.\Funções')
from funcoes import gera_excel, reorder_columns
from classes import Get
import warnings
warnings.filterwarnings("ignore")

mes = "Setembro"

relatorio = pd.read_excel(r"G20\Relatorio\Relatório G20 " + mes + ".xlsx")
assessores = Get.assessores()

caminho_excel = r"G20\Relatorio\Ranking G20 " + mes + ".xlsx"
caminho_detalhado = r"G20\Relatorio\Ranking G20 Detalhado " + mes + ".xlsx"

df_list = []
list_categories = ['Captação Líquida', 'Faturamento', 'Contas Novas', 'NPS Aniversário', 'NPS Onboarding']

mask_felipe_ribeiro = relatorio['Código assessor'] == '24152'
relatorio = relatorio[~mask_felipe_ribeiro]

relatorio['Contas Novas'] = (relatorio['Conta Nova +300k']*1 + relatorio['Conta Nova +1M']*2) - (relatorio['Conta Perdida +300k']*0.5 + relatorio['Conta Perdida +1M']*1)

points_weights = {

    'Captação Líquida': [
    [-100e6,-100],
    [-10e6,-50],
    [-9e6,-45],
    [-8e6,-40],
    [-7e6,-35],
    [-6e6,-30],
    [-5e6,-25],
    [-4e6,-20],
    [-3e6,-15],
    [-2e6,-10],
    [-1e6,-5],
    [-.5e6,-3],
    [-.3e6,-0],
    [.3e6,3],
    [.4e6,4],
    [.5e6,5],
    [.6e6,6],
    [.7e6,7],
    [.8e6,8],
    [.9e6,9],
    [1e6,10],
    [1.5e6,15],
    [2e6,20],
    [2.5e6,25],
    [3e6,30],
    [3.5e6,35],
    [4e6,40],
    [5e6,50],
    [6e6,60],
    [7e6,70],
    [8e6,80],
    [9e6,90],
    [10e6,100]
    ],

    'Faturamento': [
    [0,0],
    [.3e4,3],
    [.5e4,5],
    [1e4,10],
    [2e4,20],
    [3e4,30],
    [4e4,40],
    [5e4,50],
    [6e4,60],
    [7e4,70],
    [8e4,80],
    [9e4,90],
    [10e4,100]    
    ],

    'Contas Novas': [
    [-10,-100],
    [-9,-90],
    [-8,-80],
    [-7,-70],
    [-6,-60],
    [-5,-50],
    [-4,-40],
    [-3,-30],
    [-2,-20],
    [-1,-10],
    [0,0],
    [1,10],
    [2,20],
    [3,30],
    [4,40],
    [5,50],
    [6,60],
    [7,70],
    [8,80],
    [9,90],
    [10,100]
    ],

    'NPS Aniversário': [
    [0,0],
    [80,20],
    [85,40],
    [90,60],
    [95,80],
    [100,100]
    ],

    'NPS Onboarding': [
    [0,0],
    [80,20],
    [85,40],
    [90,60],
    [95,80],
    [100,100]
]
}

writer = pd.ExcelWriter(f'G20\Relatorio\Faixas de Pontuação\Faixas de Pontuação.xlsx' , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')

for categorie in list_categories:

    pontos = pd.DataFrame(points_weights[categorie], columns=['Valor Mínimo da Faixa', 'Pontuação'])

    if categorie in ['Captação Líquida', 'Faturamento']:

        bottom_formatado = []
        for bottom in pontos['Valor Mínimo da Faixa']:
            bottom = "R$ {:,.2f}".format(bottom)
            bottom_formatado.append(bottom)
        
        pontos['Valor Mínimo da Faixa'] = bottom_formatado

    pontos.to_excel(writer, categorie, index=False)

    df = relatorio[['Código assessor', categorie]]

    df.sort_values(categorie, inplace=True, ascending = False)
    df.reset_index(inplace=True, drop=True)

    df[f'Pontos {categorie}'] = np.zeros( len(df.index) )

    for bottom_value, points in points_weights[categorie]:
        mask = df[categorie] >= bottom_value
        df[f'Pontos {categorie}'][mask] = points
    
    df_list.append(df)

peso = pd.DataFrame({
    'Valor Mínimo da Faixa': ['Qualificação','40%','20%','0%','Zerados'],
    'Peso': ['4+ envios', '1.0', '0.95', '0.9', '0.8']})

peso.to_excel(writer, sheet_name='Percentual Respostas', index=False)

writer.save()

rank_captacao = df_list[0]
rank_receitas = df_list[1]
rank_ativacao = df_list[2]
rank_nps_aniversario = df_list[3]
rank_nps_onboarding = df_list[4]

rank_geral = rank_captacao.merge(rank_receitas, how='outer', on='Código assessor')
rank_geral = rank_geral.merge(rank_ativacao, how='outer', on='Código assessor')
rank_geral = rank_geral.merge(rank_nps_aniversario, how='outer', on='Código assessor')
rank_geral = rank_geral.merge(rank_nps_onboarding, how='outer', on='Código assessor')

rank_geral.fillna(0, inplace=True)

pesos = pd.Series({
    'Captação Líquida': 0.3,
    'Faturamento': 0.3,
    'Contas Novas': 0.3,
    'NPS Aniversário': 0.05,
    'NPS Onboarding': 0.05
})

classificacoes = rank_geral[[
    'Pontos Captação Líquida', 
    'Pontos Faturamento', 
    'Pontos Contas Novas', 
    'Pontos NPS Aniversário', 
    'Pontos NPS Onboarding']]

pontos = np.dot(classificacoes, np.transpose(pesos))

rank_geral['Pontuação Geral'] = pontos

rank_geral.sort_values('Pontuação Geral', inplace=True, ascending=False)
rank_geral.reset_index(inplace=True, drop=True)

# pondera a partir do percentual de respostas

percentual_respostas = relatorio[['Código assessor', 'Tamanho da amostra', 'Percentual de Resposta']]
percentual_respostas['Percentual de Resposta'] = percentual_respostas['Percentual de Resposta'].round(2)

nao_tem_min = relatorio['Número de Envios'] < 4

percentual_respostas.loc[nao_tem_min, 'Percentual de Resposta'] = np.nan

percentual_respostas['Fator de Peso'] = [1 for i in percentual_respostas.index]

mask_40 = percentual_respostas['Percentual de Resposta'] < 0.4
mask_20 = percentual_respostas['Percentual de Resposta'] < 0.2
mask_zerados = percentual_respostas['Percentual de Resposta'] == 0

percentual_respostas.loc[mask_40, 'Fator de Peso'] = 0.95
percentual_respostas.loc[mask_20, 'Fator de Peso'] = 0.90
percentual_respostas.loc[mask_zerados, 'Fator de Peso'] = 0.80

rank_geral = rank_geral.merge(percentual_respostas[['Código assessor','Percentual de Resposta', 'Fator de Peso']], how='left', on='Código assessor')

rank_geral['Pontuação Geral'] *= rank_geral['Fator de Peso']

# assimila a posição

posicoes = rank_geral['Pontuação Geral'].sort_values(ascending=False).reset_index(drop=True).drop_duplicates()
posicoes = posicoes.reset_index()

rank_geral = rank_geral.merge(posicoes, how='left', on='Pontuação Geral')
rank_geral.rename(columns={'index': 'Posição Geral'}, inplace=True)
rank_geral['Posição Geral'] += 1

rank_geral = pd.merge(rank_geral, assessores, on='Código assessor')
rank_geral = reorder_columns(rank_geral, 'Nome assessor', 1)
rank_geral = reorder_columns(rank_geral, 'Time', 2)
rank_geral['Nome assessor'].fillna(rank_geral['Código assessor'], inplace=True)

rank_geral = reorder_columns(rank_geral, 'Pontuação Geral', 15)

gera_excel(rank_geral, caminho_excel, index=False)

# detalhado

assessores = rank_geral['Nome assessor'].drop_duplicates()

relatorio_detalhado = pd.DataFrame(columns=['Valores', 'Nome assessor'])

for assessor in assessores:
    mask_assessor = rank_geral['Nome assessor'] == assessor
    dataset_assessor = rank_geral[mask_assessor]
    dataset_assessor = dataset_assessor[['Pontos Captação Líquida', 'Pontos Faturamento', 'Pontos Contas Novas', 'Pontos NPS Aniversário', 'Pontos NPS Onboarding']]
    dataset_assessor = dataset_assessor.transpose()
    dataset_assessor['Nome assessor'] = assessor
    dataset_assessor.columns = ['Valores', 'Nome assessor']
    relatorio_detalhado = pd.concat([relatorio_detalhado, dataset_assessor], axis=0)

gera_excel(relatorio_detalhado, caminho_detalhado, index=True)
