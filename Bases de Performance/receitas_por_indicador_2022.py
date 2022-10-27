import pandas as pd
import sys
sys.path.insert(1, r'.\Funções')
from funcoes import gera_excel
from classes import Data, Dataframe, Get

list_fechamentos= ['310122', '250222', '310322', '290422', '310522', '300622', '280722','310822']
list_data_obj = [Data(fechamento) for fechamento in list_fechamentos]

assessores_leal = Get.assessores()
suitability = Get.suitability()
tags = Get.tags_comissoes()

suitability['CodigoBolsa'] = suitability['CodigoBolsa'].astype(str)
assessores_leal['Código assessor'] = assessores_leal['Código assessor'].astype(str)

writer = pd.ExcelWriter('receitas_2022_jan-agosto.xlsx' , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')

for data_obj in list_data_obj:
    mes = data_obj.text_month

    comissao = Get.receitas(mes)
    comissao['Cliente'] = comissao['Cliente'].astype(str)

    clientes_rodrigo = Get.clientes_rodrigo(mes)
    clientes_rodrigo['Conta'] = clientes_rodrigo['Conta'].astype(str)

    mask_ajustes = comissao['Categoria'].isin(['Ajuste', 'Ajustes XP', 'Outros Ajustes', 'Incentivo Comercial', 'Complemento de Comissão Mínima', 'Desconto de Transferência de Clientes', 'Incentivo', 'Erro Operacional'])

    comissao = comissao[~mask_ajustes]

    # assimila cada cliente unica e exclusivamente ao seu indicador

    indicadores = clientes_rodrigo[['Conta', 'Assessor Indicador']]
    comissao = comissao.merge(indicadores, how='left', left_on='Cliente', right_on='Conta')
    del comissao['Conta']

    comissao['Assessor Dono'] = comissao['Assessor Indicador'].fillna(comissao['Assessor Dono'])

    # separa por centro de custo e por assessor

    by_tag = comissao.merge(tags, how='outer', on='Categoria')
    by_tag = by_tag.groupby(['Assessor Dono', 'Centro de Custo']).sum().loc[: , 'Valor Bruto Recebido']
    by_tag = by_tag.reset_index(drop=False)
    by_tag['Mês'] = mes

    by_tag['Assessor Dono'] = by_tag['Assessor Dono'].str.lstrip('A')
    by_tag['Assessor Dono'].replace("DRIANO MENEGUITE", 'ADRIANO MENEGUITE', inplace=True)
    by_tag['Assessor Dono'].replace("LINY MANZIERI", 'ALINY MANZIERI', inplace=True)
    by_tag['Assessor Dono'].replace("TENDIMENTO FATORIAL", 'ATENDIMENTO FATORIAL', inplace=True)
    
    by_tag.rename(columns={'Assessor Dono': 'Código assessor'}, inplace=True)
    by_tag.rename(columns={'Valor Bruto Recebido': f'Receita'}, inplace=True)

    if mes == 'Janeiro':
        df_tag = by_tag
    else:
        df_tag = pd.concat([df_tag, by_tag])

    # separa só por assessor

    receita = comissao.groupby('Assessor Dono').sum()['Valor Bruto Recebido']
    receita = receita.reset_index(drop=False)

    receita['Assessor Dono'] = receita['Assessor Dono'].str.lstrip('A')
    receita['Assessor Dono'].replace("DRIANO MENEGUITE", 'ADRIANO MENEGUITE', inplace=True)
    receita['Assessor Dono'].replace("LINY MANZIERI", 'ALINY MANZIERI', inplace=True)
    receita['Assessor Dono'].replace("TENDIMENTO FATORIAL", 'ATENDIMENTO FATORIAL', inplace=True)

    receita.rename(columns={'Assessor Dono': 'Código assessor'}, inplace=True)
    receita.rename(columns={'Valor Bruto Recebido': f'Receita {mes}'}, inplace=True)

    if mes == 'Janeiro':
        df_assessores = assessores_leal.merge(receita, how='outer', on='Código assessor')
    else:
        df_assessores = df_assessores.merge(receita, how='outer', on='Código assessor')

posi = Get.captacao(data_obj.cod_data, sheet_name='Positivador M')

posi['Cliente'] = posi['Cliente'].astype(str)
posi['Assessor'] = posi['Assessor'].astype(str)

posi = Dataframe(posi)

posi.add_assessor_indicador(clientes_rodrigo)

carteira_indicador = posi.dataframe.groupby('Assessor Indicador').sum()[['Net Em M']]

df_assessores['Código assessor'] = df_assessores['Código assessor'].astype(str)

df_assessores = df_assessores.merge(carteira_indicador, how='left', right_index=True, left_on='Código assessor')

df_assessores.to_excel(writer , sheet_name= 'Resumo Assessores', index=False)

df_tag = pd.merge(assessores_leal, df_tag, how='right', on='Código assessor')

df_tag.to_excel(writer, sheet_name='Resumo Tags', index=False)

writer.save()