import pandas as pd
import sys
sys.path.insert(1, r'.\Funções')
from funcoes import gera_excel

meses= [['Janeiro', '310122'], ['Fevereiro', '250222'], ['Março', '310322'], ['Abril', '290422'], ['Maio', '310522'], ['Junho', '300622']]

assessores_leal = pd.read_excel (r"bases_dados\Assessores leal_Pablo.xlsx")
suitability = pd.read_excel (r"bases_dados\Suitability.xlsx")

suitability['CodigoBolsa'] = suitability['CodigoBolsa'].astype(str)
assessores_leal['Código assessor'] = assessores_leal['Código assessor'].astype(str)

writer = pd.ExcelWriter('teste_receitas_2022.xlsx' , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')

for mes, data_hoje in meses:

    comissao = pd.read_csv(r"Comissões\Receitas\Bases SplitC\dados_comissão_" + str.lower(mes.replace('ç', 'c')) + ".csv",sep=';', decimal=',')
    comissao['Cliente'] = comissao['Cliente'].astype(str)

    clientes_rodrigo = pd.read_excel(r"Clientes Rodrigo\Clientes Rodrigo " + mes + ".xlsx")
    clientes_rodrigo['Conta'] = clientes_rodrigo['Conta'].astype(str)

    mask_ajustes = comissao['Categoria'] == "Ajuste"
    mask_incentivo = comissao['Categoria'] == 'Incentivo Comercial'

    comissao = comissao[~(mask_ajustes | mask_incentivo)]

    # assimila cada cliente do B2C unica e exclusivamente ao seu indicador

    indicadores = clientes_rodrigo[['Conta', 'Assessor Indicador']]
    comissao = comissao.merge(indicadores, how='left', left_on='Cliente', right_on='Conta')
    del comissao['Conta']

    mask_celulas = comissao['Assessor Dono'].isin(['A26839', 'ATENDIMENTO FATORIAL', 'A26877', 'A26994'])

    comissao_pra_distribuir = comissao[mask_celulas]
    comissao_sem_distribuir = comissao[~mask_celulas]

    comissao_pra_distribuir['Assessor Dono'] = comissao_pra_distribuir['Assessor Indicador'].fillna(comissao_pra_distribuir['Assessor Dono'])

    comissao = pd.concat([comissao_pra_distribuir, comissao_sem_distribuir])

    # dividir a receita dos indicadores e do B2B

    mask_sem_indicador = comissao['Assessor Indicador'].isna()
    
    nao_dividir =  mask_celulas | mask_sem_indicador

    comissao_a_dividir = comissao[~nao_dividir]
    comissao_sem_divisao = comissao[nao_dividir]

    comissao_indicador = comissao_a_dividir.copy()
    comissao_indicador['Assessor Dono'] = comissao_indicador['Assessor Indicador']
    comissao_indicador['Valor Bruto Recebido'] *= 0.5

    comissao_relacionamento = comissao_a_dividir.copy()
    comissao_relacionamento['Valor Bruto Recebido'] *= 0.5

    comissao_dividida = pd.concat([comissao_relacionamento, comissao_indicador])

    comissao = pd.concat([comissao_sem_divisao, comissao_dividida])

    #comissao.to_excel(writer, mes, index=False)

    receita = comissao.groupby('Assessor Dono').sum()['Valor Bruto Recebido']
    receita = receita.reset_index(drop=False)

    receita['Assessor Dono'] = receita['Assessor Dono'].str.lstrip('A')
    receita['Assessor Dono'].replace("DRIANO MENEGUITE", 'ADRIANO MENEGUITE', inplace=True)
    receita['Assessor Dono'].replace("LINY MANZIERI", 'ALINY MANZIERI', inplace=True)
    receita['Assessor Dono'].replace("TENDIMENTO FATORIAL", 'ATENDIMENTO FATORIAL', inplace=True)

    receita.rename(columns={'Assessor Dono': 'Código assessor'}, inplace=True)
    receita.rename(columns={'Valor Bruto Recebido': f'Receita {mes}'}, inplace=True)

    if mes == 'Janeiro':
        df = assessores_leal.merge(receita, how='outer', on='Código assessor')
    else:
        df = df.merge(receita, how='outer', on='Código assessor')

df.to_excel(writer , sheet_name= 'Resumo', index=False)

writer.save()