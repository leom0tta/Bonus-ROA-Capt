import pandas as pd
import sys
sys.path.insert(1, r'.\Funções')
from funcoes import gera_excel
from classes import Get, Data

meses= ['310122', '250222', '310322', '290422', '310522','300622', '280722', '310822', '300922']

assessores_leal = Get.assessores()
suitability = Get.suitability()

for data_hoje in meses:

    mes = Data(data_hoje).text_month

    captacao_mes = Get.captacao(data_hoje)

    captacao_mes = captacao_mes[captacao_mes['Nome assessor'] != 'Total Fatorial']

    captacao_mes['Mes'] = mes

    if mes == 'Janeiro':
        df = captacao_mes
    else:
        df = pd.concat([df, captacao_mes])

gera_excel(df, 'Bases de Performance\Base Dados\captacao_2022.xlsx', 'Captação 2022')



