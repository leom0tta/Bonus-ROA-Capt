"""
Esse código é uma rotina voltada para o monitoramento da entrada de clientes, 
por mês. Essa baseé usada na alimentação do dashboard_unificado, e esses números
são importantes para o acompanhamento da posição da Fatoial no ranking G20.
"""

import pandas as pd
import sys
sys.path.insert(1, r'.\Funções')
from funcoes import add_assessor_indicador, add_nome_assessor, gera_excel

data_hoje = '050422'
mes = 'Abril'

contas_novas = pd.read_excel(r'captacao_diario\captacao_' + data_hoje + '.xlsx', sheet_name='Contas novas')
contas_novas = contas_novas[['Assessor', 'Cliente', 'Net Em M']]
contas_novas['Assessor'] = contas_novas['Assessor'].astype(str)

clientes_rodrigo = pd.read_excel(r'bases_dados\Clientes do Rodrigo.xlsx', sheet_name='Troca')
clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Indicador']]
clientes_rodrigo['Assessor Indicador'] = clientes_rodrigo['Assessor Indicador'].astype(str)

assessores = pd.read_excel (r"bases_dados\Assessores leal_Pablo.xlsx", sheet_name='Plan1')
assessores['Código assessor'] = assessores['Código assessor'].astype(str)

caminho_excel = r"G20\Contas novas\Contas_novas_" + mes + ".xlsx"

contas_novas = add_assessor_indicador(contas_novas, clientes_rodrigo)

contas_novas = add_nome_assessor(contas_novas, 'Assessor Indicador', assessores)

contas_novas = contas_novas[['Assessor Indicador', 'Nome assessor', 'Cliente', 'Net Em M']]

gera_excel(contas_novas,caminho_excel)





