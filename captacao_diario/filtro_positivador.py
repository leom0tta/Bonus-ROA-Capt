import datetime
import pandas as pd
import numpy as np
from pandas.io import excel
import xlsxwriter as xlsx
from pathlib import Path

# planilhas usadas: positivador final do mês M-1 ; positivador mês M ; planilha transferências ; clientes rodrigo ; clientes meta ; Assessores ; relatório captação mês m-1 ; relatório captação mês m-2 ; relatório captação d-1 ; suitability

#importando as planilhas utilizadas ----------------------------------------------------------------------------------------------------------------------------

year = 2022
responsavel_digital = "Atendimento Fatorial"
data_hoje = '190422'
data_ontem = '140422'

#importando positivador mês M 
caminho_novo = Path(r'captacao_diario\positivador_' + data_hoje + '.xlsx') #! positivador do dia de run do código
posi_novo = pd.read_excel(caminho_novo, skiprows= 2)

# importando positivador do dia D-1
caminho_d1 = Path(r'captacao_diario\positivador_' + data_ontem + '.xlsx') #! positivador do dia anterior ao de run do codigo
posi_d1 = pd.read_excel(caminho_d1 , skiprows= 2) 

# importando positivador mês M - 1 
caminho_velho = Path(r'captacao_diario\arquivos\2022\Março\positivador_310322.xlsx') #! último positivador do mês anterior !
posi_velho = pd.read_excel(caminho_velho , skiprows=2)

#importando clientes rodrigo e fazendo tratamentos (selecionando as colunas , substituindo "Atendimento" pelo A do yago e tirando o prefixo A)
caminho_rodrigo = Path(r'bases_dados\Clientes do Rodrigo.xlsx') #! verificar se a clientes do rodrigo tá atualizada

clientes_rodrigo = pd.read_excel(caminho_rodrigo , sheet_name='Troca')

clientes_rodrigo = clientes_rodrigo.loc[: , ['Conta' , 'Assessor Relacionamento']]

#clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'] == 'Atendimento Fatorial' , 'Assessor Relacionamento'] = "A" + responsavel_digital

clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'].isin(['Selma' , 'Paulo Barros' , 'Paulo Monfort','Atendimento Fatorial']) == False , 'Assessor Relacionamento'] = pd.Series(clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'].isin(['Selma' , 'Paulo Barros' , 'Paulo Monfort','Atendimento Fatorial']) == False , 'Assessor Relacionamento']).str.lstrip('A')

#importando lista de assessores 
caminho_assessores = Path(r'bases_dados\Assessores leal_Pablo.xlsx') #!
assessores = pd.read_excel(caminho_assessores)
assessores = assessores.astype({'Código assessor' : str})

#importando lista de transferências
caminho_transf = Path(r'captacao_diario\transferencias_' + data_hoje + '.xlsx') #! pegar no connect a do mesmo dia

lista_transf = pd.read_excel(caminho_transf)

lista_transf = lista_transf.loc[lista_transf['Status'] == 'CONCLUÍDO' , :]

#importando o relatório de captação D-1

caminho_ontem = Path(r'captacao_diario\captacao_' + data_ontem + '.xlsx') #!#!#!#!#!#

clientes_novos_ontem = pd.read_excel(caminho_ontem , sheet_name='Novos + Transf')

#importando o suitability

caminho_suitability = Path(r'bases_dados\Suitability.xlsx') #!#!#!#!#!#!#!#

suitability = pd.read_excel(caminho_suitability)

#diretorio arquivo final

caminho_excel = Path(r'captacao_diario\positivador_'  + data_hoje + '.xlsx',index = False) #!

#diretorio arquivo de registro

caminho_registro = Path(r'bases_dados\Registro de Transferências\registro_transferência_12_21.xlsx') #MUDAR PARA O MÊS DE RUN DO CODIGO

registro_transf = pd.read_excel(caminho_registro)

registro_transf.columns = ['Assessor', 'Cliente', 'Data de Chegada', 'Net de Chegada']

#montando tabela clientes perdidos ----------------------------------------------------------------------------------------------------------------------------

#encontrando quais são os clientes perdidos
posi_velho['Clientes perdidos'] = posi_velho['Cliente'].where(posi_velho['Cliente'].isin(posi_novo['Cliente']) == True)

#renomeando os valores: na -> "Saiu" ; código cliente -> "Permanece"
posi_velho['Clientes perdidos'].fillna('Saiu' , inplace = True)

posi_velho.loc[posi_velho['Clientes perdidos'] != 'Saiu' , 'Clientes perdidos'] = 'Permanece'

#montagem do dataframe
tabela_perdidos = posi_velho.loc[posi_velho['Clientes perdidos'] == 'Saiu' , :]

#criação de coluna de assessor correto
tabela_perdidos = pd.merge(tabela_perdidos , clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how= 'left')

tabela_perdidos['Assessor Relacionamento'].fillna(tabela_perdidos['Assessor'] , inplace = True)

tabela_perdidos.rename(columns={'Assessor Relacionamento' : 'Assessor correto'} , inplace= True)

del tabela_perdidos['Conta']

tabela_perdidos.loc[tabela_perdidos['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

# montando tabela de clientes velhos ----------------------------------------------------------------------------------------------------------------------------

#encontrando quais são os clientes velhos
#posi_novo['Status conta'] = posi_novo['Cliente'].where(posi_novo['Cliente'].isin(posi_velho['Cliente']) == True)

posi_novo.loc [posi_novo["Cliente"].isin (posi_velho['Cliente']),"Status conta"] = 'conta velha'

posi_novo['Status conta'].fillna('conta nova' , inplace = True)

posi_novo.drop_duplicates (subset = "Cliente",inplace=True)

#posi_novo.loc[posi_novo['Status conta'] != 'conta nova' , 'Status conta'] = 'conta velha'

#seleção das contas velhas
tabela_velhos = posi_novo.loc[posi_novo['Status conta'] == 'conta velha' , :]

#criação da coluna de assessor correto
tabela_velhos = pd.merge(tabela_velhos , clientes_rodrigo , left_on= 'Cliente' , right_on= 'Conta' , how= 'left')

tabela_velhos['Assessor Relacionamento'].fillna(tabela_velhos['Assessor'] , inplace = True)

tabela_velhos.rename(columns={'Assessor Relacionamento':'Assessor correto'} , inplace= True)

del tabela_velhos['Conta']

tabela_velhos.loc[tabela_velhos['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

# montando tabela de clientes novos e transferencias ----------------------------------------------------------------------------------------------------------------------------

#tabela clientes novos + transferencias
tabela_novos_transf = posi_novo.loc[posi_novo['Status conta'] == 'conta nova' , :]


#identificando quais são as transferências
print(lista_transf.loc[: , 'Cliente'])
tabela_novos_transf.loc[: ,'Transferência?'] = tabela_novos_transf.loc[: , 'Cliente'].where(tabela_novos_transf.loc[: , 'Cliente'].isin(lista_transf.loc[: , 'Cliente']) == True)

tabela_novos_transf['Transferência?'].fillna('Não' , inplace = True)

tabela_novos_transf.loc[tabela_novos_transf['Transferência?'] !='Não' , 'Transferência?'] = 'Sim'

tabela_novos_transf.loc[tabela_novos_transf['Net em M-1'] > 0 , 'Transferência?'] = 'Sim'

print(tabela_novos_transf[tabela_novos_transf['Assessor'] == 24999])

#tabela clientes novos
tabela_novos = tabela_novos_transf.loc[tabela_novos_transf['Transferência?'] == 'Não' , :]

#cliente rodrigo para os novos
tabela_novos = pd.merge(tabela_novos , clientes_rodrigo , left_on= 'Cliente' , right_on= 'Conta' , how = 'left')

tabela_novos['Assessor Relacionamento'].fillna(tabela_novos['Assessor'] , inplace = True)

tabela_novos.rename(columns = {'Assessor Relacionamento' : 'Assessor correto'} , inplace= True)

del tabela_novos['Conta']

tabela_novos.loc[tabela_novos['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

#tabela transferências
tabela_transf = tabela_novos_transf.loc[tabela_novos_transf['Transferência?'] == 'Sim' , :]

#cliente rodrigo para as transferências
tabela_transf = pd.merge(tabela_transf , clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how='left')

tabela_transf['Assessor Relacionamento'].fillna(tabela_transf['Assessor'] , inplace = True)

tabela_transf.rename(columns = {'Assessor Relacionamento' : 'Assessor correto'} , inplace = True)

del tabela_transf['Conta']

tabela_transf.loc[tabela_transf['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

#Procura a transferencia em positivador D-1

clientes_transferidos = tabela_transf[['Cliente']]

clientes_ate_ontem = posi_d1['Cliente'].to_numpy()

mask_antigos = clientes_transferidos.isin(clientes_ate_ontem)

transferidos_antigos = clientes_transferidos[mask_antigos].dropna().to_numpy()

transferidos_novos = clientes_transferidos[~mask_antigos].dropna().to_numpy()

print('\n\nTransferências antigas:\n', transferidos_antigos, '\n\nTransferências novas:\n', transferidos_novos, '\n\n')

# Atualizar o Net em M-1 dos clientes que foram transferidos antes de hoje

# Filtrar os clientes transferidos antigos no positivador do dia atual

posi_mask_transf_antigos = []

for *_ , cliente in posi_novo['Cliente'].items():
    if cliente in transferidos_antigos:
        posi_mask_transf_antigos.append(True)
    else:
        posi_mask_transf_antigos.append(False)

positivador_transf_antigos = posi_novo[posi_mask_transf_antigos]

positivador_transf_antigos.sort_values('Cliente', inplace = True)

# Filtrar os clientes transferidos antigos no positivador do dia anterior

posi_1_mask_transf_antigos = []
posi_1_index_transf_antigo = []

for *_ , cliente in posi_d1['Cliente'].items():
    if cliente in transferidos_antigos:
        posi_1_mask_transf_antigos.append(True)
    else:
        posi_1_mask_transf_antigos.append(False)

positivador_1_transf_antigos = posi_d1[posi_1_mask_transf_antigos]
positivador_1_transf_antigos.sort_values('Cliente', inplace = True)

# Gerar dois arrays, um com os patrimônios registrados no posi_antigo, outro com os index a terem seu patrimônio modificado no posi_novo

net_m_1 = positivador_1_transf_antigos['Net em M-1'].to_numpy()
index_transf_antigos = positivador_transf_antigos.index.to_numpy()

# Atualizar o valor de M-1 no positivador do dia atual, de acordo com o valor no dia anterior

for i, patrimonio in enumerate(net_m_1):
    index = index_transf_antigos[i]
    posi_novo.loc[ index , 'Net em M-1'] = patrimonio

posi_novo.sort_index(inplace = True)

# Vê se tem algum cliente novo, se não, nem passa por aqui o código

if len(transferidos_novos) > 0:

    # Considera o Net em M-1 igual ao Net em M, para o caso dos transferidos no dia corrente

    posi_mask_transf_novos = []

    for *_ , cliente in posi_novo['Cliente'].items():
        if cliente in transferidos_novos:
            posi_mask_transf_novos.append(True)
        else:
            posi_mask_transf_novos.append(False)

    positivador_transf_novos = posi_novo[posi_mask_transf_novos]
    
    # Computa em duas listas os nets em M e o index dos clientes que estão errados

    lista_nets = positivador_transf_novos['Net Em M'].to_numpy()

    lista_index_novos = positivador_transf_novos.index.to_numpy()

    # Registra os assessores que ganharam clientes

    lista_assessores = positivador_transf_novos['Assessor'].to_numpy()

    # Atualiza os clientes no positivador, através da lista do inde dos clientes novos e do patrimônio certo

    for i, patrimonio in enumerate(lista_nets):
        index = lista_index_novos[i]
        posi_novo.loc[index, 'Net em M-1'] = patrimonio

    # Registra os clientes novos

    lista_dias = [datetime.date(year = year, month = int(data_hoje[2:4]), day = int(data_hoje[:2])) for i in range(len(transferidos_novos))] # MUDAR PARA O DIA EM QUESTÃO

    colunas_registro = registro_transf.columns

    lista_transferidos = [cod_cliente[0] for cod_cliente in transferidos_novos.astype(int)]

    data_transf = np.transpose(np.array([lista_assessores, lista_transferidos, lista_dias, lista_nets]))

    df_transf_novos = pd.DataFrame(
        data_transf, 
        columns = colunas_registro, 
        index = None)

    registro_transf = registro_transf.append(df_transf_novos)

    registro_transf.drop_duplicates(subset=['Cliente'], inplace = True)

    # Escreve a planilha de registro

    writer = pd.ExcelWriter(caminho_registro , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
    registro_transf.to_excel(writer , sheet_name='Registro',index=False)

    arquivo = writer.book
    aba = writer.sheets['Registro']

    writer.save ()

# coloca os aniversários que vem faltando

sem_aniversario = posi_novo[posi_novo['Data de Nascimento'].isnull()]

clientes_sem_aniversario = sem_aniversario['Cliente'].to_numpy()

clientes_ontem = posi_d1['Cliente'].to_numpy()

for cliente in clientes_sem_aniversario:
    if cliente in clientes_ontem:
        aniversario_cliente = posi_d1.loc[ posi_d1['Cliente'] == cliente, 'Data de Nascimento' ].to_numpy()[0]
        posi_novo.loc[ posi_novo['Cliente'] == cliente, 'Data de Nascimento' ] = aniversario_cliente

print("\nClientes sem aniversário:\n\n", posi_novo[posi_novo['Data de Nascimento'].isnull()], '\n')


# Escreve o positivador no excel

writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
posi_novo.to_excel(writer , sheet_name='Sheet1',index=False, startrow=2)

arquivo = writer.book
aba = writer.sheets['Sheet1']

writer.save ()

print('arquivo criado')













