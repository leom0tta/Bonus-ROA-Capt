import pandas as pd
import xlsxwriter as xlsx
from pathlib import Path

# planilhas usadas: positivador final do mês M-1 ; positivador mês M ; planilha transferências ; clientes rodrigo ; clientes meta ; Assessores ; relatório captação mês m-1 ; relatório captação mês m-2 ; relatório captação d-1 ; suitability

#importando as planilhas utilizadas ----------------------------------------------------------------------------------------------------------------------------

responsavel_digital = "Atendimento Fatorial"
data_hoje = '290322'
data_ontem = '280322'


#importando positivador mês M 
caminho_novo = Path(r'captacao_diario\positivador_' + data_hoje + '.xlsx') #! positivador do dia de run do código
posi_novo = pd.read_excel(caminho_novo , skiprows= 2)

# importando positivador mês M - 1 
caminho_velho = Path(r'captacao_diario\arquivos\2022\Fevereiro\positivador_250222.xlsx') #! último positivador do mês anterior
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

#importando lista meta actio
caminho_meta = Path(r'bases_dados\Clientes meta actio.xlsx') #! não precisa mudar
lista_meta = pd.read_excel(caminho_meta)
lista_meta = lista_meta.loc[: , ['Código Cliente' , 'Meta actio']]

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

#diretorios relatorios de captação M-1 e M-2

caminho_m1 = Path(r'captacao_diario\arquivos\2022\Fevereiro\captacao_250222.xlsx') #!
caminho_m2 = Path(r'captacao_diario\arquivos\2022\Janeiro\captacao_310122.xlsx') #!

#parametros de dias úteis do mês

#!#!#!#!#!#!#!#!#!#!#!#!
nm1 = 19 #!
nm2 = 21 #!
nm =  20  #!nº de dias uteis até o positivador

#Nomes dos meses passados

captacao_m1 = 'Total captado Fevereiro'
captacao_m2 = 'Total captado Janeiro'
media_m1 = 'Média diária Fevereiro'
media_m2 = 'Média diária Janeiro'

#diretorio arquivo final

caminho_excel = Path(r'captacao_diario\captacao_' + data_hoje + '.xlsx',index = False) #!

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

#meta actio para os perdidos
tabela_perdidos = pd.merge(tabela_perdidos , lista_meta , left_on = 'Cliente' , right_on = 'Código Cliente' , how = 'left')

tabela_perdidos['Meta actio'].fillna('Não' , inplace = True)

del tabela_perdidos['Código Cliente']

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

#meta actio para os velhos
tabela_velhos = pd.merge(tabela_velhos , lista_meta , left_on = 'Cliente' , right_on = 'Código Cliente' , how = 'left')

tabela_velhos['Meta actio'].fillna('Não' , inplace = True)

del tabela_velhos['Código Cliente']

# montando tabela de clientes novos e transferencias ----------------------------------------------------------------------------------------------------------------------------

#tabela clientes novos + transferencias
tabela_novos_transf = posi_novo.loc[posi_novo['Status conta'] == 'conta nova' , :]


#identificando quais são as transferências
tabela_novos_transf.loc[: ,'Transferência?'] = tabela_novos_transf.loc[: , 'Cliente'].where(tabela_novos_transf.loc[: , 'Cliente'].isin(lista_transf.loc[: , 'Cliente']) == True)

tabela_novos_transf['Transferência?'].fillna('Não' , inplace = True)

tabela_novos_transf.loc[tabela_novos_transf['Transferência?'] !='Não' , 'Transferência?'] = 'Sim'

tabela_novos_transf.loc[tabela_novos_transf['Net em M-1'] > 0 , 'Transferência?'] = 'Sim'

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

#montando dados para de captação ----------------------------------------------------------------------------------------------------------------------------------

#dados contas velhas (com e sem meta actio)
dados_velhos = tabela_velhos.loc[: , ['Cliente' , 'Assessor correto' , 'Captação Líquida em M' , 'Meta actio']]

dados_velhos = dados_velhos.astype({'Assessor correto': str})

dados_velhos_meta = tabela_velhos.loc[(tabela_velhos['Meta actio'] == 'Não') | ((tabela_velhos['Meta actio'] == 'Sim')&(tabela_velhos['Captação Líquida em M']>0)) , ['Cliente' , 'Assessor correto' , 'Captação Líquida em M' , 'Meta actio']]

dados_velhos_meta = dados_velhos_meta.astype({'Assessor correto': str})

#dados contas efetivamente novas
dados_novos = tabela_novos.loc[: , ['Cliente' , 'Assessor correto' , 'Captação Líquida em M']]

dados_novos = dados_novos.astype({'Assessor correto': str})

#dados contas transferências (valor a ser considerado : Soma de NET M-1 e Captação Líquida)
dados_transf = tabela_transf.loc[: , ['Cliente' , 'Assessor correto' , 'Captação Líquida em M' , 'Net em M-1']]

dados_transf['Total transferências'] = dados_transf['Captação Líquida em M'] + dados_transf['Net em M-1']

dados_transf = dados_transf.astype({'Assessor correto': str})

#dados contas perdidas (com e sem meta actio) (valor a ser considerado : Soma de Net em M)
dados_perdidos = tabela_perdidos.loc[: , ['Cliente' , 'Assessor correto' , 'Net Em M' , 'Meta actio']]

dados_perdidos.loc[: , 'Total contas perdidas'] = dados_perdidos['Net Em M'] * -1

dados_perdidos = dados_perdidos.astype({'Assessor correto': str})

dados_perdidos_meta = dados_perdidos.loc[dados_perdidos['Meta actio'] == 'Não' , :]

dados_perdidos_meta.rename(columns = {'Total contas perdidas' : 'Total contas perdidas s/ meta'} , inplace = True)

dados_perdidos_meta = dados_perdidos_meta.astype({'Assessor correto': str})

#transformando os dados em resumos -------------------------------------------------------------------------------------------------------------------------------------------------------

#resumo das contas velhas (com e sem meta actio)
#com meta
resumo_velhos = dados_velhos.loc[: , ['Assessor correto' , 'Captação Líquida em M']].groupby('Assessor correto').sum()

resumo_velhos.rename(columns = {'Captação Líquida em M' : 'Total conta velha'} , inplace = True)

#sem meta
resumo_velhos_meta = dados_velhos_meta.loc[: , ['Assessor correto' , 'Captação Líquida em M']].groupby('Assessor correto').sum()

resumo_velhos_meta.rename(columns = {'Captação Líquida em M' : 'Total conta velha s/ meta'}, inplace = True)

#resumo das contas efetivamente novas
resumo_novos = dados_novos.loc[: , ['Assessor correto' , 'Captação Líquida em M']].groupby('Assessor correto').sum()

resumo_novos.rename(columns = {'Captação Líquida em M' : 'Total conta nova'}, inplace = True)

#resumo das transferências
resumo_transf = dados_transf.loc[: , ['Assessor correto' , 'Total transferências']].groupby('Assessor correto').sum()

#resumo contas perdidas (com meta e sem meta)
#com meta
resumo_perdidos = dados_perdidos.loc[: , ['Assessor correto' , 'Total contas perdidas']].groupby('Assessor correto').sum()

#sem meta
resumo_perdidos_meta = dados_perdidos_meta.loc[: , ['Assessor correto' , 'Total contas perdidas s/ meta']].groupby('Assessor correto').sum()

#construção dos dados da carteira atual (base positivador M) ------------------------------------------------------------------------------------------------------------------------------------------------

#criando dados base
dados_carteira = posi_novo.loc[: , ['Assessor' , 'Cliente' , 'Net Em M']]

#inserindo assessores corretos 
dados_carteira = pd.merge(dados_carteira , clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how='left')

dados_carteira['Assessor Relacionamento'].fillna(dados_carteira['Assessor'] , inplace = True)

dados_carteira.rename(columns = {'Assessor Relacionamento' : 'Assessor correto'} , inplace = True)

del dados_carteira['Conta']

dados_carteira.loc[dados_carteira['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

dados_carteira = dados_carteira.astype({'Assessor correto': str})

#retirando os clientes zerados da base

dados_carteira = dados_carteira.loc[dados_carteira['Net Em M'] != 0 , :]

#achando a quantidade de clientes que cada assessor possui

dados_carteira_qtd = dados_carteira.loc[: , ['Assessor correto' , 'Cliente']].groupby('Assessor correto').count()

#achando o somatório de NET para cada assessor

dados_carteira_net = dados_carteira.loc[: , ['Assessor correto' , 'Net Em M']].groupby('Assessor correto').sum()

#juntando as informações e achando o ticket médio

dados_carteira_assessor = pd.merge(dados_carteira_qtd , dados_carteira_net , left_index=True , right_index=True , how= 'left')

dados_carteira_assessor.loc[: , 'Ticket Médio'] = dados_carteira_assessor['Net Em M'] / dados_carteira_assessor['Cliente']

dados_carteira_assessor.rename(columns={'Cliente' : 'Qtd Clientes n/ zerados'} , inplace=True)

#montando dados sob o ponto de vista da XP

dados_carteira_xp = posi_novo.loc[posi_novo['Net Em M'] != 0 , ['Assessor' , 'Cliente' , 'Net Em M']]

dados_carteira_xp = dados_carteira_xp.loc[: , ['Assessor' , 'Cliente']].groupby('Assessor').count()

variavel = posi_novo[['Assessor' , 'Net Em M']].groupby('Assessor').sum()

dados_carteira_xp = pd.merge(dados_carteira_xp , posi_novo[['Assessor' , 'Net Em M']].groupby('Assessor').sum() , how='left' , on='Assessor')

dados_carteira_xp.rename(columns={'Cliente':'Qtd Clientes XP' , 'Net Em M' : 'NET XP'} , inplace=True)

dados_carteira_xp.reset_index(inplace=True)

dados_carteira_xp = dados_carteira_xp.astype({'Assessor':str})

#montargem do tabelão de captação -------------------------------------------------------------------------------------------------------------------------------

tabela_captacao = pd.merge(assessores, resumo_velhos , left_on='Código assessor' , right_index= True , how= 'left')

tabela_captacao = pd.merge(tabela_captacao , resumo_novos , left_on='Código assessor' , right_index= True , how = 'left')

tabela_captacao = pd.merge(tabela_captacao , resumo_transf , left_on = 'Código assessor' , right_index= True , how = 'left')

tabela_captacao = pd.merge(tabela_captacao , resumo_perdidos , left_on = 'Código assessor' , right_index = True , how = 'left')

tabela_captacao['Captação Líquida'] = tabela_captacao.sum(axis = 1)

tabela_captacao = pd.merge(tabela_captacao , dados_carteira_assessor , left_on='Código assessor' , right_index=True , how='left')

tabela_captacao = pd.merge(tabela_captacao , dados_carteira_xp , how='left' , left_on='Código assessor' , right_on='Assessor')

del tabela_captacao['Assessor']

for n in tabela_captacao.columns:
    tabela_captacao[n].fillna(0 , inplace = True)

#Linha total da tabela captação assessores

tabela_captacao.set_index('Código assessor', inplace = True)

tabela_captacao.loc['Total Fatorial'] = tabela_captacao[['Total conta velha' , 'Total conta nova' , 'Total transferências' , 'Total contas perdidas' , 'Captação Líquida','Qtd Clientes n/ zerados','Net Em M' , 'Qtd Clientes XP','NET XP']].sum()

tabela_captacao.reset_index(inplace=True)

tabela_captacao['Ticket Médio'].iloc[-1] = tabela_captacao['Net Em M'].iloc[-1] / tabela_captacao['Qtd Clientes n/ zerados'].iloc[-1]

tabela_captacao.fillna('Total Fatorial',inplace=True)

#montagem por células

tabela_celulas = tabela_captacao.loc[: , tabela_captacao.columns.isin(['Nome assessor','Ticket Médio']) == False].groupby('Time').sum()

tabela_celulas.loc[: , 'Ticket Médio'] = tabela_celulas['Net Em M'] / tabela_celulas['Qtd Clientes n/ zerados']

qtd_assessores = assessores.loc[: , ['Nome assessor','Time']].groupby('Time').count()

tabela_celulas = pd.merge(tabela_celulas , qtd_assessores , left_index=True , right_index=True , how='left')

tabela_celulas.loc['Total Fatorial','Nome assessor'] = tabela_celulas['Nome assessor'].sum()

tabela_celulas['Captação / Assessor'] = tabela_celulas['Captação Líquida'] / tabela_celulas['Nome assessor']

tabela_celulas = tabela_celulas.loc[: , ['Total conta velha' , 'Total conta nova' , 'Total transferências' , 'Total contas perdidas' , 'Captação Líquida' , 'Captação / Assessor' , 'Qtd Clientes n/ zerados' , 'Net Em M' , 'Ticket Médio', 'Qtd Clientes XP','NET XP']]

tabela_celulas.reset_index(inplace=True)

#montargem do tabelão de captação meta -------------------------------------------------------------------------------------------------------------------------------

tabela_captacao_meta = pd.merge(assessores, resumo_velhos_meta , left_on='Código assessor' , right_index= True , how= 'left')

tabela_captacao_meta = pd.merge(tabela_captacao_meta , resumo_novos , left_on='Código assessor' , right_index= True , how = 'left')

tabela_captacao_meta = pd.merge(tabela_captacao_meta , resumo_transf , left_on = 'Código assessor' , right_index= True , how = 'left')

tabela_captacao_meta = pd.merge(tabela_captacao_meta , resumo_perdidos_meta , left_on = 'Código assessor' , right_index = True , how = 'left')

tabela_captacao_meta['Captação Líquida'] = tabela_captacao_meta.sum(axis = 1)

tabela_captacao_meta = pd.merge(tabela_captacao_meta , dados_carteira_assessor , left_on='Código assessor' , right_index=True , how='left')

tabela_captacao_meta = pd.merge(tabela_captacao_meta , dados_carteira_xp , how='left' , left_on='Código assessor' , right_on='Assessor')

del tabela_captacao_meta['Assessor']

for n in tabela_captacao_meta.columns:
    tabela_captacao_meta[n].fillna(0 , inplace = True)

#Linha total da tabela captação assessores

tabela_captacao_meta.set_index('Código assessor', inplace = True)

tabela_captacao_meta.loc['Total Fatorial'] = tabela_captacao_meta[['Total conta velha s/ meta' , 'Total conta nova' , 'Total transferências' , 'Total contas perdidas s/ meta' , 'Captação Líquida','Qtd Clientes n/ zerados','Net Em M' , 'Qtd Clientes XP','NET XP']].sum()

tabela_captacao_meta.reset_index(inplace=True)

tabela_captacao_meta['Ticket Médio'].iloc[-1] = tabela_captacao_meta['Net Em M'].iloc[-1] / tabela_captacao_meta['Qtd Clientes n/ zerados'].iloc[-1]

tabela_captacao_meta.fillna('Total Fatorial',inplace=True)

#montagem por células

tabela_celulas_meta = tabela_captacao_meta.loc[: , tabela_captacao_meta.columns.isin(['Nome assessor','Ticket Médio']) == False].groupby('Time').sum()

tabela_celulas_meta.loc[: , 'Ticket Médio'] = tabela_celulas_meta['Net Em M'] / tabela_celulas_meta['Qtd Clientes n/ zerados']

tabela_celulas_meta = pd.merge(tabela_celulas_meta , qtd_assessores , left_index=True , right_index=True , how='left')

tabela_celulas_meta.loc['Total Fatorial','Nome assessor'] = tabela_celulas_meta['Nome assessor'].sum()

tabela_celulas_meta['Captação / Assessor'] = tabela_celulas_meta['Captação Líquida'] / tabela_celulas_meta['Nome assessor']

tabela_celulas_meta = tabela_celulas_meta.loc[: , ['Total conta velha s/ meta' , 'Total conta nova' , 'Total transferências' , 'Total contas perdidas s/ meta' , 'Captação Líquida' , 'Captação / Assessor' , 'Qtd Clientes n/ zerados' , 'Net Em M' , 'Ticket Médio', 'Qtd Clientes XP','NET XP']]

tabela_celulas_meta.reset_index(inplace=True)


#montar top aportes e saques -----------------------------------------------------------------------------------------------------------------------------------

#criando coluna de assessores corretos no positivador novo

posi_novo = pd.merge(posi_novo , clientes_rodrigo , left_on = 'Cliente' , right_on = 'Conta' , how = 'left')

posi_novo['Assessor Relacionamento'].fillna(posi_novo['Assessor'] , inplace = True)

posi_novo.rename(columns={'Assessor Relacionamento': 'Assessor correto'} , inplace=True)

del posi_novo['Conta']

posi_novo.loc[posi_novo['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

posi_novo = posi_novo.astype({'Assessor correto': str})

#montando tabela (código dos assessores e seus nomes repetidos n vezes ou a quantidade máxima)

n = 10 #número de elementos no ranking

#aportes
aporte = pd.DataFrame()

for i in range(assessores.loc[: , 'Código assessor'].count()):
    aporte_n = posi_novo.loc[posi_novo['Assessor correto'] == assessores.loc[: , 'Código assessor'][i] , ['Assessor correto','Cliente','Captação Bruta em M' , 'Captação Líquida em M']]

    aporte_n.sort_values('Captação Bruta em M' , ascending = False , inplace = True)

    aporte_n = aporte_n.reset_index(drop = True)

    if aporte_n['Assessor correto'].count() < n:
        aporte_n = aporte_n.iloc[range(aporte_n['Assessor correto'].count()),:]
    else:
        aporte_n = aporte_n.iloc[range(n)]

    aporte = aporte.append(aporte_n)

    aporte = aporte.reset_index(drop=True)

aporte.rename(columns={'Assessor correto':'Código assessor - aportes','Cliente':'Código do cliente - aportes','Captação Bruta em M':'Valor do aporte' , 'Captação Líquida em M':'Valor líquido final - aporte'} , inplace=True)

#saques
saque = pd.DataFrame()

for i in range(assessores.loc[: , 'Código assessor'].count()):
    saque_n = posi_novo.loc[posi_novo['Assessor correto'] == assessores.loc[: , 'Código assessor'][i] , ['Assessor correto','Cliente','Resgate em M' , 'Captação Líquida em M']]

    saque_n.sort_values('Resgate em M' , ascending = True , inplace = True)

    saque_n = saque_n.reset_index(drop = True)

    if saque_n['Assessor correto'].count() < n:
        saque_n = saque_n.iloc[range(saque_n['Assessor correto'].count()),:]
    else:
        saque_n = saque_n.iloc[range(n)]

    saque = saque.append(saque_n)

    saque = saque.reset_index(drop=True)

saque.rename(columns={'Assessor correto':'Código assessor - saques','Cliente':'Código do cliente - saques','Resgate em M':'Valor do saque' , 'Captação Líquida em M':'Valor líquido final - saques'} , inplace=True)

#juntando os aportes e saques

aporte_saque = pd.concat([aporte , saque] , axis=1)

aporte_saque = pd.merge(aporte_saque , assessores.loc[: , ['Código assessor' , 'Nome assessor']] , left_on='Código assessor - aportes' , right_on='Código assessor' , how='left')

aporte_saque = aporte_saque.loc[: , ['Nome assessor' , 'Código assessor' , 'Código do cliente - aportes' , 'Valor do aporte' , 'Valor líquido final - aporte' , 'Código do cliente - saques' , 'Valor do saque' , 'Valor líquido final - saques']]

aporte_saque.rename(columns = {'Nome assessor':'Assessor'}, inplace=True)

# clientes novos no dia D ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#pegando a lista acumuladados clientes novos que entraram

clientes_novos_acum = tabela_novos_transf.loc[: , ['Cliente' , 'Assessor' ,'Aplicação Financeira Declarada' , 'Data de Nascimento' ,'Transferência?']]

#cruzando com os clientes rodrigo

clientes_novos_acum = pd.merge(clientes_novos_acum , clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how='left')

clientes_novos_acum['Assessor Relacionamento'].fillna(clientes_novos_acum['Assessor'] , inplace = True)

clientes_novos_acum.rename(columns = {'Assessor Relacionamento' : 'Assessor correto'} , inplace = True)

del clientes_novos_acum['Conta']

clientes_novos_acum.loc[clientes_novos_acum['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

clientes_novos_acum = clientes_novos_acum.astype({'Assessor correto': str})

#nome do cliente

clientes_novos_acum = pd.merge(clientes_novos_acum ,suitability[['CodigoBolsa' , 'NomeCliente']] , how='left' , left_on='Cliente' , right_on='CodigoBolsa' )

del clientes_novos_acum['CodigoBolsa']

clientes_novos_acum.rename(columns={'NomeCliente':'Nome'},inplace=True)

clientes_novos_acum['Nome'].fillna('Não encontrado' , inplace=True)


# lista de clientes novos acumulados no relatório D-1

clientes_novos_ontem = clientes_novos_ontem.loc[: , ['Cliente' , 'Transferência?']]

clientes_novos_ontem['Presente'] = 'Ontem'

# cruzando os clientes novos de ontem e os de hoje

clientes_novos_acum = pd.merge(clientes_novos_acum , clientes_novos_ontem[['Cliente' , 'Presente']] , how='left' , on='Cliente')

clientes_novos_acum['Presente'].fillna('Hoje' , inplace=True)

#colocando nome dos assessores

clientes_novos_acum = pd.merge(clientes_novos_acum , assessores[['Código assessor','Nome assessor']] , how='left' , left_on='Assessor correto' , right_on='Código assessor')

#colocando a profissão do cliente

clientes_novos_acum = pd.merge(clientes_novos_acum , posi_novo[['Cliente','Profissão']] , how='left' , on='Cliente')

# filtrando quais são os clientes de hoje

clientes_novos_hj = clientes_novos_acum.loc[clientes_novos_acum['Presente']=='Hoje' , ['Cliente' , 'Nome' , 'Profissão' , 'Data de Nascimento' ,'Assessor correto' , 'Nome assessor' , 'Aplicação Financeira Declarada']]

clientes_novos_hj.rename(columns = {'Assessor correto':'Assessor' , 'Aplicação Financeira Declarada':'PL Declarado'} , inplace=True)

clientes_novos_hj.reset_index(inplace=True)

#cockpit assessores ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#arquivo meses anteriores (com transformação da coluna de códigos em strings)

arquivo_m1 = pd.read_excel(caminho_m1 , sheet_name='Resumo')

arquivo_m1 = arquivo_m1.astype({'Código assessor':str})

arquivo_m2 = pd.read_excel(caminho_m2 , sheet_name='Resumo')

arquivo_m2= arquivo_m2.astype({'Código assessor':str})

#montando a tabela de cockpit

#coluna de captação total no mês

cockpit = assessores

cockpit = pd.merge(cockpit , tabela_captacao[['Código assessor' , 'Captação Líquida']] , how='left' , on='Código assessor')

cockpit.rename(columns={'Captação Líquida':'Total captado nesse mês'} , inplace=True)

cockpit.drop_duplicates(inplace=True)

# coluna de captação da célula

cockpit = pd.merge(cockpit , tabela_celulas[['Time' , 'Captação Líquida']] , how='left' , on='Time')

cockpit.rename(columns={'Captação Líquida':'Total célula'} , inplace=True)

# coluna de captação da Fatorial

cockpit['Total Fatorial'] = tabela_celulas.loc [tabela_celulas["Time"]=="Total Fatorial","Captação Líquida"].values.sum ()
# coluna total captado mês m1

cockpit = pd.merge(cockpit , arquivo_m1[['Código assessor' , 'Captação Líquida']] , how='left' , on='Código assessor')

cockpit.rename(columns={'Captação Líquida':captacao_m1} , inplace=True)

# coluna total captado mês m2

cockpit = pd.merge(cockpit , arquivo_m2[['Código assessor' , 'Captação Líquida']] , how='left' , on='Código assessor')

cockpit.rename(columns={'Captação Líquida':captacao_m2} , inplace=True)

#coluna média diária desse mês

cockpit['Média diária nesse mês'] = cockpit['Total captado nesse mês']/nm

#coluna média diária da célula desse mês

cockpit['Média diária célula'] = cockpit['Total célula']/nm

#coluna média diária da Fatorial desse mês

cockpit['Média diária Fatorial'] = cockpit['Total Fatorial']/nm

#coluna média diária do mês m1

cockpit[media_m1] = cockpit[captacao_m1]/nm1

#coluna média diária do mês m2

cockpit[media_m2] = cockpit[captacao_m2] / nm2

#coluna média diária da captação per capta da célula

cockpit = pd.merge(cockpit , tabela_celulas[['Time' , 'Captação / Assessor']] , how='left' , on='Time')

cockpit['Média diária célula / assessor'] = cockpit['Captação / Assessor']/nm

#deletando as colunas "código assessor" e "Captação / Assessor"

del cockpit['Código assessor']

del cockpit['Captação / Assessor']

# colocando "0" para os valores NaN

for c in cockpit.columns:
    cockpit[c].fillna(0 , inplace=True)


#montar excel final -----------------------------------------------------------------------------------------------------------------------------------

lista_tabelas = [
(tabela_captacao, 'Resumo', 'Table Style Medium 2'), 
(tabela_captacao_meta, 'Resumo meta', 'Table Style Medium 2'), 
(tabela_celulas, 'Resumo times', 'Table Style Medium 2'),
(tabela_celulas_meta, 'Resumo times meta', 'Table Style Medium 2'), 
(aporte_saque, 'Lista top 10', 'Table Style Medium 2'), 
(cockpit, 'Cockpit assessores', 'Table Style Medium 2'),
(tabela_celulas, 'Resumo times', 'Table Style Medium 2'),
(tabela_velhos, 'Contas velhas', 'Table Style Medium 2'), 
(tabela_novos, 'Contas novas', 'Table Style Medium 2'), 
(tabela_transf, 'Transferêcnias', 'Table Style Medium 2'),
(tabela_novos_transf, 'Novos + Transf', 'Table Style Medium 2'), 
(clientes_novos_hj, 'Clientes novos D', 'Table Style Medium 2'), 
(clientes_novos_acum, 'Clientes novos total', 'Table Style Medium 2'),
(tabela_perdidos, 'Perdidos', 'Table Style Medium 2'), 
(posi_novo, 'Positivador M', 'Table Style Medium 2'), 
(dados_velhos, 'Dados velhos', 'Table Style Medium 2'),
(dados_novos, 'Dados Novos', 'Table Style Medium 2'), 
(dados_transf, 'Dados transferidos', 'Table Style Medium 2'), 
(dados_perdidos, 'Dados perdidos', 'Table Style Medium 2'),
]

# Com a lista de tabelas, gera as abas necessárias

writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy') # pylint: disable=abstract-class-instantiated

for tabela, nome_tabela, *_ in lista_tabelas:
    tabela.to_excel(writer , sheet_name= nome_tabela , index= False)

# Formatação das tabelas

for tabela, nome_aba, estilo in lista_tabelas:

    arquivo = writer.book
    aba = writer.sheets[nome_aba]
    
    colunas = [{'header':column} for column in tabela.columns ]
    (lin, col) = tabela.shape

    aba.add_table(0 , 0 , lin , col-1 , {
        'columns': colunas,
        'style': estilo, 
        'autofilter': False
        })

    #if nome_aba == 'Resumo' or 'Resumo times':

        # Formata a ultima linha como negrito

        #last_line = tabela.tail(1).fillna('.').values.tolist()

        #cell_format = arquivo.add_format()
        #cell_format.set_bold()

        #list_index_col = range(col)
        
        #for i, column in enumerate( list_index_col ):
            #aba.write(lin, column, last_line[0][i], cell_format)
            

writer.save()
print ("arquivo criado")