import pandas as pd
from pathlib import Path

# planilhas usadas: positivador final do mês M-1 ; positivador mês M ; planilha transferências ; clientes rodrigo ; Assessores ; relatório captação mês m-1 ; relatório captação mês m-2 ; relatório captação d-1 ; suitability; transferidso D-1; Novos+Trasf D-1

#importando as planilhas utilizadas ----------------------------------------------------------------------------------------------------------------------------

responsavel_digital = "Atendimento Fatorial"
data_hoje = '190422'
data_ontem = '140422'


#importando positivador mês M 
caminho_novo = Path(r'captacao_diario\positivador_' + data_hoje + '.xlsx') #! positivador do dia de run do código
posi_novo = pd.read_excel(caminho_novo , skiprows= 2)

# importando positivador mês M - 1 
caminho_velho = Path(r'captacao_diario\arquivos\2022\Março\positivador_310322.xlsx') #! último positivador do mês anterior
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

# montando tabela de clientes velhos ----------------------------------------------------------------------------------------------------------------------------

#encontrando quais são os clientes velhos
#posi_novo['Status conta'] = posi_novo['Cliente'].where(posi_novo['Cliente'].isin(posi_velho['Cliente']) == True)

posi_novo.loc [posi_novo["Cliente"].isin (posi_velho['Cliente']),"Status conta"] = 'conta velha'

posi_novo['Status conta'].fillna('conta nova' , inplace = True)

posi_novo.drop_duplicates (subset = "Cliente",inplace=True)

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

#dados contas velhas
dados_velhos = tabela_velhos.loc[: , ['Cliente' , 'Assessor correto' , 'Captação Líquida em M']]

dados_velhos = dados_velhos.astype({'Assessor correto': str})

#dados contas efetivamente novas
dados_novos = tabela_novos.loc[: , ['Cliente' , 'Assessor correto' , 'Captação Líquida em M']]

dados_novos = dados_novos.astype({'Assessor correto': str})

#dados contas transferências (valor a ser considerado : Soma de NET M-1 e Captação Líquida)
dados_transf = tabela_transf.loc[: , ['Cliente' , 'Assessor correto' , 'Captação Líquida em M' , 'Net em M-1']]

dados_transf['Total transferências'] = dados_transf['Captação Líquida em M'] + dados_transf['Net em M-1']

dados_transf = dados_transf.astype({'Assessor correto': str})

#dados contas perdidas (valor a ser considerado : Soma de Net em M)
dados_perdidos = tabela_perdidos.loc[: , ['Cliente' , 'Assessor correto' , 'Net Em M' ]]

dados_perdidos.loc[: , 'Total contas perdidas'] = dados_perdidos['Net Em M'] * -1

dados_perdidos = dados_perdidos.astype({'Assessor correto': str})

#transformando os dados em resumos -------------------------------------------------------------------------------------------------------------------------------------------------------

#resumo das contas velhas
resumo_velhos = dados_velhos.loc[: , ['Assessor correto' , 'Captação Líquida em M']].groupby('Assessor correto').sum()

resumo_velhos.rename(columns = {'Captação Líquida em M' : 'Total conta velha'} , inplace = True)

#resumo das contas efetivamente novas
resumo_novos = dados_novos.loc[: , ['Assessor correto' , 'Captação Líquida em M']].groupby('Assessor correto').sum()

resumo_novos.rename(columns = {'Captação Líquida em M' : 'Total conta nova'}, inplace = True)

#resumo das transferências
resumo_transf = dados_transf.loc[: , ['Assessor correto' , 'Total transferências']].groupby('Assessor correto').sum()

#resumo contas perdidas
resumo_perdidos = dados_perdidos.loc[: , ['Assessor correto' , 'Total contas perdidas']].groupby('Assessor correto').sum()

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

#criando coluna de assessores corretos no positivador novo

posi_novo = pd.merge(posi_novo , clientes_rodrigo , left_on = 'Cliente' , right_on = 'Conta' , how = 'left')

posi_novo['Assessor Relacionamento'].fillna(posi_novo['Assessor'] , inplace = True)

posi_novo.rename(columns={'Assessor Relacionamento': 'Assessor correto'} , inplace=True)

del posi_novo['Conta']

posi_novo.loc[posi_novo['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

posi_novo = posi_novo.astype({'Assessor correto': str})

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

#montar excel final -----------------------------------------------------------------------------------------------------------------------------------

lista_tabelas = [
(tabela_captacao, 'Resumo', 'Table Style Medium 2'),
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