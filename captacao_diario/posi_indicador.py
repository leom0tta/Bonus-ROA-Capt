import pandas as pd
import sys
sys.path.insert(1, r'.\Funções')
from funcoes import export_files_captacao, gera_excel

data_hoje = '110822'
data_ontem = '090822'

dataframes = export_files_captacao(data_hoje, data_ontem)

posi_novo = dataframes[0]
posi_d1 = dataframes[1]
posi_velho = dataframes[2]
clientes_rodrigo = dataframes[3]
assessores = dataframes[4]
lista_transf = dataframes[5]
clientes_novos_ontem = dataframes[6]
suitability = dataframes[7]
registro_transf = dataframes[8]

def relatorio_diario(posi_novo, posi_velho, clientes_rodrigo, assessores, lista_transf, clientes_novos_ontem, suitability, gera_excel=True, caminho_excel = r'captacao_diario\relatorio_diario_' + data_hoje + '.xlsx'):
    """Tanto separando os assessores em times, quanto analisando individualmente, esse código gera um
    relatório diário, com base nos arquivos da XP, reportando a captação e o AUM dos assessores, pela ótica da XP e do Indicador."""

    import pandas as pd

    #diretorio arquivo final

    clientes_rodrigo = clientes_rodrigo.loc[:, ['Conta', 'Assessor Indicador']]
    clientes_rodrigo['Assessor Indicador'] = clientes_rodrigo['Assessor Indicador'].str.lstrip('A')

    #montando tabela clientes perdidos ----------------------------------------------------------------------------------------------------------------------------

    #encontrando quais são os clientes perdidos
    posi_velho['Clientes perdidos'] = posi_velho['Cliente'].where(posi_velho['Cliente'].isin(posi_novo['Cliente']) == True)

    #renomeando os valores: na -> "Saiu" ; código cliente -> "Permanece"
    posi_velho['Clientes perdidos'].fillna('Saiu' , inplace = True)

    posi_velho.loc[posi_velho['Clientes perdidos'] != 'Saiu' , 'Clientes perdidos'] = 'Permanece'

    #montagem do dataframe
    tabela_perdidos = posi_velho.loc[posi_velho['Clientes perdidos'] == 'Saiu' , :]

    #criação de coluna de assessor Indicador
    tabela_perdidos = tabela_perdidos.merge(clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how= 'left')

    tabela_perdidos['Assessor Indicador'].fillna(tabela_perdidos['Assessor'] , inplace = True)

    del tabela_perdidos['Conta']

    tabela_perdidos.loc[tabela_perdidos['Assessor Indicador'] == '1618' , 'Assessor Indicador'] = 'Atendimento Fatorial'

    tabela_perdidos['Net Indicador'] = tabela_perdidos['Net Em M']

    del tabela_perdidos['Net Em M']

    # montando tabela de clientes velhos ----------------------------------------------------------------------------------------------------------------------------

    #encontrando quais são os clientes velhos

    posi_novo.loc [posi_novo["Cliente"].isin (posi_velho['Cliente']),"Status conta"] = 'conta velha'

    posi_novo['Status conta'].fillna('conta nova' , inplace = True)

    posi_novo.drop_duplicates (subset = "Cliente",inplace=True)

    #seleção das contas velhas
    tabela_velhos = posi_novo.loc[posi_novo['Status conta'] == 'conta velha' , :]

    #criação da coluna de assessor indicador
    tabela_velhos = pd.merge(tabela_velhos , clientes_rodrigo , left_on= 'Cliente' , right_on= 'Conta' , how= 'left')

    tabela_velhos['Assessor Indicador'].fillna(tabela_velhos['Assessor'] , inplace = True)

    del tabela_velhos['Conta']

    tabela_velhos.loc[tabela_velhos['Assessor Indicador'] == '1618' , 'Assessor Indicador'] = 'Atendimento Fatorial'

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

    tabela_novos['Assessor Indicador'].fillna(tabela_novos['Assessor'] , inplace = True)

    del tabela_novos['Conta']

    tabela_novos.loc[tabela_novos['Assessor Indicador'] == '1618' , 'Assessor Indicador'] = 'Atendimento Fatorial'

    #tabela transferências
    tabela_transf = tabela_novos_transf.loc[tabela_novos_transf['Transferência?'] == 'Sim' , :]

    #cliente rodrigo para as transferências
    tabela_transf = pd.merge(tabela_transf , clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how='left')

    tabela_transf['Assessor Indicador'].fillna(tabela_transf['Assessor'] , inplace = True)

    del tabela_transf['Conta']

    tabela_transf.loc[tabela_transf['Assessor Indicador'] == '1618' , 'Assessor Indicador'] = 'Atendimento Fatorial'

    #montando dados para de captação ----------------------------------------------------------------------------------------------------------------------------------

    #dados contas velhas
    dados_velhos = tabela_velhos.loc[: , ['Cliente' , 'Assessor Indicador', 'Assessor' , 'Captação Líquida em M']]

    dados_velhos = dados_velhos.astype({'Assessor Indicador': str, 'Assessor': str})

    #dados contas efetivamente novas
    dados_novos = tabela_novos.loc[: , ['Cliente' , 'Assessor Indicador', 'Assessor' , 'Captação Líquida em M']]

    dados_novos = dados_novos.astype({'Assessor Indicador': str, 'Assessor': str})

    #dados contas transferências (valor a ser considerado : Soma de NET M-1 e Captação Líquida)
    dados_transf = tabela_transf.loc[: , ['Cliente' , 'Assessor Indicador', 'Assessor' , 'Captação Líquida em M' , 'Net em M-1']]

    dados_transf['Total transferências'] = dados_transf['Captação Líquida em M'] + dados_transf['Net em M-1']

    dados_transf = dados_transf.astype({'Assessor Indicador': str, 'Assessor': str})

    #dados contas perdidas (valor a ser considerado : Soma de Net em M)
    dados_perdidos = tabela_perdidos.loc[: , ['Cliente' , 'Assessor Indicador', 'Assessor' , 'Net Indicador' ]]

    dados_perdidos.loc[: , 'Total contas perdidas'] = dados_perdidos['Net Indicador'] * -1

    dados_perdidos = dados_perdidos.astype({'Assessor Indicador': str, 'Assessor': str})

    #transformando os dados em resumos -------------------------------------------------------------------------------------------------------------------------------------------------------

    #resumo das contas velhas
    resumo_velhos_indicador = dados_velhos.loc[: , ['Assessor Indicador' , 'Captação Líquida em M']].groupby('Assessor Indicador').sum()
    resumo_velhos_indicador.rename(columns = {'Captação Líquida em M' : 'Total conta velha Indic.'} , inplace = True)
    
    resumo_velhos_xp = dados_velhos.loc[: , ['Assessor' , 'Captação Líquida em M']].groupby('Assessor').sum()
    resumo_velhos_xp.rename(columns = {'Captação Líquida em M' : 'Total conta velha XP'} , inplace = True)

    #resumo das contas efetivamente novas
    resumo_novos_indicador = dados_novos.loc[: , ['Assessor Indicador' , 'Captação Líquida em M']].groupby('Assessor Indicador').sum()
    resumo_novos_indicador.rename(columns = {'Captação Líquida em M' : 'Total conta nova Indic.'}, inplace = True)

    resumo_novos_xp = dados_novos.loc[: , ['Assessor' , 'Captação Líquida em M']].groupby('Assessor').sum()
    resumo_novos_xp.rename(columns = {'Captação Líquida em M' : 'Total conta nova XP'}, inplace = True)

    #resumo das transferências
    resumo_transf_indicador = dados_transf.loc[: , ['Assessor Indicador' , 'Total transferências']].groupby('Assessor Indicador').sum()
    resumo_transf_indicador.rename(columns={'Total transferências': 'Total Transferências Indic.'}, inplace=True)
    
    resumo_transf_xp = dados_transf.loc[: , ['Assessor' , 'Total transferências']].groupby('Assessor').sum()
    resumo_transf_xp.rename(columns={'Total transferências': 'Total Transferências XP'}, inplace=True)

    #resumo contas perdidas
    resumo_perdidos_indicador = dados_perdidos.loc[: , ['Assessor Indicador' , 'Total contas perdidas']].groupby('Assessor Indicador').sum()
    resumo_perdidos_indicador.rename( columns={'Total contas perdidas' : 'Total contas perdidas Indic.'} , inplace=True)

    resumo_perdidos_xp = dados_perdidos.loc[: , ['Assessor' , 'Total contas perdidas']].groupby('Assessor').sum()
    resumo_perdidos_xp.rename( columns={'Total contas perdidas' : 'Total contas perdidas XP'} , inplace=True)

    #construção dos dados da carteira atual (base positivador M) ------------------------------------------------------------------------------------------------------------------------------------------------

    #criando dados base
    dados_carteira = posi_novo.loc[: , ['Assessor' , 'Cliente' , 'Net Em M']]
    dados_carteira.rename(columns = {'Net Em M':'Net Indicador'}, inplace=True)

    #inserindo assessores Indicador 
    dados_carteira = pd.merge(dados_carteira , clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how='left')

    dados_carteira['Assessor Indicador'].fillna(dados_carteira['Assessor'] , inplace = True)

    del dados_carteira['Conta']

    dados_carteira.loc[dados_carteira['Assessor Indicador'] == '1618' , 'Assessor Indicador'] = 'Atendimento Fatorial'

    dados_carteira = dados_carteira.astype({'Assessor Indicador': str})

    #retirando os clientes zerados da base

    dados_carteira = dados_carteira.loc[dados_carteira['Net Indicador'] != 0 , :]

    #achando a quantidade de clientes que cada assessor possui

    dados_carteira_qtd = dados_carteira.loc[: , ['Assessor Indicador' , 'Cliente']].groupby('Assessor Indicador').count()

    #achando o somatório de NET para cada assessor

    dados_carteira_net = dados_carteira.loc[: , ['Assessor Indicador' , 'Net Indicador']].groupby('Assessor Indicador').sum()

    #juntando as informações e achando o ticket médio

    dados_carteira_assessor = pd.merge(dados_carteira_qtd , dados_carteira_net , left_index=True , right_index=True , how= 'left')

    dados_carteira_assessor.loc[: , 'Ticket Médio'] = dados_carteira_assessor['Net Indicador'] / dados_carteira_assessor['Cliente']

    dados_carteira_assessor.rename(columns={'Cliente' : 'Qtd Clientes Indicados'} , inplace=True)

    #montando dados sob o ponto de vista da XP

    dados_carteira_xp = posi_novo.loc[posi_novo['Net Em M'] != 0 , ['Assessor' , 'Cliente' , 'Net Em M']]
    
    dados_carteira_xp.rename(columns={'Net Em M':'Net Indicador'}, inplace=True)

    dados_carteira_xp = dados_carteira_xp.loc[: , ['Assessor' , 'Cliente']].groupby('Assessor').count()

    dados_carteira_xp = pd.merge(dados_carteira_xp , posi_novo[['Assessor' , 'Net Em M']].groupby('Assessor').sum() , how='left' , on='Assessor')

    dados_carteira_xp.rename(columns={'Cliente':'Qtd Clientes XP' , 'Net Em M' : 'NET XP'} , inplace=True)

    dados_carteira_xp.reset_index(inplace=True)

    dados_carteira_xp = dados_carteira_xp.astype({'Assessor':str})

    #montargem do tabelão de captação -------------------------------------------------------------------------------------------------------------------------------

    tabela_captacao = pd.merge(assessores, resumo_velhos_indicador , left_on='Código assessor' , right_index= True , how= 'left')

    tabela_captacao = pd.merge(tabela_captacao , resumo_novos_indicador , left_on='Código assessor' , right_index= True , how = 'left')

    tabela_captacao = pd.merge(tabela_captacao , resumo_transf_indicador , left_on = 'Código assessor' , right_index= True , how = 'left')

    tabela_captacao = pd.merge(tabela_captacao , resumo_perdidos_indicador , left_on = 'Código assessor' , right_index = True , how = 'left')

    tabela_captacao['Captação Indicador'] = tabela_captacao.sum(axis = 1)

    tabela_captacao = pd.merge(tabela_captacao, resumo_velhos_xp , left_on='Código assessor' , right_index= True , how= 'left')

    tabela_captacao = pd.merge(tabela_captacao , resumo_novos_xp , left_on='Código assessor' , right_index= True , how = 'left')

    tabela_captacao = pd.merge(tabela_captacao , resumo_transf_xp , left_on = 'Código assessor' , right_index= True , how = 'left')

    tabela_captacao = pd.merge(tabela_captacao , resumo_perdidos_xp , left_on = 'Código assessor' , right_index = True , how = 'left')

    tabela_captacao['Captação XP'] = tabela_captacao[['Total conta velha XP' , 'Total conta nova XP' , 'Total Transferências XP' , 'Total contas perdidas XP']].sum(axis = 1)

    tabela_captacao = pd.merge(tabela_captacao , dados_carteira_assessor , left_on='Código assessor' , right_index=True , how='left')

    tabela_captacao = pd.merge(tabela_captacao , dados_carteira_xp , how='left' , left_on='Código assessor' , right_on='Assessor')

    del tabela_captacao['Assessor']

    for n in tabela_captacao.columns:
        tabela_captacao[n].fillna(0 , inplace = True)

    #Linha total da tabela captação assessores

    tabela_captacao.set_index('Código assessor', inplace = True)

    tabela_captacao.loc['Total Fatorial'] = tabela_captacao[['Total conta velha Indic.' , 'Total conta nova Indic.' , 'Total Transferências Indic.' , 'Total contas perdidas Indic.' , 'Captação Indicador', 'Total conta velha XP' , 'Total conta nova XP' , 'Total Transferências XP' , 'Total contas perdidas XP' , 'Captação XP', 'Qtd Clientes Indicados','Net Indicador' , 'Qtd Clientes XP','NET XP']].sum()

    tabela_captacao.reset_index(inplace=True)

    tabela_captacao['Ticket Médio'].iloc[-1] = tabela_captacao['Net Indicador'].iloc[-1] / tabela_captacao['Qtd Clientes Indicados'].iloc[-1]

    tabela_captacao.fillna('Total Fatorial',inplace=True)
    
    #Tira do resumo quem não tem patrimonio

    mask_sem_net = (tabela_captacao['Net Indicador'] == 0) & (tabela_captacao['NET XP'] == 0)

    tabela_captacao = tabela_captacao[~mask_sem_net]

    #montagem por células

    tabela_celulas = tabela_captacao.loc[: , tabela_captacao.columns.isin(['Nome assessor','Ticket Médio']) == False].groupby('Time').sum()

    tabela_celulas.loc[: , 'Ticket Médio Indic.'] = tabela_celulas['Net Indicador'] / tabela_celulas['Qtd Clientes Indicados']

    tabela_celulas.loc[: , 'Ticket Médio XP'] = tabela_celulas['NET XP'] / tabela_celulas['Qtd Clientes XP']

    qtd_assessores = assessores.loc[: , ['Nome assessor','Time']].groupby('Time').count()

    tabela_celulas = pd.merge(tabela_celulas , qtd_assessores , left_index=True , right_index=True , how='left')

    tabela_celulas.loc['Total Fatorial','Nome assessor'] = tabela_celulas['Nome assessor'].sum()

    tabela_celulas['Captação Indic. / Assessor'] = tabela_celulas['Captação Indicador'] / tabela_celulas['Nome assessor']

    tabela_celulas['Captação XP / Assessor'] = tabela_celulas['Captação Indicador'] / tabela_celulas['Nome assessor']

    tabela_celulas = tabela_celulas.loc[: , ['Total conta velha Indic.' , 'Total conta nova Indic.' , 'Total Transferências Indic.' , 'Total contas perdidas Indic.' , 'Captação Indicador', 'Captação Indic. / Assessor', 'Total conta velha XP' , 'Total conta nova XP' , 'Total Transferências XP' , 'Total contas perdidas XP' , 'Captação Indicador' , 'Captação XP / Assessor' , 'Qtd Clientes Indicados', 'Ticket Médio Indic.', 'Net Indicador' , 'Qtd Clientes XP', 'Ticket Médio XP','NET XP']]

    tabela_celulas.reset_index(inplace=True)

    #criando coluna de assessores Indicador no positivador novo

    posi_novo = pd.merge(posi_novo , clientes_rodrigo , left_on = 'Cliente' , right_on = 'Conta' , how = 'left')

    posi_novo['Assessor Indicador'].fillna(posi_novo['Assessor'] , inplace = True)

    del posi_novo['Conta']

    posi_novo.loc[posi_novo['Assessor Indicador'] == '1618' , 'Assessor Indicador'] = 'Atendimento Fatorial'

    posi_novo = posi_novo.astype({'Assessor Indicador': str})

    # clientes novos no dia D ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    #pegando a lista acumuladados clientes novos que entraram

    clientes_novos_acum = tabela_novos_transf.loc[: , ['Cliente' , 'Assessor' ,'Aplicação Financeira Declarada' , 'Data de Nascimento' ,'Transferência?']]

    #cruzando com os clientes rodrigo

    clientes_novos_acum = pd.merge(clientes_novos_acum , clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how='left')

    clientes_novos_acum['Assessor Indicador'].fillna(clientes_novos_acum['Assessor'] , inplace = True)

    del clientes_novos_acum['Conta']

    clientes_novos_acum.loc[clientes_novos_acum['Assessor Indicador'] == '1618' , 'Assessor Indicador'] = 'Atendimento Fatorial'

    clientes_novos_acum = clientes_novos_acum.astype({'Assessor Indicador': str})

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

    clientes_novos_acum = pd.merge(clientes_novos_acum , assessores[['Código assessor','Nome assessor']] , how='left' , left_on='Assessor Indicador' , right_on='Código assessor')

    #colocando a profissão do cliente

    clientes_novos_acum = pd.merge(clientes_novos_acum , posi_novo[['Cliente','Profissão']] , how='left' , on='Cliente')

    # filtrando quais são os clientes de hoje

    clientes_novos_hj = clientes_novos_acum.loc[clientes_novos_acum['Presente']=='Hoje' , ['Cliente' , 'Nome' , 'Profissão' , 'Data de Nascimento' ,'Assessor Indicador' , 'Nome assessor' , 'Aplicação Financeira Declarada']]

    clientes_novos_hj.rename(columns = {'Assessor Indicador':'Assessor' , 'Aplicação Financeira Declarada':'PL Declarado'} , inplace=True)

    clientes_novos_hj.reset_index(inplace=True)

    if gera_excel == False:
        return tabela_captacao, tabela_transf

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
                

    writer.save()

    if tabela_captacao['NET XP'].sum() != tabela_captacao['Net Indicador'].sum():
        print("Os valores de AUM não batem, verificar assessores")

    print ("\nRelatório Diário")

relatorio_diario(posi_novo, posi_velho, clientes_rodrigo, assessores, lista_transf, clientes_novos_ontem, suitability)

exit()

meses= [['Janeiro', '310122', '270122'], ['Fevereiro', '250222', '220222'], ['Março', '310322', '290322'], ['Abril', '290422', '270422'], ['Maio', '310522', '300522'], ['Junho', '300622', '280622']]

caminho_assessores = r'bases_dados\Assessores leal_Pablo.xlsx' #!
assessores = pd.read_excel(caminho_assessores)
assessores = assessores.astype({'Código assessor' : str})

for i, [mes, data_hoje, data_ontem] in enumerate(meses):

    caminho_rodrigo = r'Clientes Rodrigo\Clientes Rodrigo ' + mes + '.xlsx'
    clientes_rodrigo = pd.read_excel(caminho_rodrigo)
    clientes_rodrigo = clientes_rodrigo.loc[: , ['Conta' , 'Assessor Relacionamento', 'Assessor Indicador']]
    clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'].isin(['Atendimento Fatorial']) == False , 'Assessor Relacionamento'] = pd.Series(clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'].isin(['Atendimento Fatorial']) == False , 'Assessor Relacionamento']).str.lstrip('A')

    caminho_novo = r'captacao_diario\arquivos\2022\\' + mes + '\positivador_' + data_hoje + '.xlsx' 
    posi_novo = pd.read_excel(caminho_novo , skiprows= 2)
    posi_novo['Assessor'] = posi_novo['Assessor'].astype(str)

    if mes == 'Janeiro':

        caminho_velho = r'captacao_diario\arquivos\2021\Dezembro\positivador_301221.xlsx' #! último positivador do mês anterior
        posi_velho = pd.read_excel(caminho_velho , skiprows=2)
        posi_velho['Assessor'] = posi_velho['Assessor'].astype(str)

        caminho_ontem = r'captacao_diario\arquivos\2021\Dezembro\captacao_291221.xlsx' #!#!#!#!#!#
        clientes_novos_ontem = pd.read_excel(caminho_ontem , sheet_name='Novos + Transf')

    else:
        mes_anterior, data_ultimo_posi, *_ = meses[i-1]
        caminho_velho = r'captacao_diario\arquivos\2022\\' + mes_anterior + '\positivador_' + data_ultimo_posi + '.xlsx' #! último positivador do mês anterior
        posi_velho = pd.read_excel(caminho_velho , skiprows=2)
        posi_velho['Assessor'] = posi_velho['Assessor'].astype(str)

        caminho_ontem = r'captacao_diario\arquivos\2022\\' + mes + '\captacao_' + data_ontem + '.xlsx' #!#!#!#!#!#
        clientes_novos_ontem = pd.read_excel(caminho_ontem , sheet_name='Novos + Transf')

    caminho_transf = f'captacao_diario/arquivos/2022/{mes}/transferencias_{data_hoje}.xlsx' #! pegar no connect a do mesmo dia
    lista_transf = pd.read_excel(caminho_transf)
    lista_transf = lista_transf.loc[lista_transf['Status'] == 'CONCLUÍDO' , :]

    caminho_suitability = r'bases_dados\Suitability.xlsx' #!#!#!#!#!#!#!#
    suitability = pd.read_excel(caminho_suitability)

    resumo, *_ = relatorio_diario(posi_novo, posi_velho, clientes_rodrigo, assessores, lista_transf, clientes_novos_ontem, suitability=suitability, gera_excel=False)

    res_bia = resumo.loc[  resumo['Código assessor'] == '24220', :]

    if mes == 'Janeiro':
        df = res_bia[['Captação Indicador', 'Net Indicador']]
        df.index = ['Janeiro']
    else:
        df.loc[mes] = res_bia[['Captação Indicador', 'Net Indicador']].to_numpy()[0]

gera_excel(df, 'captacao_bia_2022.xlsx', index=True)

