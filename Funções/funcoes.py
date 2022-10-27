'''
Esse arquivo armazena funções importantes e úteis
    1. get_fundos_investidos
    2. add_assessor_relacionamento
    3. add_assessor_indicador
    4. reorder_columns
    5. add_nome_cliente
    6. add_nome_assessor
    7. gera_excel
    8. envia_email_notificacao
    9. atualiza_relacionamento_indicador
    10. get_receita_gerada
    11. export_files_captacao_mes_passado
    12. export_files_captacao
    13. filtro_positivador
    14. captacao
    15. relatorio_diario
    16. gera_pipeline
    17. ranking_diario
    18. rotina_coe
    19. monitora_vencimentos_RF
    20. avisos_novos_transf
    21. distribution
    22. confere_bases_b2b
    23. envia_avisos_clientes_b2b
'''

import io
import win32com.client
import pandas as pd
from classes import Data

def send_mail(mail_from=None, mail_to=None, cc=None, subject=None, body=None, attachment=None, mensagem=''):
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace('MAPI')
    email = outlook.CreateItem(0)
    
    email._oleobj_.Invoke(*(64209,0,8,0,mapi.Accounts.Item(mail_from)))

    if type(mail_to) == list:
        email.To = ";".join(mail_to)
    elif type(mail_to) == str:
        email.To = mail_to
    else:
        print('The mail_to object must be a string or a list.')
        exit()
    
    if not cc == None:
        if type(cc) == str:
            email.CC = cc
        if type(cc) == list:
            email.CC = ";".join(cc)

    email.Subject = subject
    email.HTMLBody = body 
    
    if not attachment == None:
        if type(attachment) == str:
            email.Attachments.Add(attachment)
        if type(attachment) == list:
            for i in range(len(attachment)):
                email.Attachments.Add(attachment[i])
    
    email.Send()
    print(mensagem)

def get_fundos_investidos(diversificador, original_index = False):
    """Retorna os fundos investidos em um certo diversificador, com o index original, ou sem"""
    
    diversificador = diversificador[['Assessor', 'Cliente', 'Produto', 'Sub Produto', 'CNPJ Fundo', 'Ativo', 'NET']]
    fundo = diversificador['Produto'] == 'Fundos'
    fundos_investidos = diversificador[fundo]
    if not original_index:
        fundos_investidos.reset_index(inplace=True, drop=True)
    
    return fundos_investidos

def add_assessor_relacionamento(dataframe, clientes_rodrigo, column_conta='Cliente', positivador = None, column_assessor='Assessor', assessores_com_A = False):
    """Essa função  adiciona a coluna de assessor relacionamento a um dataframe com coluna de contas"""

    if assessores_com_A == False:
        clientes_rodrigo['Assessor Relacionamento'] = clientes_rodrigo['Assessor Relacionamento'].str.strip('A')
        clientes_rodrigo['Assessor Relacionamento'] = clientes_rodrigo['Assessor Relacionamento'].str.replace('tendimento Fatorial', 'Atendimento Fatorial')
    clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Relacionamento']]
    dataframe = dataframe.merge(clientes_rodrigo, how='left', left_on=column_conta, right_on='Conta')
    dataframe.drop(['Conta'], axis=1, inplace=True)

    null = dataframe['Assessor Relacionamento'].isnull().to_numpy()
    clientes = dataframe[column_conta].to_numpy()
    
    for i, *_ in enumerate(dataframe['Assessor Relacionamento'].to_numpy()):        
        is_null = null[i]
        if is_null:
            
            if column_assessor != None:
                assessor_relacionamento = dataframe.loc[i, column_assessor]
                dataframe.loc[i,'Assessor Relacionamento'] = assessor_relacionamento
    
            elif column_assessor == None:
                cliente_selecionado = clientes[i]
                mask_cliente = positivador['Cliente'] == cliente_selecionado
                assessor_relacionamento = positivador.loc[mask_cliente , 'Assessor'].iloc[0]
                dataframe.loc[i,'Assessor Relacionamento'] = assessor_relacionamento
    
    return dataframe

def add_assessor_indicador(dataframe, clientes_rodrigo, column_conta='Cliente', column_assessor='Assessor', assessores_com_A = False, positivador=None):
    """Essa função adiciona a coluna de assessor indicador a um dataframe com coluna de contas"""

    if assessores_com_A == False:
        clientes_rodrigo['Assessor Indicador'] = clientes_rodrigo['Assessor Indicador'].str.strip('A')
        clientes_rodrigo['Assessor Indicador'] = clientes_rodrigo['Assessor Indicador'].str.replace('tendimento Fatorial', 'Atendimento Fatorial')
    clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Indicador']]
    dataframe = dataframe.merge(clientes_rodrigo, how='left', left_on=column_conta, right_on='Conta')
    dataframe.drop(['Conta'], axis=1, inplace=True)

    null = dataframe['Assessor Indicador'].isnull().to_numpy()
    clientes = dataframe[column_conta].to_numpy()
    
    for i, *_ in enumerate(dataframe['Assessor Indicador'].to_numpy()):        
        is_null = null[i]
        if is_null:
            
            if column_assessor != None:
                assessor_indicador = dataframe.loc[i, column_assessor]
                dataframe.loc[i,'Assessor Indicador'] = assessor_indicador
    
            elif column_assessor == None:
                cliente_selecionado = clientes[i]
                mask_cliente = positivador['Cliente'] == cliente_selecionado
                assessor_indicador = positivador.loc[mask_cliente , 'Assessor'].iloc[0]
                dataframe.loc[i,'Assessor Indicador'] = assessor_indicador
    
    return dataframe

def reorder_columns(dataframe, col_name, position):
    """Reorder a dataframe's column.
    Args:
        dataframe (pd.DataFrame): dataframe to use
        col_name (string): column name to move
        position (0-indexed position): where to relocate column to
    Returns:
        pd.DataFrame: re-assigned dataframe
    """
    temp_col = dataframe[col_name]
    dataframe = dataframe.drop(columns=[col_name])
    dataframe.insert(loc=position, column=col_name, value=temp_col)
    return dataframe

def add_nome_cliente(dataframe, column_conta, suitability):
    """Essa função adiciona o nome de um cliente, com base na Suitability"""
    suitability = suitability [['CodigoBolsa', 'NomeCliente']]
    dataframe = dataframe.merge(suitability, how='left', left_on=column_conta, right_on='CodigoBolsa')
    dataframe = dataframe.drop('CodigoBolsa', axis = 1)
    
    return dataframe

def add_nome_assessor(dataframe, column_assessor, assessores):
    """Essa função adiciona o nome do assessor, com base na assessores leal pablo"""
    assessores = assessores [['Código assessor', 'Nome assessor']]
    dataframe = dataframe.merge(assessores, how='left', left_on=column_assessor, right_on='Código assessor')
    if column_assessor != 'Código assessor':
        dataframe = dataframe.drop('Código assessor', axis = 1)
    
    return dataframe

def gera_excel(tabela, caminho_excel, sheet_name='Sheet1', index=False, skiprows=0, mensagem='arquivo criado'):
    """Gera uma planilha em excel com a tabela indicada"""
    import pandas as pd
    writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')

    tabela.to_excel(writer , sheet_name= sheet_name , index= index, startrow=skiprows)

    writer.save()

    print(mensagem)

def envia_email_notificacao(assessor, email_destino, dataframe):
    """Essa função envia e-mails automáticos para assessores, 
    a fim de notificá-los de clientes problemáticos, ou seja, 
    que entraram há menos de 3 meses e não superaram 300 mil""" 

    html_df = dataframe.to_html(index=False, justify = 'center')

    # criar a integração com o outlook
    outlook = win32com.client.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # corpo do email
    body = f"""
    <p>Olá, {assessor}<p>
    
    <p>Aqui é o Leonardo Motta, da Inteligência da Fatorial.<p>

    <p>Fizemos um levantamento dos clientes que entraram desde novembro, e estão com patrimônio 
    menor do que R$ 300.000,00. Nessa filtragem, encontramos estes clientes seus nessa situação:<p>

    <p>{html_df}<p>

    <p><p>Att,<p>
    <p> Leonardo Gonçalves Motta <p>
    """

    # configurar as informações do seu e-mail
    email.To = email_destino
    email.Subject = "Aviso: Clientes abaixo de 300 mil"
    email.HTMLBody = body


    email.Send()
    print(f"\nEmail Enviado para {assessor}\n")

def atualiza_relacionamento_indicador(dataframe, clientes_rodrigo, column_indicador= 'Assessor Indicador', column_relacionamento = 'Assessor Relacionamento', only_indicador = False, only_relacionamento = False):
    """Recebe um dataframe com colunas de assessor relacionamento e 
    indicador e atualiza de acordo com a clientes rodrigo mais recente"""

    if only_relacionamento == True:
        dataframe.drop([column_relacionamento], inplace=True)

        dataframe = add_assessor_relacionamento(dataframe, clientes_rodrigo, column_conta = 'Cliente')

        dataframe = reorder_columns(dataframe, column_relacionamento, 1)

        return dataframe

    if only_indicador == True:
        dataframe.drop([column_indicador], inplace=True)

        dataframe = add_assessor_indicador(dataframe, clientes_rodrigo, column_conta = 'Cliente')

        dataframe = reorder_columns(dataframe, column_indicador, 0)

        return dataframe

    dataframe.drop([column_indicador, column_relacionamento], axis=1, inplace=True)

    dataframe = add_assessor_relacionamento(dataframe, clientes_rodrigo, column_conta = 'Cliente')

    dataframe = add_assessor_indicador(dataframe, clientes_rodrigo)

    dataframe = reorder_columns(dataframe, column_indicador, 0)

    dataframe = reorder_columns(dataframe, column_relacionamento, 2)
    
    return dataframe

def get_receita_gerada(df_comissao, clientes_rodrigo, split_method='50'):
    """Essa função recebe de entrada uma base de dados com todas as movimentações de receita
    ocorridas no mês, para calcular o ganho de cada assessor, baseado em três tipos: indicador
    50/50 e relacionamento"""

    import numpy as np

    df_comissao = df_comissao.astype({'Cliente':str})
    df_comissao['Cliente'] == df_comissao['Cliente'].replace('nan', np.NaN, inplace=True)

    clientes_rodrigo = clientes_rodrigo.astype({'Conta':str})

    clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Relacionamento', 'Assessor Indicador']]

    df_comissao = df_comissao[['ID', 'Cliente', 'Assessor Dono', 'Repasse Corretora']]

    df_comissao['Assessor Dono'] = df_comissao['Assessor Dono'].str.lower()

    name_assessores = df_comissao['Assessor Dono'].to_numpy()
    list_names_corrigido = []
    for assessor in name_assessores:
        list_words = assessor.split()
        list_words = [name.capitalize() for name in list_words]
        list_names_corrigido.append(" ".join(list_words))
    
    df_comissao['Assessor Dono'] = list_names_corrigido

    qtd_envolvido=[]
    for id in df_comissao['ID'].to_numpy():
        mask_id = df_comissao['ID'] == id
        recorrencia = len(df_comissao['ID'][mask_id])
        qtd_envolvido.append(recorrencia)

    df_comissao['Número Envolvidos'] = qtd_envolvido

    df_comissao.drop_duplicates(subset=['ID'], inplace=True)

    df_comissao.dropna(subset=['Cliente'], inplace=True)

    df_comissao = add_assessor_indicador(df_comissao, clientes_rodrigo, column_assessor='Assessor Dono', assessores_com_A=True)

    df_comissao = reorder_columns(df_comissao, 'Assessor Indicador', 3)

    df_comissao.rename(columns={'Assessor Dono': 'Assessor Relacionamento'}, inplace=True)

    if split_method == 'Indicador':
        df_comissao['Receita Relacionamento'] = df_comissao['Repasse Corretora'] * 0
        df_comissao['Receita Indicador'] = df_comissao['Repasse Corretora'] * 1

    elif split_method == 'Relacionamento':
        df_comissao['Receita Relacionamento'] = df_comissao['Repasse Corretora'] * 1
        df_comissao['Receita Indicador'] = df_comissao['Repasse Corretora'] * 0

    elif split_method == '50':
        df_comissao['Receita Relacionamento'] = df_comissao['Repasse Corretora'] / 2
        df_comissao['Receita Indicador'] = df_comissao['Repasse Corretora'] / 2

    return df_comissao

def export_files_captacao_mes_passado(data_hoje, data_ontem, mes_passado, datafechamento):

    meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    idx_mes_passado = meses.index(str.capitalize(mes_passado))
    mes_retrasado = meses[idx_mes_passado - 1]

    #importando positivador mês M 

    caminho_novo = r'captacao_diario\arquivos\2022\\' + str.capitalize(mes_passado) + '\positivador_' + data_hoje + '.xlsx' #! positivador do dia de run do código
    posi_novo = pd.read_excel(caminho_novo , skiprows= 2)

    # importando positivador do dia D-1
    
    caminho_d1 = r'captacao_diario\arquivos\2022\\' + str.capitalize(mes_passado) + '\positivador_' + data_ontem + '.xlsx' #! positivador do dia anterior ao de run do codigo
    posi_d1 = pd.read_excel(caminho_d1 , skiprows= 2) 

    # importando positivador mês M - 1 
    caminho_velho = r'captacao_diario\arquivos\2022\\' + str.capitalize(mes_retrasado) + '\positivador_' + datafechamento + '.xlsx' #! último positivador do mês anterior
    posi_velho = pd.read_excel(caminho_velho , skiprows=2)

    #importando clientes rodrigo e fazendo tratamentos (selecionando as colunas , substituindo "Atendimento" pelo A do yago e tirando o prefixo A)
    
    caminho_rodrigo = r'Clientes Rodrigo\Clientes Rodrigo ' + mes_passado + '.xlsx' #! verificar se a clientes do rodrigo tá atualizada
    clientes_rodrigo = pd.read_excel(caminho_rodrigo , sheet_name='Troca')
    clientes_rodrigo = clientes_rodrigo.loc[: , ['Conta' , 'Assessor Relacionamento', 'Assessor Indicador']]
    clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'].isin(['Atendimento Fatorial']) == False , 'Assessor Relacionamento'] = pd.Series(clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'].isin(['Atendimento Fatorial']) == False , 'Assessor Relacionamento']).str.lstrip('A')

    #importando lista de assessores 
    
    caminho_assessores = r'bases_dados\Assessores leal_Pablo.xlsx' #!
    assessores = pd.read_excel(caminho_assessores)
    assessores = assessores.astype({'Código assessor' : str})

    #importando lista de transferências
    
    caminho_transf = r'captacao_diario\arquivos\2022\\' + str.capitalize(mes_passado) + '\\transferencias_' + data_hoje + '.xlsx' #! pegar no connect a do mesmo dia
    lista_transf = pd.read_excel(caminho_transf)
    lista_transf = lista_transf.loc[lista_transf['Status'] == 'CONCLUÍDO' , :]

    #importando o relatório de captação D-1

    caminho_ontem = r'captacao_diario\arquivos\2022\\' + mes_passado + '\captacao_' + data_ontem + '.xlsx' #!#!#!#!#!#
    clientes_novos_ontem = pd.read_excel(caminho_ontem , sheet_name='Novos + Transf')

    #importando o suitability

    caminho_suitability = r'bases_dados\Suitability.xlsx' #!#!#!#!#!#!#!#
    suitability = pd.read_excel(caminho_suitability)

    #diretorio arquivo de registro

    caminho_registro = r'bases_dados\Registro de Transferências\registro_transferência_12_21.xlsx' #MUDAR PARA O MÊS DE RUN DO CODIGO
    registro_transf = pd.read_excel(caminho_registro)
    registro_transf.columns = ['Assessor', 'Cliente', 'Data de Chegada', 'Net de Chegada']

    return posi_novo, posi_d1, posi_velho, clientes_rodrigo, assessores, lista_transf, clientes_novos_ontem, suitability, registro_transf

def export_files_captacao(data_hoje, data_ontem, responsavel_digital = "Atendimento Fatorial"):

    import pandas as pd

    #importando positivador mês M 

    caminho_novo = r'captacao_diario\positivador_' + data_hoje + '.xlsx' #! positivador do dia de run do código
    posi_novo = pd.read_excel(caminho_novo , skiprows= 2)

    # importando positivador do dia D-1
    
    caminho_d1 = r'captacao_diario\positivador_' + data_ontem + '.xlsx' #! positivador do dia anterior ao de run do codigo
    posi_d1 = pd.read_excel(caminho_d1 , skiprows= 2) 

    # importando positivador mês M - 1 
    caminho_velho = r'captacao_diario\arquivos\2022\Setembro\positivador_300922.xlsx' #! último positivador do mês anterior
    posi_velho = pd.read_excel(caminho_velho , skiprows=2)

    #importando clientes rodrigo e fazendo tratamentos (selecionando as colunas , substituindo "Atendimento" pelo A do yago e tirando o prefixo A)
    
    caminho_rodrigo = r'bases_dados\Clientes do Rodrigo.xlsx' #! verificar se a clientes do rodrigo tá atualizada
    clientes_rodrigo = pd.read_excel(caminho_rodrigo , sheet_name='Troca')
    clientes_rodrigo = clientes_rodrigo.loc[: , ['Conta' , 'Assessor Relacionamento', 'Assessor Indicador']]
    clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'].isin(['Atendimento Fatorial']) == False , 'Assessor Relacionamento'] = pd.Series(clientes_rodrigo.loc[clientes_rodrigo['Assessor Relacionamento'].isin(['Atendimento Fatorial']) == False , 'Assessor Relacionamento']).str.lstrip('A')

    #importando lista de assessores 
    
    caminho_assessores = r'bases_dados\Assessores leal_Pablo.xlsx' #!
    assessores = pd.read_excel(caminho_assessores)
    assessores = assessores.astype({'Código assessor' : str})

    #importando lista de transferências
    
    caminho_transf = r'captacao_diario\transferencias_' + data_hoje + '.xlsx' #! pegar no connect a do mesmo dia
    lista_transf = pd.read_excel(caminho_transf)
    lista_transf = lista_transf.loc[lista_transf['Status'] == 'CONCLUÍDO' , :]

    #importando o relatório de captação D-1

    caminho_ontem = r'captacao_diario\captacao_' + data_ontem + '.xlsx' #!#!#!#!#!#
    clientes_novos_ontem = pd.read_excel(caminho_ontem , sheet_name='Novos + Transf')

    #importando o suitability

    caminho_suitability = r'bases_dados\Suitability.xlsx' #!#!#!#!#!#!#!#
    suitability = pd.read_excel(caminho_suitability)

    #diretorio arquivo de registro

    caminho_registro = r'bases_dados\Registro de Transferências\registro_transferência_12_21.xlsx' #MUDAR PARA O MÊS DE RUN DO CODIGO
    registro_transf = pd.read_excel(caminho_registro)
    registro_transf.columns = ['Assessor', 'Cliente', 'Data de Chegada', 'Net de Chegada']

    return posi_novo, posi_d1, posi_velho, clientes_rodrigo, assessores, lista_transf, clientes_novos_ontem, suitability, registro_transf

def filtro_positivador(posi_novo, posi_d1, posi_velho, clientes_rodrigo, lista_transf, registro_transf, data_hoje, responsavel_digital = "Atendimento Fatorial", year=2022):
    """Esse código tem por objetivo fazer uma filtragem nas transferências do positivador da XP, 
    registrando o Net Em M-1 deles como sendo, idêntico, ao Net de entrada. Esse código é rodado
    previamente ao de captação."""

    import datetime
    import pandas as pd
    import datetime
    import numpy as np

    #diretorio arquivo final

    caminho_excel = r'captacao_diario\positivador_'  + data_hoje + '.xlsx' #!
    caminho_registro = r'bases_dados\Registro de Transferências\registro_transferência_12_21.xlsx'

    clientes_rodrigo = clientes_rodrigo.loc[:, ['Conta', 'Assessor Relacionamento']]

    #montando tabela clientes perdidos ----------------------------------------------------------------------------------------------------------------------------

    #encontrando quais são os clientes perdidos
    posi_velho['Clientes perdidos'] = posi_velho['Cliente'].where(posi_velho['Cliente'].isin(posi_novo['Cliente']) == True)

    #renomeando os valores: na -> "Saiu" ; código cliente -> "Permanece"
    posi_velho['Clientes perdidos'].fillna('Saiu' , inplace = True)

    posi_velho.loc[posi_velho['Clientes perdidos'] != 'Saiu' , 'Clientes perdidos'] = 'Permanece'

    #montagem do dataframe
    tabela_perdidos = posi_velho.loc[posi_velho['Clientes perdidos'] == 'Saiu' , :]

    #criação de coluna de assessor correto
    tabela_perdidos = tabela_perdidos.merge(clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how= 'left')

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
    tabela_velhos = tabela_velhos.merge(clientes_rodrigo , left_on= 'Cliente' , right_on= 'Conta' , how= 'left')

    tabela_velhos['Assessor Relacionamento'].fillna(tabela_velhos['Assessor'] , inplace = True)

    tabela_velhos.rename(columns={'Assessor Relacionamento':'Assessor correto'} , inplace= True)

    del tabela_velhos['Conta']

    tabela_velhos.loc[tabela_velhos['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

    # montando tabela de clientes novos e transferencias ----------------------------------------------------------------------------------------------------------------------------

    #tabela clientes novos + transferencias
    tabela_novos_transf = posi_novo.loc[posi_novo['Status conta'] == 'conta nova' , :]


    #identificando quais são as transferências

    mask_transferencia = tabela_novos_transf.loc[: , 'Cliente'].isin(lista_transf.loc[: , 'Cliente'])

    print(mask_transferencia)

    tabela_novos_transf.loc[mask_transferencia, 'Transferência?'] = "Sim"

    tabela_novos_transf['Transferência?'].fillna('Não' , inplace = True)

    tabela_novos_transf.loc[tabela_novos_transf['Net em M-1'] > 0 , 'Transferência?'] = 'Sim'

    #tabela clientes novos
    tabela_novos = tabela_novos_transf.loc[tabela_novos_transf['Transferência?'] == 'Não' , :]

    #cliente rodrigo para os novos
    tabela_novos = tabela_novos.merge(clientes_rodrigo , left_on= 'Cliente' , right_on= 'Conta' , how = 'left')

    tabela_novos['Assessor Relacionamento'].fillna(tabela_novos['Assessor'] , inplace = True)

    tabela_novos.rename(columns = {'Assessor Relacionamento' : 'Assessor correto'} , inplace= True)

    del tabela_novos['Conta']

    tabela_novos.loc[tabela_novos['Assessor correto'] == '1618' , 'Assessor correto'] = responsavel_digital

    #tabela transferências
    tabela_transf = tabela_novos_transf.loc[tabela_novos_transf['Transferência?'] == 'Sim' , :]

    #cliente rodrigo para as transferências
    tabela_transf = tabela_transf.merge(clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how='left')

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

        writer.save ()

    # coloca os aniversários que vem faltando

    sem_aniversario = posi_novo[posi_novo['Data de Nascimento'].isnull()]

    clientes_sem_aniversario = sem_aniversario['Cliente'].to_numpy()

    clientes_ontem = posi_d1['Cliente'].to_numpy()

    for cliente in clientes_sem_aniversario:
        if cliente in clientes_ontem:
            aniversario_cliente = posi_d1.loc[ posi_d1['Cliente'] == cliente, 'Data de Nascimento' ].to_numpy()[0]
            posi_novo.loc[ posi_novo['Cliente'] == cliente, 'Data de Nascimento' ] = aniversario_cliente

    if not any(posi_novo['Data de Nascimento'].isnull()): # se todos os clientes têm aniversário
        print("\nTodos os clientes tem aniversário registrado")

    else:
        print("\nClientes sem aniversário:\n\n", posi_novo[posi_novo['Data de Nascimento'].isnull()], '\n')


    # Escreve o positivador no excel

    writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
    posi_novo.to_excel(writer , sheet_name='Sheet1',index=False, startrow=2)

    writer.save ()

    # salva um positivador dinamico

    gera_excel(posi_novo, r'captacao_diario\positivador_dinamico.xlsx')

    print('\nPositivador filtrado')

def captacao(posi_novo, posi_velho, clientes_rodrigo, assessores, lista_transf, clientes_novos_ontem, suitability, data_hoje, responsavel_digital = "Atendimento Fatorial", year=2022, gera_excel = True):

    """Tanto separando os assessores em times, quanto analisando individualmente, esse código gera um
    relatório diário, com base nos arquivos da XP, reportando a ccaptação e o AUM dos assessores."""

    import pandas as pd

    #diretorio arquivo final

    caminho_excel = r'captacao_diario\captacao_' + data_hoje + '.xlsx' #!

    clientes_rodrigo = clientes_rodrigo.loc[:, ['Conta', 'Assessor Relacionamento']]

    #montando tabela clientes perdidos ----------------------------------------------------------------------------------------------------------------------------

    #encontrando quais são os clientes perdidos
    posi_velho['Clientes perdidos'] = posi_velho['Cliente'].where(posi_velho['Cliente'].isin(posi_novo['Cliente']) == True)

    #renomeando os valores: na -> "Saiu" ; código cliente -> "Permanece"
    posi_velho['Clientes perdidos'].fillna('Saiu' , inplace = True)

    posi_velho.loc[posi_velho['Clientes perdidos'] != 'Saiu' , 'Clientes perdidos'] = 'Permanece'

    #montagem do dataframe
    tabela_perdidos = posi_velho.loc[posi_velho['Clientes perdidos'] == 'Saiu' , :]

    #criação de coluna de assessor correto
    tabela_perdidos = tabela_perdidos.merge(clientes_rodrigo , left_on='Cliente' , right_on='Conta' , how= 'left')

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
    
    #Tira do resumo quem não tem patrimonio

    mask_sem_net = tabela_captacao['Net Em M'] == 0

    mask_sem_net_xp = tabela_captacao['NET XP'] == 0

    mask_sem_capt = tabela_captacao['Captação Líquida'] == 0

    tabela_captacao = tabela_captacao[~(mask_sem_net & mask_sem_capt & mask_sem_net_xp)]


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

    if tabela_captacao['NET XP'].sum() != tabela_captacao['Net Em M'].sum():
        print("Os valores de AUM não batem, verificar assessores")

    print ("\nRelatório de Captação gerado")

def gera_pipeline(data_pipe, assessores):
    """Esse código tem por objetivo registrar os dados de pipeline que são exportados do
    CRM. Essas informações são usadas na geração do ranking."""
    
    import pandas as pd

    oportunidades= pd.read_excel(r"pipeline\pipeline_CRM\Pipeline_" + data_pipe + ".xlsx", sheet_name='Oportunidades')
    leads = pd.read_excel(r"pipeline\pipeline_CRM\Pipeline_" + data_pipe + ".xlsx", sheet_name='Leads')

    caminho_excel = r"pipeline\Relatórios\Relatório de Oportunidades e Leads_" + data_pipe + ".xlsx"

    # tirar o A dos códigos de assessores

    oportunidades['Assessor'] = oportunidades['Assessor'].str.lstrip('A')
    oportunidades['Criador'] = oportunidades['Criador'].str.lstrip('A')

    leads['Assessor'] = leads['Assessor'].str.lstrip('A')
    leads['Criador'] = leads['Criador'].str.lstrip('A')

    # adiciona nome do assessor e nome do criador

    # oportunidades

    oportunidades = add_nome_assessor(oportunidades, column_assessor='Criador',assessores=assessores)
    oportunidades.rename(columns={'Nome assessor':'Quem criou?'}, inplace=True)
    oportunidades['Quem criou?'].fillna('Leonardo Motta', inplace=True)

    oportunidades = add_nome_assessor(oportunidades, column_assessor='Assessor',assessores=assessores)
    oportunidades = reorder_columns(oportunidades, col_name='Nome assessor', position = 4)

    # leads
    leads = add_nome_assessor(leads, column_assessor='Criador',assessores=assessores)
    leads.rename(columns={'Nome assessor':'Quem criou?'}, inplace=True)
    leads['Quem criou?'].fillna('Leonardo Motta', inplace=True)

    leads = add_nome_assessor(leads, column_assessor='Assessor',assessores=assessores)
    leads = reorder_columns(leads, col_name='Nome assessor', position = 5)

    # adiciona zero aos aportes vazios

    oportunidades['Valor'].fillna(0, inplace=True)
    leads['1º aporte estimado'].fillna(0, inplace=True)

    # estrutura do relatório geral, com base das oportunidades

    relatorio_geral = oportunidades
    del relatorio_geral['Data de Fechamento']
    relatorio_geral = reorder_columns(relatorio_geral, 'Fase', 6)

    # assimila o padrão aos leads

    leads_formatado = leads

    leads_formatado['Cliente'] = [None for i in range(len(leads_formatado.index))]
    leads_formatado['Nome do Cliente'] = [None for i in range(len(leads_formatado.index))]

    del leads_formatado['Data de Criação']
    del leads_formatado['Não Abriu']

    leads_formatado = leads_formatado[['Nome da oportunidade' , 'Cliente', 'Nome do Cliente' ,'Assessor', 'Nome assessor', '1º aporte estimado', 'Fase', 'Criador', 'Quem criou?']]

    # append das bases

    columns = relatorio_geral.columns
    leads_formatado.columns = columns

    relatorio_geral = pd.concat([relatorio_geral, leads_formatado])

    relatorio_assessores = relatorio_geral.groupby('Nome assessor').sum()

    # gera excel

    writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')

    relatorio_geral.to_excel(writer , sheet_name= 'Detalhado')

    relatorio_assessores.to_excel(writer, sheet_name='Resumo')

    writer.save()

    print('\nPipeline dos assessores Registrado')

    return relatorio_assessores

def ranking_diario(pipeline, data_hoje):

    import pandas as pd

    # Base de Dados

    captacao_acumulado = pd.read_excel (r"captacao_diario\arquivos\2022\captacao_2022.xlsx")

    captacao_acumulado = captacao_acumulado[['Código assessor', 'Acumulado 2022']]

    captacao_acumulado = captacao_acumulado.astype({'Código assessor':str})

    captacao_hj = pd.read_excel (r"captacao_diario\captacao_" + data_hoje + ".xlsx",sheet_name="Resumo")

    captacao_hj = captacao_hj.merge(captacao_acumulado, how='left', on='Código assessor')

    captacao_hj = captacao_hj [["Time","Nome assessor","Captação Líquida", "Acumulado 2022"]]

    captacao_hj['Acumulado 2022'] += captacao_hj['Captação Líquida']

    captacao_hj = captacao_hj.loc [captacao_hj["Time"]!="Fora da Fatorial",:]

    captacao_hj = captacao_hj.loc [captacao_hj["Time"]!="Total Fatorial",:]

    captacao_hj = captacao_hj.sort_values  (by="Captação Líquida",ascending=False)

    captacao_hj = captacao_hj [["Nome assessor","Time","Captação Líquida", "Acumulado 2022"]]

    # Diretório do Excel
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Nome assessor'] == "Paulo Valinote"].index)
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Nome assessor'] == "Paulo Barros"].index)
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Nome assessor'] == "Paulo Monfort"].index)
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Nome assessor'] == "Base Fatorial"].index)

    caminho_excel = r"ranking_diario\arquivos\2022\ranking_" + data_hoje + "_geral.xlsx" #!

    writer = pd.ExcelWriter(caminho_excel , 
                            engine='xlsxwriter', 
                            datetime_format = 'dd/mm/yyyy')

    captacao_hj.to_excel(writer , sheet_name= 'Sheet1' , index= False)

    aba = writer.sheets['Sheet1']
        
    colunas = [{'header':column} for column in captacao_hj.columns ]
    (lin, col) = captacao_hj.shape

    aba.add_table(0 , 0 , lin , col-1 , {
        'columns': colunas,
        'style': 'Table Style Medium 2', 
        'autofilter': False
        })

    writer.save()

    # Ranking filtrado

    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Time'] == "Digital"].index)
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Time'] == "Mesa RV"].index)
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Time'] == "Não Comercial"].index)
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Time'] == "Saiu da Fatorial"].index)

    '''captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Nome assessor'] == "Beatriz Paiva"].index)
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Nome assessor'] == "Flávio Camargo"].index)
    captacao_hj = captacao_hj.drop(captacao_hj[captacao_hj['Nome assessor'] == "Renata Schneider"].index)'''

    socios = ('Rodrigo Cabral', 'Jansen Costa', 'Octavio Bastos', 'Pablo Langenbach')
    captacao_hj = captacao_hj[~captacao_hj['Nome assessor'].isin(socios)]

    captacao_hj["Pipeline"] = ''
    captacao_hj = captacao_hj [["Nome assessor","Time","Captação Líquida", "Acumulado 2022"]]

    captacao_hj = pd.merge(captacao_hj, pipeline, on='Nome assessor', how='left')

    captacao_hj = captacao_hj.rename(columns={"Valor":"Pipeline"})

    caminho_excel = r"ranking_diario\arquivos\2022\ranking_" + data_hoje + "_filtrado.xlsx" #!

    writer = pd.ExcelWriter(caminho_excel , 
                            engine='xlsxwriter', 
                            datetime_format = 'dd/mm/yyyy')

    captacao_hj.to_excel(writer , sheet_name= 'Sheet1' , index= False)

    aba = writer.sheets['Sheet1']
        
    colunas = [{'header':column} for column in captacao_hj.columns ]
    (lin, col) = captacao_hj.shape

    aba.add_table(0 , 0 , lin , col-1 , {
        'columns': colunas,
        'style': 'Table Style Medium 2', 
        'autofilter': False
        })

    writer.save()

    print ("\nRanking gerado")

def relatorio_diario(posi_novo, posi_velho, clientes_rodrigo, assessores, lista_transf, clientes_novos_ontem, suitability, data_hoje, gera_excel=True):
    """Tanto separando os assessores em times, quanto analisando individualmente, esse código gera um
    relatório diário, com base nos arquivos da XP, reportando a captação e o AUM dos assessores, pela ótica da XP e do Indicador."""

    import pandas as pd

    #diretorio arquivo final

    caminho_excel = r'captacao_diario\relatorio_diario_' + data_hoje + '.xlsx' #!

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

def rotina_coe(dia_hoje, mes):
    import pandas as pd
    caminho_excel = r"COE\Análise de COE\Relatórios\Relatório_COE_" + mes + ".xlsx"

    COE_dataset = pd.read_excel(r"COE\Análise de COE\COE_" + mes + ".xlsx")

    clientes_rodrigo = pd.read_excel(r'bases_dados\Clientes do Rodrigo.xlsx', sheet_name='Troca')

    assessores = pd.read_excel(r'bases_dados\Assessores leal_Pablo.xlsx')
    assessores = assessores.astype({'Código assessor' : str})

    positivador = pd.read_excel(r"captacao_diario\positivador_" + dia_hoje + ".xlsx", skiprows=2)
    positivador['Assessor'] = positivador['Assessor'].astype(str)

    COE_dataset = add_assessor_relacionamento(COE_dataset, clientes_rodrigo =clientes_rodrigo, column_assessor=None, positivador=positivador)

    COE_dataset = add_nome_assessor(COE_dataset, 'Assessor Relacionamento',assessores)

    COE_dataset = reorder_columns(COE_dataset, 'Assessor Relacionamento', 0)

    COE_dataset = reorder_columns(COE_dataset, 'Nome assessor', 1)

    COE_dataset['Financeiro'] = COE_dataset['Financeiro'].str[3:]
    COE_dataset['Financeiro'] = COE_dataset['Financeiro'].str.replace(',', '')
    COE_dataset['Financeiro'] = COE_dataset['Financeiro'].str.replace('.', '')
    COE_dataset['Financeiro'] = COE_dataset['Financeiro'].astype(float)/100

    COE_dataset['Comissão'].replace('---', 0, inplace=True)

    COE_dataset['Comissão assessor'] = COE_dataset['Financeiro']*COE_dataset['Comissão']

    COE_dataset = reorder_columns(COE_dataset, 'Comissão assessor', 6)

    assessores_geraram = COE_dataset['Assessor Relacionamento'].drop_duplicates()

    gerou_COE = assessores['Código assessor'].isin(assessores_geraram)

    controle_COE = assessores[~gerou_COE]

    comercial = controle_COE['Time'] != 'Não Comercial'

    controle_COE = controle_COE[comercial]

    controle_COE['Controle COE'] = 'Não gerou'

    writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')

    COE_dataset.to_excel(writer , sheet_name= 'Registro COE' , index= False)

    controle_COE.to_excel(writer, sheet_name= 'Controle COE', index = False)

    writer.save()

    print('Registros de COE finalizados')

def monitora_vencimentos_RF(data_hoje, clientes_rodrigo, assessores):

    import pandas as pd
    import numpy
    from datetime import date, timedelta

    diversificador = pd.read_excel(r"COE\arquivos\diversificacao_" + data_hoje + ".xlsx", skiprows=2)
    diversificador['Assessor'] = diversificador['Assessor'].astype(str)

    caminho_excel = r"COE\Monitora Vencimentos\Monitoramento de Vencimentos_" + data_hoje + ".xlsx" # mudar mês

    diversificador = diversificador[['Assessor', 'Cliente', 'Produto', 'Ativo', 'Data de Vencimento', 'NET']]

    # define o assessor relacionamento e bota o nome dele

    diversificador = add_assessor_relacionamento(diversificador, clientes_rodrigo)
    diversificador = add_nome_assessor(diversificador, 'Assessor Relacionamento', assessores)
    diversificador = reorder_columns(diversificador, 'Assessor Relacionamento', 0)
    diversificador = reorder_columns(diversificador, 'Nome assessor', 1)
    diversificador.drop('Assessor', axis=1,inplace=True)

    # definição dos objetos de data como numpy.datetime64

    hoje = date.today()
    dez_dias = hoje + timedelta(days=10)

    hoje = numpy.datetime64(hoje)
    dez_dias = numpy.datetime64(dez_dias)

    # definição das máscaras para ver os que estão para vencer

    mask_vai_vencer = diversificador['Data de Vencimento'] < dez_dias

    mask_RF = diversificador['Produto'] == 'Renda Fixa'

    vai_vencer = diversificador[mask_RF & mask_vai_vencer]

    # isola o que já venceu

    mask_venceu = vai_vencer['Data de Vencimento'] < hoje

    vencidos = vai_vencer[mask_venceu]

    vai_vencer = vai_vencer[~mask_venceu]

    vencidos.sort_values('Data de Vencimento', inplace=True)
    vai_vencer.sort_values('Data de Vencimento', inplace=True)

    # exporta em excel

    writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
    vencidos.to_excel(writer , sheet_name='Vencidos',index=False)
    vai_vencer.to_excel(writer , sheet_name='Vencem em 10 dias',index=False)

    writer.save()

    print("Relatório de Monitoramento de Vencimentos Atualizado")

def avisos_novos_transf(mes, mes1, posi):


    import pandas as pd

    posi['Cliente'] = posi['Cliente'].astype(int)

    caminho_relatorio = r'Aviso Novos Transf\ novos_transf_' + mes + '.xlsx'

    clientes_rodrigo = pd.read_excel(r'bases_dados\Clientes do Rodrigo.xlsx' , sheet_name='Troca')

    registro_ascensoes = pd.read_excel(r"Aviso Novos Transf\Ascensões\Ascensões_acumulado.xlsx")
    registro_ascensoes['Cliente'] = registro_ascensoes['Cliente'].astype(str)

    base_dados = pd.read_excel(r"Aviso Novos Transf\ novos_transf_" + mes1 +".xlsx", sheet_name='Dataset') #mês anterior
    base_dados['Cliente'] = base_dados['Cliente'].astype(int)

    # tira da base de dados os clientes que saíram

    cliente_na_fatorial = base_dados['Cliente'].isin(posi['Cliente'])
    clientes_sairam = base_dados[~cliente_na_fatorial]
    base_dados = base_dados[cliente_na_fatorial]

    print('\n\nClientes que sairam:\n', clientes_sairam)

    # atualiza a os positivadores antigos com o positivador novo, além da base de dados

    net_atualizado = []
    for conta in base_dados['Cliente'].to_numpy():
        net_atualizado.append(posi.loc[posi['Cliente'] == conta, 'Net Em M'].values[0])
    base_dados['Net Em M'] = net_atualizado

    # gera a base de dados dos clientes que superaram 300 mil

    is_maior_300 = base_dados['Net Em M'] > 300000
    superaram_300 = base_dados[is_maior_300]
    base_dados = base_dados[~is_maior_300]

    print('\n\nClientes que superaram 300 mil:\n', superaram_300)

    # registra os clientes que ascenderam

    ascensoes = superaram_300
    ascensoes['Mês de Ascensão'] = [mes for i in ascensoes.index]
    registro_ascensoes = pd.concat([registro_ascensoes, superaram_300])
    registro_ascensoes.drop_duplicates(subset=['Cliente'], inplace=True)
    registro_ascensoes.sort_values("Mês de Ascensão", inplace=True)

    #filtrando positiadores para clientes problemáticos

    is_conta_nova = posi['Status conta'] == 'conta nova'
    is_menor_300 = posi['Net Em M'] < 3e5
    posi = posi[['Assessor', 'Cliente', 'Net Em M', 'Data de Cadastro']]
    posi = posi[is_conta_nova & is_menor_300]
    posi['Entrada'] = [mes for i in posi.index]

    posi = add_assessor_relacionamento(posi, clientes_rodrigo, column_conta = 'Cliente')

    posi = add_assessor_indicador(posi, clientes_rodrigo)

    posi = reorder_columns(posi, 'Assessor Indicador', 0)

    posi = reorder_columns(posi, 'Assessor Relacionamento', 2)

    print(f'\n\nNovos problemáticos do mês de {mes}:\n', posi)

    base_dados = atualiza_relacionamento_indicador(base_dados, clientes_rodrigo)

    base_dados = base_dados.append(posi)

    print(f'\n\nBase de dados de {mes}:\n', base_dados)

    writer = pd.ExcelWriter(caminho_relatorio , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')

    tabelas = (
        (base_dados, 'Dataset'),
        (superaram_300, 'Superaram 300k'),
        (clientes_sairam, 'Clientes Perdidos'),
        (posi, 'Clientes Novos')
    )

    for tabela, sheet_name in tabelas:
        tabela.to_excel(writer , sheet_name= sheet_name , index= False)

    writer.save()

    print('Novos Transf atualizados')

    gera_excel(registro_ascensoes, r"Aviso Novos Transf\Ascensões\Ascensões_acumulado.xlsx")

def distribution():
    import os
    import shutil

    stock_path = r"landing_point\stock"

    generator = os.walk(stock_path)

    files_result = [x for x in generator]

    new_files = files_result[0][2]

    for file_name in new_files:

        # destina o positivador

        if file_name[0:11] == "positivador":

            where_to = "captacao_diario"        
            current_day, current_month, current_year = file_name[18:20], file_name[16:18], file_name[14:16]
            current_date = "".join(["_",current_day,current_month,current_year])
        
            new_name = stock_path + "\positivador" + current_date + ".xlsx"
            complete_old_name = stock_path + "\\" + file_name

            os.rename(complete_old_name, new_name)

            src_dir_name = new_name
            dst_dir_name = new_name.replace('landing_point\stock', where_to)

            shutil.move(src_dir_name, dst_dir_name)

            print('positivador transferido')

        #destina o diversificador

        if file_name[0:14] == "diversificacao":

            where_to = r"COE\\arquivos"  

            current_day, current_month, current_year = file_name[21:23], file_name[19:21], file_name[17:19]
            current_date = "".join(["_",current_day,current_month,current_year])
        
            new_name = stock_path + "\diversificacao" + current_date + ".xlsx"
            complete_old_name = stock_path + "\\" + file_name

            os.rename(complete_old_name, new_name)

            src_dir_name = new_name
            dst_dir_name = new_name.replace('landing_point\stock', where_to)

            shutil.move(src_dir_name, dst_dir_name)

            print('diversificador transferido')

        # destina a clientes do rodrigo

        if file_name == "Clientes do Rodrigo.xlsx":

            where_to = r"bases_dados"

            src_dir_name = r"landing_point\stock\Clientes do Rodrigo.xlsx"
            dst_dir_name = r"bases_dados\Clientes do Rodrigo.xlsx"

            shutil.move(src_dir_name, dst_dir_name)

            print('clientes rodrigo transferido')

        # destina a suitability

        if file_name == "Suitability.xlsx":

            where_to = r"bases_dados"

            src_dir_name = r"landing_point\stock\Suitability.xlsx"
            dst_dir_name = r"bases_dados\Suitability.xlsx"

            shutil.move(src_dir_name, dst_dir_name)

            print('suitability transferido')

        # movimentações de fundos

        if file_name[:23] == "relatório_movimentações":
            
            where_to = r"COE\\Movimentações de Fundos"  

            stop = input("""\nO nome do fundo monitorado é SPX Private Equity I Advisory FIP MM?
            1 --> Sim
            2 --> Não
            """)

            if stop == '2':
                print('\nAltere o nome do fundo a ser redirecionado\n')
                exit()

            new_name = stock_path + "\Movimentações SPX Private Equity I Advisory FIP MM.xlsx"
            complete_old_name = stock_path + "\\" + file_name

            os.rename(complete_old_name, new_name)

            src_dir_name = new_name
            dst_dir_name = new_name.replace('landing_point\stock', where_to)

            shutil.move(src_dir_name, dst_dir_name)

            print('Movimentação de Fundos transferido')

        # relatorio de transferencias

        if file_name[:20] == 'ExportTransfAssessor':
            
            where_to = "captacao_diario"        
            current_day, current_month, current_year = file_name[18:20], file_name[16:18], file_name[14:16]
            current_date = input('Dia do Relatório de transferências ')
        
            new_name = stock_path + "\\transferencias_" + current_date + ".xlsx"
            complete_old_name = stock_path + "\\" + file_name

            os.rename(complete_old_name, new_name)

            src_dir_name = new_name
            dst_dir_name = new_name.replace('landing_point\stock', where_to)

            shutil.move(src_dir_name, dst_dir_name)

            print('transferencia transferido')

def confere_bases_b2b(positivador, mes, suitability, A_digital = '26839', A_exclusive = '26994'):
    
    ## importando os dados

    base_digital = pd.read_excel(f"Relatórios Digital\Base Clientes\distribuicao_clientes_digital_{str.lower(mes)}.xlsx", sheet_name='Ativo')
    base_exclusive = pd.read_excel(f"Relatórios Digital\Base Clientes\distribuicao_clientes_exclusive_{str.lower(mes)}.xlsx", sheet_name='Ativo')

    #tratamento dos dados

    suitability = suitability[['CodigoBolsa', 'NomeCliente']]
    suitability['CodigoBolsa'] = suitability['CodigoBolsa'].astype(str)

    mask_nao_zerado = positivador['Net Em M'] >= 1e3
    mask_ativo = positivador['Status'] == 'ATIVO'
    positivador = positivador[mask_nao_zerado & mask_ativo]

    cols = ['Assessor', 'Cliente', 'Responsável']
    base_digital = base_digital[cols]
    base_exclusive = base_exclusive[cols]

    positivador['Assessor'] = positivador['Assessor'].astype(str)
    base_digital['Assessor'] = base_digital['Assessor'].astype(str)
    base_exclusive['Assessor'] = base_exclusive['Assessor'].astype(str)

    positivador['Cliente'] = positivador['Cliente'].astype(str)
    base_digital['Cliente'] = base_digital['Cliente'].astype(str)
    base_exclusive['Cliente'] = base_exclusive['Cliente'].astype(str)

    # execução do código
    
    mask_digital = positivador['Assessor'] == A_digital
    mask_exclusive = positivador['Assessor'] == A_exclusive

    positivador_digital = positivador[mask_digital]
    positivador_exclusive = positivador[mask_exclusive]

    mask_in_digital = positivador_digital['Cliente'].isin(base_digital['Cliente'])
    mask_in_exclusive = positivador_exclusive['Cliente'].isin(base_exclusive['Cliente'])

    nao_listados_digital = positivador_digital['Cliente'][~mask_in_digital]
    nao_listados_exclusive = positivador_exclusive['Cliente'][~mask_in_exclusive]

    nao_listados_digital = pd.DataFrame(nao_listados_digital.reset_index(drop=True))
    nao_listados_exclusive = pd.DataFrame(nao_listados_exclusive.reset_index(drop=True))

    nao_listados_digital = nao_listados_digital.merge(suitability, how='left', left_on='Cliente', right_on='CodigoBolsa')
    nao_listados_exclusive = nao_listados_exclusive.merge(suitability, how='left', left_on='Cliente', right_on='CodigoBolsa')

    del nao_listados_digital['CodigoBolsa']
    del nao_listados_exclusive['CodigoBolsa']

    print(f"\n\nClientes não listados do Digital: ({len(nao_listados_digital)} clientes)\n", nao_listados_digital)
    print(f"\n\nClientes não listados do Exclusive: ({len(nao_listados_exclusive)} clientes)\n", nao_listados_exclusive)

    return nao_listados_exclusive, nao_listados_digital

def envia_avisos_clientes_b2b(nao_listados_exclusive, nao_listados_digital, 
    emails_exclusive = ['pamella.teixeira@fatorialinvest.com.br', 'rodrigo.cabral@fatorialinvest.com.br', 'moises.rodrigues@fatorialinvest.com.br', 'soraya.brum@fatorialinvest.com.br', 'alex.marinho@fatorialinvest.com.br'],
    emails_digital = ['soraya.brum@fatorialinvest.com.br', 'alex.marinho@fatorialinvest.com.br', 'rodrigo.cabral@fatorialinvest.com.br', 'moises.rodrigues@fatorialinvest.com.br'],
    adress_mail = 'leonardo.motta@fatorialadvisors.com.br'):
    
    import io

    ## clientes do digital

    total_clientes_digital = len(nao_listados_digital)

    if total_clientes_digital != 0:

        html_df = nao_listados_digital.to_html(index=False, justify = 'center')

        # corpo do email
        body = f"""
        <p>Olá, Time Digital<p>
        
        <p>Aqui é o Leonardo Motta, da Inteligência da Fatorial.<p>

        <p>Fizemos um levantamento e, de acordo com a última base de separação de contas, os seguintes clientes estão sem responsável designado:<p>

        <p>{html_df}<p>

        <p><p>Att,<p>
        <p> Leonardo Gonçalves Motta <p>
        """

    else:
        body = f"""
        <p>Olá, Time Digital<p>
        
        <p>Aqui é o Leonardo Motta, da Inteligência da Fatorial.<p>

        <p>Fizemos um levantamento e, de acordo com a última base de separação de contas, não há clientes sem responsável designado.<p>

        <p><p>Att,<p>
        <p> Leonardo Gonçalves Motta <p>
        """

    send_mail(
        mail_from=adress_mail,
        mail_to=emails_digital,
        subject="[Inteligência Fatorial] Atualização da Base de Clientes Digital",
        body=body,
        mensagem='\n\nE-mail enviado para o Digital.'
    )

    ## clientes do exclusive

    total_clientes_exclusive = len(nao_listados_exclusive)

    if total_clientes_exclusive != 0:

        html_df = nao_listados_exclusive.to_html(index=False, justify = 'center')

        # corpo do email
        body = f"""
        <p>Olá, Time Exclusive<p>
        
        <p>Aqui é o Leonardo Motta, da Inteligência da Fatorial.<p>

        <p>Fizemos um levantamento e, de acordo com a última base de separação de contas, os seguintes clientes estão sem responsável designado:<p>

        <p>{html_df}<p>

        <p><p>Att,<p>
        <p> Leonardo Gonçalves Motta <p>
        """

    else:
        body = f"""
        <p>Olá, Time Exclusive<p>
        
        <p>Aqui é o Leonardo Motta, da Inteligência da Fatorial.<p>

        <p>Fizemos um levantamento e, de acordo com a última base de separação de contas, não há clientes sem responsável designado.<p>

        <p><p>Att,<p>
        <p> Leonardo Gonçalves Motta <p>
        """

    send_mail(
        mail_from=adress_mail,
        mail_to=emails_exclusive,
        subject="[Inteligência Fatorial] Atualização da Base de Clientes Exclusive",
        body=body,
        mensagem='\n\nE-mail enviado para o Exclusive.'
    )

def get_meses():
    return ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

def get_posi_b2c(data_hoje, celula, generate_excel=False):

    obj_data = Data(str(data_hoje))
    mes = obj_data.text_month

    positivador = pd.read_excel (f"captacao_diario\positivador_{data_hoje}.xlsx",skiprows=2)
    base = pd.read_excel(f"Relatórios Digital\Base Clientes\distribuicao_clientes_" + str.lower(celula) + f"_{str.lower(mes)}.xlsx", sheet_name='Ativo')
    suitability = pd.read_excel(r"bases_dados\Suitability.xlsx")
    clientes_rodrigo = pd.read_excel (r"bases_dados\Clientes do Rodrigo.xlsx",sheet_name="Troca")
    assessores = pd.read_excel (r"bases_dados\Assessores leal_Pablo.xlsx")

    caminho_excel = f"Relatórios Digital\Positivadores\positivador_{celula}_{data_hoje}.xlsx"
    clientes_rodrigo['Conta'] = clientes_rodrigo['Conta'].astype(str)

    suitability['CodigoBolsa'] = suitability['CodigoBolsa'].astype(str)

    cols = ['Cliente', 'Responsável']
    base = base[cols]

    positivador['Assessor'] = positivador['Assessor'].astype(str)
    assessores['Código assessor'] = assessores['Código assessor'].astype(str)

    positivador['Cliente'] = positivador['Cliente'].astype(str)
    base['Cliente'] = base['Cliente'].astype(str)

    cod_cels = {
        'digital':'26839',
        'exclusive':'26994'
    }

    positivador = positivador[positivador['Status'] == 'ATIVO']
    positivador = positivador[positivador['Net Em M'] != 0]
    positivador = positivador[positivador['Assessor'] == cod_cels[str.lower(celula)]]

    positivador = add_assessor_indicador(positivador, clientes_rodrigo)

    suitability = suitability[['CodigoBolsa', 'NomeCliente', 'Telefone', 'Celular', 'EmailCliente']]

    df = suitability.merge(positivador, how='right', left_on='CodigoBolsa', right_on='Cliente')

    del df['CodigoBolsa']

    df = df.merge(base, how='left', on='Cliente')

    df = add_nome_assessor(df, 'Assessor Indicador', assessores)

    df.rename(columns={'Nome assessor': 'Indicador'}, inplace=True)

    df  = df[['Indicador', 'Assessor', 'Responsável', 'Cliente', 'NomeCliente', 'Profissão', 
        'Telefone', 'Celular', 'EmailCliente', 'Sexo', 'Segmento', 'Data de Cadastro',
        'Fez Segundo Aporte?', 'Data de Nascimento',
        'Status', 'Ativou em M?', 'Evadiu em M?', 'Operou Bolsa?',
        'Operou Fundo?', 'Operou Renda Fixa?', 'Aplicação Financeira Declarada',
        'Receita no Mês', 'Receita Bovespa', 'Receita Futuros',
        'Receita RF Bancários', 'Receita RF Privados', 'Receita RF Públicos',
        'Captação Bruta em M', 'Resgate em M', 'Captação Líquida em M',
        'Captação TED', 'Captação ST', 'Captação OTA', 'Captação RF',
        'Captação TD', 'Captação PREV', 'Net em M-1', 'Net Em M',
        'Net Renda Fixa', 'Net Fundos Imobiliários', 'Net Renda Variável',
        'Net Fundos', 'Net Financeiro', 'Net Previdência', 'Net Outros',
        'Receita Aluguel', 'Receita Complemento/Pacote Corretagem',
        'Status conta']]

    if generate_excel: 
        gera_excel(df, caminho_excel)
    else: 
        return df

def envia_captacao(data_hoje, mail_to, adress_mail = 'leonardo.motta@fatorialadvisors.com.br'):

    obj_data = Data(data_hoje)
    dia, mes, ano = obj_data.day, obj_data.month, obj_data.year

    captacao = pd.read_excel(f'captacao_diario\captacao_{data_hoje}.xlsx', sheet_name='Resumo')

    posi_exclusive = get_posi_b2c(data_hoje, 'exclusive')

    posi_digital = get_posi_b2c(data_hoje, 'digital')
    
    caminho_excel = r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\resumo_carteiras.xlsx'

    resumo = captacao[['Código assessor', 'Nome assessor', 'Time', 'Qtd Clientes XP', 'NET XP']]

    def get_resumo_carteira_b2c(posi):
        resumo_net = posi[['Responsável', 'Net Em M']].groupby('Responsável').sum()
        resumo_clients = posi[['Responsável', 'Cliente']].groupby('Responsável').count()
        resumo = resumo_clients.merge(resumo_net, how='left', right_index=True, left_index=True)
        resumo.rename(columns={'Net Em M': 'NET XP', 'Cliente': 'Qtd. Cliente XP'}, inplace=True)
        return resumo

    resumo_digital = get_resumo_carteira_b2c(posi_digital)
    resumo_excluisive = get_resumo_carteira_b2c(posi_exclusive)

    writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
    resumo.to_excel(writer, sheet_name='Resumo')
    resumo_digital.to_excel(writer, sheet_name='Resumo Digital')
    resumo_excluisive.to_excel(writer, sheet_name='Resumo Exclusive')
    writer.save()

    def format_html(resumo, index=False):
        html_df = resumo.copy()
        if 'Código assessor' in html_df.columns:
            html_df.drop('Código assessor', axis=1, inplace=True)
        html_df.sort_values(by='NET XP', ascending = False, inplace=True)
        html_df['NET XP'] = html_df['NET XP'].apply(lambda x: 'R$ {:,.2f}'.format(x))
        html_df = html_df.to_html(justify='center', index=index)
        return html_df

    html_df = format_html(resumo)
    html_df_digital = format_html(resumo_digital, index=True)
    html_df_exclusive = format_html(resumo_excluisive, index=True)

    body = f'''
    <p>
    Olá, Cabral! Tudo bem?
    </p>

    <p>
    Segue abaixo um resumo com as carteiras de cada assessor.
    </p>
    
    <p> <b> Toda Fatorial </b> </p>
    {html_df}

    <p>
    Aqui segue as distribuições do digital e do exclusive:
    </p>

    <p> <b> Digital </b> </p>
    {html_df_digital}
    <p> <b> Exclusive </b> </p>
    {html_df_exclusive}

    <p>
    Att,
    </p>

    <img src="https://i.ibb.co/qMt33Ny/Assinatura-Analista-Leonardo.jpg" alt="Assinatura-Analista-Leonardo" border="0">
    '''

    send_mail(
        mail_from=adress_mail,
        mail_to= mail_to,
        subject=f'Resumo das carteiras em {dia}/{mes}/{ano}',
        body=body,
        attachment=caminho_excel,
        mensagem='Resumo de Carteiras Enviado.'
        )
