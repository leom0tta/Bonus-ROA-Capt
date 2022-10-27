import win32com.client
import pandas as pd
import sys
sys.path.insert(1, r'.\Funções')
from funcoes import reorder_columns

mes = 'Setembro'
adress_mail = 'leonardo.motta@fatorialadvisors.com.br'
show='B2B'

def export_dataset(mes):
    rank_g20 = pd.read_excel(r"G20\Relatorio\Ranking G20 " + mes + ".xlsx")
    emails = pd.read_excel(r"bases_dados\Emails.xlsx", sheet_name='Emails')

    emails['Ass'] = emails['Ass'].astype(str)
    rank_g20['Código assessor'] = rank_g20['Código assessor'].astype(str)

    return rank_g20, emails

def filtra_dataset(rank_g20, show='everyone'):
    if show == 'everyone':
        rank_g20 = rank_g20[ rank_g20['Time'].isin(['B2B', 'B2C', 'Mesa RV'])]
        rank_g20 = rank_g20[~rank_g20['Nome assessor'].isin(['Lucas Mattoso', 'Rodolfo', 'Yago Meireles' ,'Pablo Langenbach', 'Base Fatorial','Aldeir Dovales'])]
        del rank_g20['Posição Geral']
    
    elif show == 'B2B':
        rank_g20 = rank_g20[ rank_g20['Time'].isin(['B2B'])]
        rank_g20 = rank_g20[~rank_g20['Nome assessor'].isin(['Jansen Costa', 'Octavio Bastos'])]
        del rank_g20['Posição Geral']
    
    posicoes = rank_g20['Pontuação Geral'].sort_values(ascending=False).reset_index(drop=True).drop_duplicates()
    posicoes = posicoes.reset_index()

    rank_g20 = rank_g20.merge(posicoes, how='left', on='Pontuação Geral')
    rank_g20.rename(columns={'index': 'Posição Geral'}, inplace=True)
    rank_g20['Posição Geral'] += 1

    ult_col = len(rank_g20.columns) - 1
    rank_g20 = reorder_columns(rank_g20, 'Pontuação Geral', ult_col)

    rank_g20['Captação Líquida'] = rank_g20['Captação Líquida'].apply(lambda x: 'R${:,.2f}'.format(x))
    rank_g20['Faturamento'] = rank_g20['Faturamento'].apply(lambda x: 'R${:,.2f}'.format(x))
    rank_g20['Contas Novas'] = rank_g20['Contas Novas'].astype(int)
    rank_g20['NPS Aniversário'] = rank_g20['NPS Aniversário'].round(2)
    rank_g20['NPS Onboarding'] = rank_g20['NPS Onboarding'].round(2)
    rank_g20['Percentual de Resposta'] = rank_g20['Percentual de Resposta'].apply(lambda x: '{:,.2f}%'.format(x*100))
    rank_g20['Fator de Peso'] = rank_g20['Fator de Peso'].apply(lambda x: '{:,.2f}%'.format(x*100))
    rank_g20['Posição Geral'] = rank_g20['Posição Geral'].apply(lambda x: "{:}ª Posição".format(x))
    rank_g20['Pontuação Geral'] = rank_g20['Pontuação Geral'].round(2)

    return rank_g20

def get_emails_array(rank_g20, emails):

    emails_list = []
    sem_email = []

    for cod_assessor in rank_g20['Código assessor']:
        nome_assessor = rank_g20.loc[ rank_g20['Código assessor'] == cod_assessor, 'Nome assessor'].values[0]
        try:
            email = emails.loc[ emails['Ass'] == cod_assessor, 'E-mail'].values[0]
            emails_list += [[nome_assessor,email]]
        except IndexError: # não foi registrado
            sem_email += [nome_assessor]
    
    if len(sem_email) > 0:
        print('\nOs seguintes assessores não receberão e-mails, deseja prosseguir?')
        for assessor in sem_email: print(assessor)
        print('1- Sim\n2- Não')
        resposta = input('Resposta: ')
        if resposta == '1':
            return emails_list
        if resposta == '2':
            exit()
    else:
        return emails_list

def build_body(rank_g20, nome_assessor, mes):

    information = rank_g20.loc[ rank_g20['Nome assessor'] == nome_assessor, :]
    del information['Nome assessor']
    del information['Código assessor']
    del information['Time']

    pont_geral, posi_geral = information['Pontuação Geral'].values[0], information['Posição Geral'].values[0]

    information = information.transpose()
    information.columns = ['Valor']

    information.reset_index(inplace=True, drop=False)

    mask_valor = information.index % 2 == 0
    mask_peso = information.index % 2 != 0

    table = pd.DataFrame([], index=['Captação', 'Faturamento', 'Contas Novas', 'NPS Aniversário', 'NPS Onboarding', 'Percentual de Respostas', 'Final'])

    table['Valor'] = information.loc[mask_valor, 'Valor'].to_numpy()
    table['Pontuação'] = information.loc[mask_peso, 'Valor'].to_numpy()
    table['Pesos'] = ['30%', '30%', '30%', '5%', '5%', 'Multiplica o Total', 'Final']
    html_table = table.to_html(justify = 'center')

    list_nomes = nome_assessor.split(' ')
    primeiro_nome = list_nomes[0].lower().capitalize()

    body = f'''
    <body>
    <font>
    <p> 
    Boa tarde, {primeiro_nome}! 
    </p>
    
    <p>
    No mês de {mes}, ranqueamos os assessores quanto ao seu desempenho na Fatorial, considerando: Captação, Receita, Abertura e Perda de Contas e NPS.
    <p>
    
    <p>
    Nesse ranking você ocupa a <b> {posi_geral} Geral</b>, acumulando {pont_geral} pontos.
    </p>

    <p>
    Abaixo, você pode ver os pontos de cada categoria:
    </p>

    <p>
    {html_table}
    </p>

    <p>
    O cálculo da nota é feito com a soma da pontuação obtida nos cinco primeiros campos, multiplicados pelos seus pesos. 
    O percentual de respostas corresponde a um peso, e esse peso multiplica a soma obtida anteriormente, gerando o resultado final.
    </p>

    <p>
    Qualquer dúvida, estou à disposição!
    </p>

    <p>
    Atenciosamente,
    </p>

    <img src="https://i.ibb.co/JQsTyJF/Assinatura-e-mail-advisors-leonardo.png" alt="Assinatura-e-mail-advisors-leonardo" border="0">

    </font>
    </body>'''

    return body

def disparar_emails(emails_array, adress_mail, mes):
    for nome, mail_to in emails_array:

        body = build_body(rank_g20, nome, mes)

        outlook = win32com.client.Dispatch('outlook.application')
        mapi = outlook.GetNamespace('MAPI')
        
        email = outlook.CreateItem(0)

        email.To = ";".join([mail_to, 'rodrigo.cabral@fatorialinvest.com.br', 'jansen@fatorialinvest.com.br'])
        #email.To = ";".join(['jansen@fatorialinvest.com.br', 'rodrigo.cabral@fatorialinvest.com.br', '3613leo@gmail.com'])
        #email.To = '3613leo@gmail.com'
        email._oleobj_.Invoke(*(64209,0,8,0,mapi.Accounts.Item(adress_mail)))
        email.Subject = f'Sua posição no Super Ranking da Fatorial de {mes}'
        email.HTMLBody = body 
        email.Send()
        print(f'Email enviado para {nome}')

rank_g20, emails = export_dataset(mes=mes)
rank_g20 = filtra_dataset(rank_g20, show=show)
emails_array = get_emails_array(rank_g20, emails)
disparar_emails(emails_array, adress_mail=adress_mail, mes=mes)