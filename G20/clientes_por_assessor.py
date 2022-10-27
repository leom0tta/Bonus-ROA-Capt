import pandas as pd
from pathlib import Path
import win32com.client as win32
import locale

data_hoje = '050522'
enviar_email = False

#importando positivador mês M 
caminho_novo = Path(r'captacao_diario\positivador_' + data_hoje + '.xlsx') #! positivador do dia de run do código
posi_novo = pd.read_excel(caminho_novo, skiprows= 2)

#importando lista de assessores 
caminho_assessores = Path(r'bases_dados\Assessores leal_Pablo.xlsx') #!
assessores = pd.read_excel(caminho_assessores)
assessores = assessores.astype({'Código assessor' : str})

caminho_excel = Path(r'G20\Clientes_por_assessor\Clientes_por_assessor_' + data_hoje + '.xlsx')

caminho_fixo = Path(r'G20\Clientes_por_assessor\Clientes_por_assessor.xlsx') # estático pro power bi

# faz uma listagem , baseada no positivador, da quantidade de clientes por assessor

mask_ativo = posi_novo['Status'] == 'ATIVO'

posi_ativo = posi_novo[mask_ativo]

posi_ativo['Assessor'] = posi_ativo['Assessor'].astype(str)

clientes_por_assessor = posi_ativo.groupby('Assessor')['Cliente'].count()

net_por_assessor = posi_ativo.groupby('Assessor')['Net Em M'].sum()

clientes_por_assessor = pd.concat([clientes_por_assessor, net_por_assessor], axis=1)

clientes_por_assessor = clientes_por_assessor.merge(assessores, how='left', right_on='Código assessor', left_index=True)

clientes_por_assessor.reset_index(drop=True, inplace=True)

clientes_por_assessor = clientes_por_assessor[['Código assessor', 'Nome assessor', 'Cliente', 'Net Em M']]

clientes_por_assessor['Ticket Médio'] = clientes_por_assessor['Net Em M']/clientes_por_assessor['Cliente']

# gera excel com o relatório de clientes por assessor

writer = pd.ExcelWriter(caminho_excel, engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
clientes_por_assessor.to_excel(writer , sheet_name='Sheet1',index=False)

writer.save()

writer = pd.ExcelWriter(caminho_fixo, engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
clientes_por_assessor.to_excel(writer , sheet_name='Sheet1',index=False)

writer.save()

print('arquivo criado')

if enviar_email == True:
    # envia por email para o cabral

    informations = [['Rodrigo Cabral', 'rodrigo.cabral@fatorialinvest.com.br']] 
    assessor = informations[0][0]
    email_destino= informations[0][1]

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)
    email.Subject = "Relação de Clientes Por Assessor"
    email.To = email_destino

    # adiciona assinatura
    file_path = str(caminho_excel.absolute())
    email.Attachments.Add(file_path)

    # corpo do email
    body = f"""
    <p>Olá, {assessor}<p>

    <p>Aqui é o Leonardo Motta, da Inteligência da Fatorial.<p>

    <p>Segue em anexo o documento que registra a relação de clientes por assessores do dia 
    {data_hoje[0:2]}/{data_hoje[2:4]}/{data_hoje[4:]}.<p>

    <p>Qualquer dúvida estou à disposição.<p>

    <p><p>Att,<p>
    <p>Leonardo Gonçalves Motta<p>
    """

    # configurar as informações do seu e-mail
    email.HTMLBody = body
    email.display()

    email.Send()
    print(f"\nEmail Enviado para {assessor}\n")