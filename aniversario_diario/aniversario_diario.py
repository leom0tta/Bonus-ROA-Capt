import pandas as pd
import numpy
from datetime import date, timedelta
from pathlib import Path
import win32com.client as win32
import sys
import os
import subprocess
sys.path.insert(1, r'.\Funções')
from funcoes import add_assessor_indicador, gera_excel
from classes import Get

ano_corrente = 2022 #! colocar o ano atual
data_hoje = '211022'
data_ontem = '201022'
enviar_email = False


# Importando Base de Dados

# Filtrando Código, Nome e Email e Relacionando com Positivador, para ter a data de Nascimento 
# É necessário atualizar as versões das planilhas

suitability = Get.suitability()
positivador = Get.positivador(data_hoje)
positivador_ontem = Get.positivador(data_ontem)
clientes_rodrigo = Get.clientes_rodrigo()
captacao_hj = Get.captacao(data_hoje, sheet_name="Positivador M")
assessores_leal = Get.assessores()

caminho_excel = Path(r"aniversario_diario\tabela_aniversariantes.xlsx")
caminho_tabela = Path(r"aniversario_diario\tabela_completa.xlsx")

assessor_indicador = clientes_rodrigo.copy ()
assessores = assessores_leal.copy()

clientes_rodrigo = clientes_rodrigo [["Conta","Assessor Relacionamento"]]
clientes_rodrigo.rename (columns={"Conta":"Cliente"},inplace=True)

suitability = suitability [["CodigoBolsa","NomeCliente","EmailCliente"]]
positivador = positivador[["Assessor","Cliente","Data de Nascimento"]]
positivador_ontem = positivador_ontem[["Assessor","Cliente","Data de Nascimento"]]

positivador.dropna(subset=['Data de Nascimento'], inplace=True)

suitability.rename (columns={"CodigoBolsa":"Cliente"},inplace=True)

tabela_completa = pd.merge (suitability,positivador,how="inner",on="Cliente")

tabela_completa ["Data de Aniversário"] = tabela_completa ["Data de Nascimento"].apply (lambda x: date (ano_corrente,x.month,x.day) ).astype (numpy.datetime64)

tabela_completa.drop ("Data de Nascimento",axis=1,inplace=True)

tabela_completa ["Data Aviso: 10 Dias de Antecedência"] = tabela_completa ["Data de Aniversário"] - timedelta (days=10)

dia_de_hoje = date.today ().strftime ("%Y-%m-%d")

tabela_aviso = tabela_completa.loc [tabela_completa["Data Aviso: 10 Dias de Antecedência"]==dia_de_hoje,:]

tabela_aviso = pd.merge (tabela_aviso,clientes_rodrigo,how="left",on="Cliente")

tabela_aviso ["Assessor Relacionamento"].fillna (tabela_aviso["Assessor"],inplace=True)

tabela_aviso.drop ("Assessor",axis=1,inplace=True)

tabela_aviso ["Assessor Relacionamento"] = tabela_aviso ["Assessor Relacionamento"].astype (str)

tabela_aviso.loc [tabela_aviso["Assessor Relacionamento"]!="Atendimento Fatorial","Assessor Relacionamento"] = tabela_aviso.loc [tabela_aviso["Assessor Relacionamento"]!="Atendimento Fatorial","Assessor Relacionamento"].apply (lambda x: x.replace ("A",""))

assessores_leal ["Código assessor"] = assessores_leal ["Código assessor"].astype (str)  

tabela_aviso = pd.merge (tabela_aviso,assessores_leal,how="left",left_on="Assessor Relacionamento",right_on="Código assessor")

tabela_aviso.drop (["Código assessor","Time"],axis=1,inplace=True)

tabela_aviso ["Nome assessor"].fillna ("Atendimento Fatorial",inplace=True)

captacao_hj.rename (columns={"Assessor correto":"Assessor Relacionamento"},inplace=True)

captacao_hj ["Assessor Relacionamento"] = captacao_hj ["Assessor Relacionamento"].astype (str)

tabela_aviso = pd.merge (tabela_aviso,captacao_hj[["Cliente","Assessor Relacionamento","Net Em M"]],how="left",on="Cliente")

tabela_aviso.drop ("Assessor Relacionamento_y",axis=1,inplace=True)

tabela_aviso.rename (columns={"Assessor Relacionamento_x":"Assessor Relacionamento"},inplace=True)

tabela_aviso.sort_values (by="Net Em M",ascending=False,inplace=True)

# Juntando Tabela Aviso com clientes_rodrigo para descobrir o Assessor Indicador

assessor_indicador.rename (columns={"Conta":"Cliente"},inplace=True)

tabela_aviso = pd.merge (tabela_aviso,assessor_indicador[["Assessor Indicador","Cliente"]],how="left",on="Cliente")

tabela_aviso["Assessor Indicador"].fillna (tabela_aviso["Assessor Relacionamento"],axis=0,inplace=True)

tabela_aviso ["Assessor Indicador"] = tabela_aviso ["Assessor Indicador"].apply (lambda x: x.replace ("A",""))

tabela_aviso = tabela_aviso [["Nome assessor","Assessor Relacionamento","Cliente","NomeCliente","EmailCliente","Net Em M","Assessor Indicador","Data de Aniversário","Data Aviso: 10 Dias de Antecedência"]]

# Tabela Aniversariantes

dia_de_hoje = date.today ().strftime ("%Y-%m-%d")

tabela_aniversariantes = tabela_completa.loc [tabela_completa["Data de Aniversário"]==dia_de_hoje,:]
    
tabela_aniversariantes.drop ("Data Aviso: 10 Dias de Antecedência",axis=1,inplace=True)

tabela_aniversariantes = pd.merge (tabela_aniversariantes,clientes_rodrigo,how="left",on="Cliente")

tabela_aniversariantes ["Assessor Relacionamento"].fillna (tabela_aniversariantes["Assessor"],inplace=True)

tabela_aniversariantes ["Assessor Relacionamento"] = tabela_aniversariantes ["Assessor Relacionamento"].astype (str) 

tabela_aniversariantes.loc [tabela_aniversariantes["Assessor Relacionamento"]!="Atendimento Fatorial","Assessor Relacionamento"] = tabela_aniversariantes.loc [tabela_aniversariantes["Assessor Relacionamento"]!="Atendimento Fatorial","Assessor Relacionamento"].apply (lambda x: x.replace ("A",""))

tabela_aniversariantes = pd.merge (tabela_aniversariantes,assessores_leal,how="left",left_on="Assessor Relacionamento",right_on="Código assessor")

tabela_aniversariantes.drop (["Código assessor","Time"],axis=1,inplace=True)

tabela_aniversariantes ["Nome assessor"].fillna (tabela_aniversariantes["Assessor Relacionamento"],inplace=True)

tabela_aniversariantes.drop ("Assessor",axis=1,inplace=True)

tabela_aniversariantes = pd.merge (tabela_aniversariantes,captacao_hj[["Cliente","Net Em M"]],how="left",on="Cliente")

tabela_aniversariantes.sort_values (by="Net Em M",ascending=False,inplace=True)

tabela_aniversariantes = pd.merge (tabela_aniversariantes,assessor_indicador[["Cliente","Assessor Indicador"]],how="left",on="Cliente")

tabela_aniversariantes ["Assessor Indicador"].fillna (tabela_aniversariantes["Assessor Relacionamento"],axis=0,inplace=True)

tabela_aniversariantes ["Assessor Indicador"] = tabela_aniversariantes ["Assessor Indicador"].apply (lambda x: x.replace ("A",""))

tabela_aniversariantes = tabela_aniversariantes [["Nome assessor","Assessor Relacionamento","Cliente","NomeCliente","EmailCliente","Net Em M","Assessor Indicador","Data de Aniversário"]]


# Colocando nome do indicador na tabela aviso

assessores_leal.rename (columns={"Código assessor":"Assessor Indicador"},inplace=True)

assessores_leal.rename (columns={"Nome assessor":"Nome assessor Indicador"},inplace=True)

assessores_leal ["Assessor Indicador"] = assessores_leal ["Assessor Indicador"].astype (str)

tabela_aviso ["Assessor Indicador"] = tabela_aviso ["Assessor Indicador"].astype (str)

tabela_aviso = pd.merge (tabela_aviso,assessores_leal[["Assessor Indicador","Nome assessor Indicador"]],how="left",on="Assessor Indicador")

tabela_aviso ["Nome assessor Indicador"].fillna (tabela_aviso["Assessor Indicador"],inplace=True)

tabela_aviso = tabela_aviso [["Nome assessor","Assessor Relacionamento","Cliente","NomeCliente","EmailCliente","Net Em M","Assessor Indicador","Nome assessor Indicador","Data de Aniversário","Data Aviso: 10 Dias de Antecedência"]]

tabela_aviso.rename (columns={"Nome assessor":"Nome assessor Relacionamento"},inplace=True)

# Colocando nome do indicador na tabela aniversariantes

tabela_aniversariantes ["Assessor Indicador"] = tabela_aniversariantes ["Assessor Indicador"].astype (str)

tabela_aniversariantes = pd.merge (tabela_aniversariantes,assessores_leal[["Assessor Indicador","Nome assessor Indicador"]],how="left",on="Assessor Indicador")

tabela_aniversariantes.rename (columns={"Nome assessor":"Nome assessor Relacionamento"},inplace=True)

tabela_aniversariantes ["Nome assessor Indicador"].fillna (tabela_aniversariantes["Assessor Indicador"],inplace=True)


tabela_aniversariantes = tabela_aniversariantes [["Nome assessor Relacionamento","Assessor Relacionamento","Cliente","NomeCliente","EmailCliente","Net Em M","Assessor Indicador","Nome assessor Indicador","Data de Aniversário"]]

# coloca o assessor indicador na tabela completa

assessor_indicador.rename (columns={"Cliente":"Conta"},inplace=True)
tabela_completa = add_assessor_indicador(tabela_completa, assessor_indicador)
tabela_completa['Assessor Indicador'] = tabela_completa['Assessor Indicador'].astype(str)

# ajusta as colunas das tabelas de aviso e aniversariantes

tabela_aviso.drop('EmailCliente', axis=1, inplace=True)
tabela_aviso.drop('Net Em M', axis=1, inplace=True)
tabela_aniversariantes.drop('EmailCliente', axis=1, inplace=True)
tabela_aniversariantes.drop('Net Em M', axis=1, inplace=True)


writer = pd.ExcelWriter(caminho_excel , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
tabela_completa.to_excel(writer , sheet_name='Tabela_Completa',index=False)
tabela_aviso.to_excel(writer , sheet_name='Tabela_Aviso10D',index=False)
tabela_aniversariantes.to_excel(writer , sheet_name='Tabela_Hoje',index=False)

#Definir as tabelas na planilha através de um loop

lista_tabelas = [
(tabela_completa, 'Tabela_Completa', 'Table Style Medium 2'), 
(tabela_aviso, 'Tabela_Aviso10D', 'Table Style Medium 2'), 
(tabela_aniversariantes, 'Tabela_Hoje', 'Table Style Medium 2')
]

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
 
writer.save ()

print ("arquivo criado")

if enviar_email == True:

    def open_outlook():
        try:
            subprocess.call(['C:\Program Files\Microsoft Office\Office15\Outlook.exe'])
            os.system("C:\Program Files\Microsoft Office\Office15\Outlook.exe")
        except:
            print("Outlook didn't open successfully")

    # atualiza o arquivo que tem só a tabela
    gera_excel(tabela_completa, caminho_tabela)

    # envio do email com o arquivo criado para a Lara

    informations = [['Renata Schneider', 'renata.schneider@fatorialinvest.com.br']]
    assessor = informations[0][0]
    email_destino= informations[0][1]

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')
    mapi = outlook.GetNameSpace('MAPI')

    # criar um email
    email = outlook.CreateItem(0)
    email.To = email_destino
    email._oleobj_.Invoke(*(64209,0,8,0,mapi.Accounts.Item('3613leo@gmail.com')))
    email.Subject = "Aniversários Fatorial"

    # adiciona assinatura
    file_path = str(caminho_tabela.absolute())
    email.Attachments.Add(file_path)

    # corpo do email
    body = f"""
    <p>Olá, {assessor}<p>

    <p>Aqui é o Leonardo Motta, da Inteligência da Fatorial.<p>

    <p>Segue em anexo o documento que compila todos os aniversários de clientes 
    da fatorial, atualizado hoje.<p>

    <p>Qualquer dúvida estou à disposição.<p>

    <p><p>Att,<p>
    <p>Leonardo Gonçalves Motta<p>
    """

    # configurar as informações do seu e-mail
    email.HTMLBody = body
    
    email.display()
    email.Send()
    print(f"\nEmail Enviado para {assessor}\n")

