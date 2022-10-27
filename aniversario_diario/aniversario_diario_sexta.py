import pandas as pd
from datetime import datetime
import numpy
from datetime import date
from datetime import datetime, timedelta
import xlsxwriter as xlsx
from pathlib import Path

# A lógica é fazer o mesmo trabalho do código "aniversário_diario", só que para os 3 dias seguintes, em 3 planilhas

ano_corrente = 2022 #! colocar o ano atual
dia_recente = '211022'
DMenos1 = '201022'

# Define os diretórios de cada planilha e os dias de hoje, amanhã e depois de amanhã (sexta sábado e domingo))

caminho_sexta_excel = r"aniversario_diario\sexta_aniversarios.xlsx"
caminho_sabado_excel = r"aniversario_diario\sabado_aniversarios.xlsx"
caminho_domingo_excel = r"aniversario_diario\domingo_aniversarios.xlsx"

sexta = date.today () - timedelta(days= 3.0)
sabado = sexta + timedelta(days = 1.0)
domingo = sabado + timedelta(days = 1.0)

lista_de_dias = [
    (sexta, caminho_sexta_excel), 
    (sabado, caminho_sabado_excel), 
    (domingo, caminho_domingo_excel)
    ]

lista_dias_str = []

for dia, caminho_excel in lista_de_dias:
    lista_dias_str.append(
        (dia.strftime ("%Y-%m-%d"), caminho_excel )
        )    

# Realiza os mesmos procesos de criação da planilha antiga, mas uma vez para cada dia, gerando uma planilha em cada diretório

for dia_de_hoje, caminho_excel in lista_dias_str:

    # Importando Base de Dados

    # Filtrando Código, Nome e Email e Relacionando com Positivador, para ter a data de Nascimento 

    suitability = pd.read_excel (r"bases_dados\Suitability.xlsx")
    positivador = pd.read_excel (r"captacao_diario\positivador_" + dia_recente + ".xlsx",skiprows=2)
    positivador_ontem = pd.read_excel (r"captacao_diario\positivador_" + DMenos1 + ".xlsx",skiprows=2)
    clientes_rodrigo = pd.read_excel (r"bases_dados\Clientes do Rodrigo.xlsx",sheet_name="Troca")
    captacao_hj = pd.read_excel (r"captacao_diario\captacao_" + dia_recente + ".xlsx",sheet_name="Positivador M")
    assessores_leal = pd.read_excel (r"bases_dados\Assessores leal_Pablo.xlsx")

    assessor_indicador = clientes_rodrigo.copy ()

    clientes_rodrigo = clientes_rodrigo [["Conta","Assessor Relacionamento"]]
    clientes_rodrigo.rename (columns={"Conta":"Cliente"},inplace=True)

    suitability = suitability [["CodigoBolsa","NomeCliente","EmailCliente"]]
    positivador = positivador[["Assessor","Cliente","Data de Nascimento"]]

    suitability.rename (columns={"CodigoBolsa":"Cliente"},inplace=True)

# coloca os aniversários que vem faltando

    sem_aniversario = positivador[positivador['Data de Nascimento'].isnull()]

    clientes_sem_aniversario = sem_aniversario['Cliente'].to_numpy()

    clientes_ontem = positivador_ontem['Cliente'].to_numpy()

    for cliente in clientes_sem_aniversario:
        if cliente in clientes_ontem:
            aniversario_cliente = positivador_ontem.loc[ positivador_ontem['Cliente'] == cliente, 'Data de Nascimento' ].to_numpy()[0]
            positivador.loc[ positivador['Cliente'] == cliente, 'Data de Nascimento' ] = aniversario_cliente

    print("\nClientes sem aniversário:\n\n", positivador[positivador['Data de Nascimento'].isnull()], '\n')

    positivador.dropna(subset=['Data de Nascimento'], inplace=True)

    tabela_completa = pd.merge (suitability,positivador,how="inner",on="Cliente")

    tabela_completa ["Data de Aniversário"] = tabela_completa ["Data de Nascimento"].apply (lambda x: date (ano_corrente,x.month,x.day) ).astype (numpy.datetime64)

    tabela_completa.drop ("Data de Nascimento",axis=1,inplace=True)

    tabela_completa ["Data Aviso: 10 Dias de Antecedência"] = tabela_completa ["Data de Aniversário"] - timedelta (days=10)
    
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

