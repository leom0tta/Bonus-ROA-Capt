import streamlit as st
import pandas as pd
import datetime
import sys
sys.path.insert (1, r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\Funções')
from classes import St
from config import bar_height, columns_proportion, month, adms, receitas, captacao
from login_page import name

def app(name, receitas=receitas, captacao=captacao):

    # header

    st.write("""
    # Receita Bruta 2022
    """)

    St.espacamento(2,st)

    # content

    if name in adms:
        nomes_filtrados = st.multiselect('Assessor', receitas['Nome assessor'].drop_duplicates().fillna(receitas['Código assessor']).sort_values())
        print(nomes_filtrados)

    else:
        nomes_filtrados = [name]

    col1, col2 = st.columns(columns_proportion)

    # captacao

    if not nomes_filtrados == []:
        receitas = receitas[ receitas['Nome assessor'].isin(nomes_filtrados) ]
        captacao = captacao[ captacao['Nome assessor'].isin(nomes_filtrados) ]

    dict_month = pd.DataFrame( [[ month[i] , datetime.datetime(2022, i+1, 1).strftime("%m - %B")] for i in range(12)] , columns = ['Mes', 'Month'])

    receitas = receitas.merge(dict_month, how='left', on='Mes')
    captacao = captacao.merge(dict_month, how='left', on='Mes')

    carteira = captacao.groupby('Month').sum()[['Net Em M']]
    carteira.columns = ['Carteira']
    resumo_receita = receitas.groupby('Month').sum()[['Receita']]
    resumo_receita['Receita Acumulada'] = resumo_receita['Receita'].cumsum()
    resumo_receita = resumo_receita.merge(carteira, how='left', left_index=True, right_index=True)
    resumo_receita['ROA'] = resumo_receita['Receita'] / resumo_receita['Carteira'] * 12
    export_files_rec = resumo_receita.copy()
    resumo_receita['Receita Acumulada'] = resumo_receita['Receita Acumulada'].apply(lambda x: 'R$ {:,.2f}'.format(x))
    resumo_receita['Receita'] = resumo_receita['Receita'].apply(lambda x: 'R$ {:,.2f}'.format(x))
    resumo_receita['Carteira'] = resumo_receita['Carteira'].apply(lambda x: 'R$ {:,.2f}'.format(x))
    resumo_receita['ROA'] = resumo_receita['ROA'].apply(lambda x: '{:,.3f} %'.format(x))

    receitas_plot = receitas.groupby('Month').sum()[['Receita']]

    col1.bar_chart(receitas_plot, height=bar_height)

    col2.write(resumo_receita)

    df = St.to_excel(export_files_rec)

    col2.download_button(
        label='Exportar Receita',
        data = df,
        file_name = f'Receita 2022.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
