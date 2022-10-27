import streamlit as st
import pandas as pd
import datetime
import sys
sys.path.insert (1, r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\Funções')
from classes import St
from config import bar_height, columns_proportion, month, adms, receitas, captacao

def app(name, captacao=captacao):

    # header

    st.write("""
    # Captação Líquida 2022
    """)

    St.espacamento(2,st)

    # content

    if name in adms:
        nomes_filtrados = st.multiselect('Assessor', captacao['Nome assessor'].drop_duplicates().fillna(captacao['Código assessor']).sort_values())

    else:
        nomes_filtrados = [name]

    col1, col2 = st.columns(columns_proportion)

    # captacao

    if not nomes_filtrados == []:
        captacao = captacao[ captacao['Nome assessor'].isin(nomes_filtrados) ]

    dict_month = pd.DataFrame( [[ month[i] , datetime.datetime(2022, i+1, 1).strftime("%m - %B")] for i in range(12)] , columns = ['Mes', 'Month'])

    captacao = captacao.merge(dict_month, how='left', on='Mes')

    resumo_captacao = captacao.groupby('Month').sum()[['Captação Líquida', 'Net Em M']]
    resumo_captacao.rename(columns={'Net Em M' : 'Carteira'}, inplace=True)
    resumo_captacao['Captação Acumulada'] = resumo_captacao['Captação Líquida'].cumsum()
    resumo_captacao = resumo_captacao[['Captação Líquida', 'Captação Acumulada', 'Carteira']]
    exporting_file_cap = resumo_captacao.copy()
    resumo_captacao['Captação Acumulada'] = resumo_captacao['Captação Acumulada'].apply(lambda x: 'R$ {:,.2f}'.format(x))
    resumo_captacao['Carteira'] = resumo_captacao['Carteira'].apply(lambda x: 'R$ {:,.2f}'.format(x))
    resumo_captacao['Captação Líquida'] = resumo_captacao['Captação Líquida'].apply(lambda x: 'R$ {:,.2f}'.format(x))

    captacao_plot = captacao.groupby('Month').sum()[['Captação Líquida']]

    col1.bar_chart(captacao_plot, height=bar_height)

    col2.write(resumo_captacao)

    df = St.to_excel(exporting_file_cap)

    col2.download_button(
        label='Exportar Captação',
        data = df,
        file_name = f'Captacao 2022.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )