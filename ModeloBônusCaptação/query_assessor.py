import streamlit as st
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from datasets import df, assessores

def app(df=df):
    st.write('''
    # Análise Individual por Assessor
    # ''')

    assessor = st.selectbox('Selecione o assessor: ', assessores['Nome assessor'].drop_duplicates().sort_values(), index=0)

    df = df[df['Nome assessor'] == assessor]

    average_capt = sum(df['Captação Líquida'])/len(df['Captação Líquida'])
    perc_capt_ext = sum(df['Captação Externa'])/sum(df['Captação Líquida'])

    average_ROA = sum(df['ROA'])/len(df['ROA'])
    average_fat = sum(df['Receita'])/len(df['Receita'])

    col1, col2, col3 = st.columns(3)

    col1.metric("Captação Mensal", 'R$ {:,.2f}'.format(average_capt))
    col2.metric("Captação Anualizada", 'R$ {:,.2f}'.format(average_capt*12))
    col3.metric("% Captação Externa", '{:,.2f} %'.format(perc_capt_ext*100))

    col1.metric("Receita Mensal", 'R$ {:,.2f}'.format(average_fat))
    col2.metric("Receita Anualizada", 'R$ {:,.2f}'.format(average_fat*12))
    col3.metric("ROA", '{:,.2f} %'.format(average_ROA*100))

    fig = make_subplots(rows=1, cols=2, specs=[[{"secondary_y": True}, {"secondary_y": True}]])

    fig.add_trace(
        go.Bar(x=df['Mes'], y=df['Captação Líquida'], name='Captação Líquida'),
        row=1, col=1
    )

    fig.add_trace(
        go.Scatter(x=df['Mes'], y=df['Net Em M'], name='Carteira'), 
        row=1, col=1, secondary_y=True
    )

    fig.add_trace(
        go.Bar(x=df['Mes'], y=df['Receita'], name='Faturamento'), 
        row=1, col=2
    )

    fig.add_trace(
        go.Scatter(x=df['Mes'], y=df['ROA'], name='ROA'), 
        row=1, col=2, secondary_y=True
    )

    st.plotly_chart(fig, use_container_width=True)