import pandas as pd
import streamlit as st
import numpy as np
from datasets import average_df, assessores
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from range_capt import range_capt_bonus, range_roa

def app(average_df = average_df ):

    st.write('''
    # Análise e Simulações
    # ''')

    time = st.multiselect('Selecione o Time: ', assessores['Time'].drop_duplicates().sort_values())

    col1, col2, col3 = st.columns([1,.1,1]) 

    if time != []:
        average_df = average_df.loc[ average_df['Time'].isin(time)]

    min_roa = col1.select_slider('Minimum ROA', range_roa*100, format_func= '{:,.2f} %'.format)/100
    capt_bonus = col3.select_slider('Bônus Captação', range_capt_bonus*100, format_func='{:,.2f} %'.format)/100

    num_beneficiados = average_df[str(min_roa.round(3))].sum()
    pagamento_em_bonus = sum(average_df[str(min_roa.round(3))] * average_df['Captação Líquida'] * 12 * capt_bonus)

    col3.metric("Pagamento anual em Bônus", 'R$ {:,.2f}'.format(pagamento_em_bonus))
    col1.metric("Total de Beneficiados", '{:,.0f} Assessores'.format(num_beneficiados))

    @st.cache
    def get_price(average_df):
        price = []
        for roa_min in range_roa:
            price += [
            [
            roa_min, 
            capt,
            sum(12*average_df['Captação Líquida']*average_df[str(roa_min)]*capt), #captação anualizada
            sum(average_df[str(roa_min)]) #assessores beneficiados
            ] 
            for capt in range_capt_bonus]

        return pd.DataFrame(price, columns = ['Roa Min', 'Bônus Capt', 'Preço', 'Qtd. Beneficiados'])

    price = get_price(average_df)

    mask_roa = price['Roa Min'] == round(min_roa,3)
    mask_bonus = price['Bônus Capt'] == capt_bonus

    left_df = price.loc[mask_roa , :]
    right_df = price.loc[mask_bonus , :]

    fig = make_subplots(
        rows=1, cols=2, 
        specs=[[{"secondary_y": True}, {"secondary_y": True}]],
        subplot_titles=[
            f'''<b>Preço em Função do Bônus de Captação, para ROA = {'{:,.2f} %'.format(min_roa*100)}</b>''', 
            f'''<b>Preço em Função do ROA, para Bônus p/ Captação = {'{:,.2f} %'.format(capt_bonus*100)}</b>'''
            ])

    fig.add_trace(
        go.Scatter(x = left_df['Bônus Capt'], y = left_df['Preço']),
        row=1, col=1
    )

    fig.update_xaxes(title_text='Bônus por Captação', row=1, col=1)
    fig.update_yaxes(title_text='Preço', row=1, col=1)


    fig.add_trace(
        go.Scatter(x = right_df['Roa Min'], y = right_df['Preço']),
        row=1, col=2
    )

    fig.update_xaxes(title_text='ROA', row=1, col=2)
    fig.update_yaxes(title_text='Preço', row=1, col=2)

    fig.update_layout(showlegend=False)

    st.plotly_chart(fig, use_container_width=True)

    st.write(f'''
    ### Assessores Beneficiados com ROA mínimo = {'{:,.2f} %'.format(min_roa*100)}''')

    display_df = average_df.loc[average_df[str(round(min_roa,3))]==1, ['Captação Líquida', 'ROA']]

    display_df['Captação Líquida'] = display_df['Captação Líquida'].apply(lambda x: 'R$ {:,.2f}'.format(x))
    display_df['ROA'] = display_df['ROA'].apply(lambda x: '{:,.3f} %'.format(x))

    st.write(display_df)





