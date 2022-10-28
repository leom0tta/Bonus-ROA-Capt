
import streamlit as st
st.set_page_config(page_title= 'Modelo de Receita' ,page_icon = 'F',layout='wide')


from multipage import MultiApp
import query_assessor
import analysis

apps = MultiApp()

apps.add_app('Análise por Assessor', query_assessor.app)
apps.add_app('Análises e Simulações', analysis.app)

apps.run()


