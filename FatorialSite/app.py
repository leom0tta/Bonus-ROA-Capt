import streamlit as st
import pandas as pd
from pathlib import Path
import pickle
import sys
import streamlit_authenticator as stauth
sys.path.insert (1, r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\Funções')
from classes import St
import captacao_page
import receita_page

path_to_vscode = r"C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos"

# settings

st.set_page_config(page_title= 'Fatorial Investimentos',page_icon = '📍',layout='wide')

def user_authentication():
    assessores = pd.read_excel ( path_to_vscode + r"\bases_dados\Assessores leal_Pablo.xlsx")

    names = assessores['Nome assessor'].to_list()
    usernames = assessores['Código assessor'].astype(str).to_list()

    file_path = Path(__file__).parent / "hashed_pw.pkl"
    with file_path.open("rb") as file:
        hashed_passwords = pickle.load(file)

    authenticator = stauth.Authenticate(names, usernames, hashed_passwords, "assessores_fatorial", "abcdef", cookie_expiry_days=30)

    name, authentication_status, username = authenticator.login("Login", "main")
    
    return name, authentication_status, username, authenticator

def sidebar():
    sidebar = st.sidebar
    sidebar.write(f'## Bem vindo, {name}!')
    St.espacamento(2, sidebar)
    sidebar.write(f'Selecione o relatório que deseja ver: ')
    St.espacamento(2, sidebar)

    if sidebar.button('Captação'):
        captacao_page.app(name=name)
    if sidebar.button('Receitas'):
        receita_page.app(name=name)

    St.espacamento(8, sidebar)

    sidebar.image("https://d335luupugsy2.cloudfront.net/cms/files/291484/1660076443/$ogiaym3dec ")
    St.espacamento(10, sidebar)
    sidebar.write('Fatorial Investimentos - XP')
    authenticator.logout("Logout", 'sidebar')


# --- USER AUTHENTICATION ---

name, authentication_status, username, authenticator = user_authentication()

if authentication_status == False:
    st.error("Username/password is incorrect")

if authentication_status:
    sidebar()

