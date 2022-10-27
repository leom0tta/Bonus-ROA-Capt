import streamlit as st
import sys
sys.path.insert (1, r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\Funções')
from classes import St
from login_page import name, authenticator

def app():

    # header

    st.write(f"""
    # {name}
    """)

    St.espacamento(2,st)