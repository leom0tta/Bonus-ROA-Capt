import pandas as pd
from pathlib import Path
import pickle
import streamlit_authenticator as stauth

path_to_vscode = r"C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos"

def user_authentication():

    assessores = pd.read_excel ( path_to_vscode + r"\bases_dados\Assessores leal_Pablo.xlsx")

    names = assessores['Nome assessor'].to_list()
    usernames = assessores['Código assessor'].astype(str).to_list()

    file_path = Path(__file__).parent / "hashed_pw.pkl"
    with file_path.open("rb") as file:
        hashed_passwords = pickle.load(file)

    authenticator = stauth.Authenticate(names, usernames, hashed_passwords, "assessores_fatorial", "abcdef", cookie_expiry_days=0)

    name, authentication_status, username = authenticator.login("Login", "main")
    
    return name, authentication_status, username, authenticator

name, authentication_status, username, authenticator = user_authentication()