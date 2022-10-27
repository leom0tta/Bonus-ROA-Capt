import streamlit as st
from multipage import MultiApp
import captacao_page
import receita_page
from login_page import authentication_status, authenticator
#import my_account

path_to_vscode = r"C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos"

# settings

st.set_page_config(page_title= 'Fatorial Investimentos',page_icon = '📍',layout='wide')

# --- USER AUTHENTICATION ---

if authentication_status == False:
    st.error("Username/password is incorrect")

if authentication_status:

    authenticator.logout("Logout", 'main')

    apps = MultiApp()

    apps.add_app("Captação", captacao_page.app)
    apps.add_app("Receita", receita_page.app)
    
    apps.run()

else:
    from login_page import authenticator
    authenticator.logout("Logout", 'main')

