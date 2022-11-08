import pandas as pd
import numpy as np

assessores = pd.read_excel('Assessores-Leal Pablo.xlsx')

path = r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\ModeloBônusCaptação\BD\métricas_captação.xlsx'

df = pd.read_excel(path, sheet_name='df')
price = pd.read_excel(path, sheet_name='price')
average_df = pd.read_excel(path, sheet_name='average_df', index_col='Nome assessor')

average_df.columns = [str(col) for col in average_df.columns]



