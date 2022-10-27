import pandas as pd
import sys
sys.path.insert (1, r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\Funções')
import numpy as np

assessores = Get.assessores()

path = r'C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\ModeloBônusCaptação\BD\métricas_captação.xlsx'

df = pd.read_excel(path, sheet_name='df')
price = pd.read_excel(path, sheet_name='price')
average_df = pd.read_excel(path, sheet_name='average_df', index_col='Nome assessor')

average_df.columns = [str(col) for col in average_df.columns]



