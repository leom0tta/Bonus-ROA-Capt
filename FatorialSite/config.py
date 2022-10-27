import pandas as pd

bar_height=500
columns_proportion=[5,2.7]
month = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
adms = ['Rodrigo Cabral', 'Jansen Costa', 'Leonardo Motta', 'Moises']

captacao = pd.read_excel(r"C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\Bases de Performance\Base Dados\captacao_2022.xlsx")
receitas = pd.read_excel(r"C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\Bases de Performance\Base Dados\receitas_2022.xlsx", sheet_name='Resumo Tags')