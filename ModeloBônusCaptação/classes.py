<<<<<<< HEAD
import pandas as pd

class Dataframe:
    def __init__(self, my_dataframe):
        self.dataframe = my_dataframe

    def rows(self):
        return len(self.dataframe.index)

    def cols(self):
        return len(self.dataframe.columns)


    def add_assessor_indicador(self, clientes_rodrigo, column_conta='Cliente', column_assessor='Assessor', assessores_com_A = False, positivador=None, inplace=True, com_obs=False):
        """Essa função  adiciona a coluna de assessor relacionamento a um dataframe com coluna de contas"""
        dataframe = self.dataframe
        if assessores_com_A == False:
            clientes_rodrigo['Assessor Indicador'] = clientes_rodrigo['Assessor Indicador'].str.strip('A')
            clientes_rodrigo['Assessor Indicador'] = clientes_rodrigo['Assessor Indicador'].str.replace('tendimento Fatorial', 'Atendimento Fatorial')
        if com_obs:
            clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Indicador', 'OBS']]
        else:
            clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Indicador']]
        if column_conta == 'Conta':
            dataframe = dataframe.merge(clientes_rodrigo, how='left', on=column_conta)
        else:
            dataframe = dataframe.merge(clientes_rodrigo, how='left', left_on=column_conta, right_on='Conta')
            dataframe.drop(['Conta'], axis=1, inplace=True)

        null = dataframe['Assessor Indicador'].isnull().to_numpy()
        clientes = dataframe[column_conta].to_numpy()
        
        for i, *_ in enumerate(dataframe['Assessor Indicador'].to_numpy()):        
            is_null = null[i]
            if is_null:
                
                if column_assessor != None:
                    assessor_indicador = dataframe.loc[i, column_assessor]
                    dataframe.loc[i,'Assessor Indicador'] = assessor_indicador
        
                elif column_assessor == None:
                    cliente_selecionado = clientes[i]
                    mask_cliente = positivador['Cliente'] == cliente_selecionado
                    assessor_indicador = positivador.loc[mask_cliente , 'Assessor'].iloc[0]
                    dataframe.loc[i,'Assessor Indicador'] = assessor_indicador

        new_dataframe = dataframe.copy()
        
        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

    def add_assessor_relacionamento(self, clientes_rodrigo, column_conta='Cliente', positivador = None, column_assessor='Assessor', assessores_com_A = False, inplace=True):
        """Essa função adiciona a coluna de assessor indicador a um dataframe com coluna de contas"""

        dataframe = self.dataframe

        if assessores_com_A == False:
            clientes_rodrigo['Assessor Relacionamento'] = clientes_rodrigo['Assessor Relacionamento'].str.strip('A')
            clientes_rodrigo['Assessor Relacionamento'] = clientes_rodrigo['Assessor Relacionamento'].str.replace('tendimento Fatorial', 'Atendimento Fatorial')
        clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Relacionamento']]
        if column_conta == 'Conta':
            dataframe = dataframe.merge(clientes_rodrigo, how='left', on=column_conta)
        else:
            dataframe = dataframe.merge(clientes_rodrigo, how='left', left_on=column_conta, right_on='Conta')
            dataframe.drop(['Conta'], axis=1, inplace=True)

        null = dataframe['Assessor Relacionamento'].isnull().to_numpy()
        clientes = dataframe[column_conta].to_numpy()
        
        for i, *_ in enumerate(dataframe['Assessor Relacionamento'].to_numpy()):        
            is_null = null[i]
            if is_null:
                
                if column_assessor != None:
                    assessor_relacionamento = dataframe.loc[i, column_assessor]
                    dataframe.loc[i,'Assessor Relacionamento'] = assessor_relacionamento
        
                elif column_assessor == None:
                    cliente_selecionado = clientes[i]
                    mask_cliente = positivador['Cliente'] == cliente_selecionado
                    assessor_relacionamento = positivador.loc[mask_cliente , 'Assessor'].iloc[0]
                    dataframe.loc[i,'Assessor Relacionamento'] = assessor_relacionamento
        
        new_dataframe = dataframe.copy()
        
        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

    def reorder_columns(self, col_name, position, inplace=True):
        dataframe = self.dataframe
        """Reorder a dataframe's column.
        Args:
            dataframe (pd.DataFrame): dataframe to use
            col_name (string): column name to move
            position (0-indexed position): where to relocate column to
        Returns:
            pd.DataFrame: re-assigned dataframe
        """
        temp_col = dataframe[col_name]
        dataframe = dataframe.drop(columns=[col_name])
        dataframe.insert(loc=position, column=col_name, value=temp_col)
        
        new_dataframe = dataframe.copy()
        
        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

    def add_nome_cliente(self, suitability, column_conta='Cliente', inplace=True):
        """Essa função adiciona o nome de um cliente, com base na Suitability"""
        dataframe = self.dataframe
        suitability = suitability [['CodigoBolsa', 'NomeCliente']]
        dataframe = dataframe.merge(suitability, how='left', left_on=column_conta, right_on='CodigoBolsa')
        dataframe = dataframe.drop('CodigoBolsa', axis = 1)

        new_dataframe = dataframe.copy()

        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

    def add_nome_assessor(self, assessores, column_assessor='Assessor', inplace=True):
        """Essa função adiciona o nome do assessor, com base na assessores leal pablo"""
        dataframe = self.dataframe
        assessores = assessores [['Código assessor', 'Nome assessor']]
        dataframe = dataframe.merge(assessores, how='left', left_on=column_assessor, right_on='Código assessor')
        if column_assessor != 'Código assessor':
            dataframe = dataframe.drop('Código assessor', axis = 1)
        dataframe['Nome assessor'].fillna(dataframe[column_assessor], inplace=True)

        new_dataframe = dataframe.copy()

        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

class Data:
    def __init__(self, data):
        self.cod_data = data
        self.day = int(data[:2])
        self.month = int(data[2:4])
        self.year = int(data[4:])
        meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        self.text_month = meses[self.month - 1]

class St:

    def __init__(self) -> None:
        pass
    
    def to_excel(df):
        import pandas as pd
        import io
        buffer = io.BytesIO()
        writer = pd.ExcelWriter(buffer, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.save()
        return buffer

    def espacamento(n, self):
        for i in range(n): self.write('')

    def get_aniversario_dataset(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\aniversario_diario\tabela_aniversariantes.xlsx"
        dataframe = pd.read_excel(path, sheet_name='Tabela_Completa')
        dataframe.drop('Data Aviso: 10 Dias de Antecedência', axis=1, inplace=True)
        dataframe['Data de Aniversário'] = dataframe['Data de Aniversário'].dt.strftime("%d/%m/%Y")

class Get:

    def __init__(self) -> None:
        pass

    def aniversario(username = 'Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\aniversario_diario\tabela_aniversariantes.xlsx"
        dataframe = pd.read_excel(path, sheet_name='Tabela_Completa')
        dataframe.drop('Data Aviso: 10 Dias de Antecedência', axis=1, inplace=True)
        dataframe['Data de Aniversário'] = dataframe['Data de Aniversário'].dt.strftime("%d/%m/%Y")
    
        return dataframe

    def assessores(username = 'Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\bases_dados\Assessores leal_Pablo.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        return dataframe

    def suitability(username='Leonardo'):
        suitability = pd.read_excel (r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\bases_dados\Suitability.xlsx')
        suitability['CodigoBolsa'] = suitability['CodigoBolsa'].astype(str)
        return suitability

    def clientes_rodrigo(mes=None, username = 'Leonardo', ano='2022'):
        if mes == None:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\bases_dados\Clientes do Rodrigo.xlsx'
            dataframe = pd.read_excel(path, sheet_name='Troca')
        else:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\Clientes Rodrigo\\' + ano + r'\Clientes Rodrigo ' + mes + r'.xlsx'
            dataframe = pd.read_excel(path)

        dataframe['Conta'] = dataframe['Conta'].astype(str)

        return dataframe

    def captacao(dia, username = 'Leonardo', sheet_name='Resumo'):
        try:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\\captacao_diario\captacao_' + dia + r'.xlsx'
            dataframe = pd.read_excel(path, sheet_name=sheet_name)
        except FileNotFoundError:
            data_obj = Data(dia)
            mes = data_obj.text_month
            ano = '20' + str(data_obj.year)
            
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\\arquivos\\" + ano + r"\\" + mes + r"\\captacao_" + dia + r".xlsx"
            dataframe = pd.read_excel(path, sheet_name=sheet_name)

        if sheet_name == 'Resumo':
            dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        if sheet_name == 'Positivador M':
            dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def positivador(dia, username = 'Leonardo', skiprows=2):
        try:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\\captacao_diario\positivador_' + dia + r'.xlsx'
            dataframe = pd.read_excel(path, skiprows=skiprows)
        except FileNotFoundError:
            data_obj = Data(dia)
            mes = data_obj.text_month
            ano = '20' + str(data_obj.year)

            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\\arquivos\\" + ano + r"\\" + mes + r"\\positivador_" + dia + r".xlsx"
            dataframe = pd.read_excel(path, skiprows=skiprows)
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        dataframe['Assessor'] = dataframe['Assessor'].astype(str)
        return dataframe

    def transferencias(dia, username = 'Leonardo'):
        try:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\\captacao_diario\transferencias_' + dia + r'.xlsx'
            dataframe = pd.read_excel(path)
        except FileNotFoundError:
            data_obj = Data(dia)
            mes = data_obj.text_month
            ano = '20' + str(data_obj.year)

            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\\arquivos\\" + ano + r"\\" + mes + r"\\transferencias_" + dia + r".xlsx"
            dataframe = pd.read_excel(path)
        dataframe = dataframe.loc[dataframe['Status'] =='CONCLUÍDO', :]
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def receitas(mes, username='Leonardo'):
        path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\Comissões\Receitas\Bases SplitC\dados_comissão_' + (str.lower(mes)).replace('ç', 'c') + r'.csv'
        dataframe = pd.read_csv(path, decimal=',', sep=';')
        dataframe['Assessor Dono'] = dataframe['Assessor Dono'].str.title()
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def diversificador(data, username='Leonardo'):
        path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\COE\arquivos\diversificacao_' + data + r'.xlsx'
        diversificador = pd.read_excel(path, skiprows=2)
        return diversificador

    def tags_comissoes(username='Leonardo'):
        return pd.read_excel(r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\Comissões\Bases de dados\Tags x Categorias.xlsx')

    def emails(username = 'Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\bases_dados\Emails.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        return dataframe

    def controle_conta_nova(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\arquivos\2022\Novos_Transf_Acumulado.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Assessor'] = dataframe['Assessor'].astype(str)
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        dataframe['Assessor'].replace('Atendimento Fatorial', '26839', inplace=True)
        return dataframe

    def controle_conta_perdida(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\arquivos\2022\Perdidos_Acumulado.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Assessor'] = dataframe['Assessor'].astype(str)
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        dataframe['Assessor'].replace('Atendimento Fatorial', '26839', inplace=True)
        return dataframe

    def nps_aniversario(mes = None, username='Leonardo'):
        if mes != None:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Mensal\\" + mes +"\Ranking Assessores.xlsx"
            dataframe = pd.read_excel(path, skiprows=4)
        else:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Ranking Assessores.xlsx"
            dataframe = pd.read_excel(path, skiprows=4)
        return dataframe
    
    def nps_onboarding(mes = None, username='Leonardo'):
        if mes != None:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Mensal\\" + mes +"\Ranking_onboarding.xlsx"
            dataframe = pd.read_excel(path, skiprows=4)
        else:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Ranking_onboarding.xlsx"
            dataframe = pd.read_excel(path, skiprows=4)
        return dataframe

    def envios_nps(mes = None, username='Leonardo'):
        if mes != None:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Mensal\\" + mes +"\Registro de Envios.xlsx"
            dataframe = pd.read_excel(path, skiprows=2)
        else:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Registro de Envios.xlsx"
            dataframe = pd.read_excel(path, skiprows=2)
        return dataframe

    def historico_novos(mes, username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\G20\Arquivos\Histórico Novos Perdidos " + mes + ".xlsx"
        dataframe = pd.read_excel(path, sheet_name='Novos')
        dataframe['Cliente'] = dataframe['Cliente'].astype(int).astype(str)
        return dataframe

    def historico_perdidos(mes, username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\G20\Arquivos\Histórico Novos Perdidos " + mes + ".xlsx"
        dataframe = pd.read_excel(path, sheet_name='Perdidos')
        dataframe['Cliente'] = dataframe['Cliente'].astype(int).astype(str)
        return dataframe

    def base_digital(mes, username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\Relatórios Digital\Base Clientes\distribuicao_clientes_digital_" + str.lower(mes) + ".xlsx"
        dataframe = pd.read_excel(path, sheet_name='Ativo')
        dataframe = dataframe[['Cliente', 'Responsável']]
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def base_exclusive(mes, username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\Relatórios Digital\Base Clientes\distribuicao_clientes_exclusive_" + str.lower(mes) + ".xlsx"
        dataframe = pd.read_excel(path, sheet_name='Ativo')
        dataframe = dataframe[['Cliente', 'Responsável']]
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def receita_acumulada(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\Bases de Performance\Base Dados\receitas_2022.xlsx"
        dataframe = pd.read_excel(path, sheet_name = 'Resumo Tags')
        dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        return dataframe

    def captacao_acumulada(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\Bases de Performance\Base Dados\captacao_2022.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        return dataframe

=======
import pandas as pd

class Dataframe:
    def __init__(self, my_dataframe):
        self.dataframe = my_dataframe

    def rows(self):
        return len(self.dataframe.index)

    def cols(self):
        return len(self.dataframe.columns)


    def add_assessor_indicador(self, clientes_rodrigo, column_conta='Cliente', column_assessor='Assessor', assessores_com_A = False, positivador=None, inplace=True, com_obs=False):
        """Essa função  adiciona a coluna de assessor relacionamento a um dataframe com coluna de contas"""
        dataframe = self.dataframe
        if assessores_com_A == False:
            clientes_rodrigo['Assessor Indicador'] = clientes_rodrigo['Assessor Indicador'].str.strip('A')
            clientes_rodrigo['Assessor Indicador'] = clientes_rodrigo['Assessor Indicador'].str.replace('tendimento Fatorial', 'Atendimento Fatorial')
        if com_obs:
            clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Indicador', 'OBS']]
        else:
            clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Indicador']]
        if column_conta == 'Conta':
            dataframe = dataframe.merge(clientes_rodrigo, how='left', on=column_conta)
        else:
            dataframe = dataframe.merge(clientes_rodrigo, how='left', left_on=column_conta, right_on='Conta')
            dataframe.drop(['Conta'], axis=1, inplace=True)

        null = dataframe['Assessor Indicador'].isnull().to_numpy()
        clientes = dataframe[column_conta].to_numpy()
        
        for i, *_ in enumerate(dataframe['Assessor Indicador'].to_numpy()):        
            is_null = null[i]
            if is_null:
                
                if column_assessor != None:
                    assessor_indicador = dataframe.loc[i, column_assessor]
                    dataframe.loc[i,'Assessor Indicador'] = assessor_indicador
        
                elif column_assessor == None:
                    cliente_selecionado = clientes[i]
                    mask_cliente = positivador['Cliente'] == cliente_selecionado
                    assessor_indicador = positivador.loc[mask_cliente , 'Assessor'].iloc[0]
                    dataframe.loc[i,'Assessor Indicador'] = assessor_indicador

        new_dataframe = dataframe.copy()
        
        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

    def add_assessor_relacionamento(self, clientes_rodrigo, column_conta='Cliente', positivador = None, column_assessor='Assessor', assessores_com_A = False, inplace=True):
        """Essa função adiciona a coluna de assessor indicador a um dataframe com coluna de contas"""

        dataframe = self.dataframe

        if assessores_com_A == False:
            clientes_rodrigo['Assessor Relacionamento'] = clientes_rodrigo['Assessor Relacionamento'].str.strip('A')
            clientes_rodrigo['Assessor Relacionamento'] = clientes_rodrigo['Assessor Relacionamento'].str.replace('tendimento Fatorial', 'Atendimento Fatorial')
        clientes_rodrigo = clientes_rodrigo[['Conta' , 'Assessor Relacionamento']]
        if column_conta == 'Conta':
            dataframe = dataframe.merge(clientes_rodrigo, how='left', on=column_conta)
        else:
            dataframe = dataframe.merge(clientes_rodrigo, how='left', left_on=column_conta, right_on='Conta')
            dataframe.drop(['Conta'], axis=1, inplace=True)

        null = dataframe['Assessor Relacionamento'].isnull().to_numpy()
        clientes = dataframe[column_conta].to_numpy()
        
        for i, *_ in enumerate(dataframe['Assessor Relacionamento'].to_numpy()):        
            is_null = null[i]
            if is_null:
                
                if column_assessor != None:
                    assessor_relacionamento = dataframe.loc[i, column_assessor]
                    dataframe.loc[i,'Assessor Relacionamento'] = assessor_relacionamento
        
                elif column_assessor == None:
                    cliente_selecionado = clientes[i]
                    mask_cliente = positivador['Cliente'] == cliente_selecionado
                    assessor_relacionamento = positivador.loc[mask_cliente , 'Assessor'].iloc[0]
                    dataframe.loc[i,'Assessor Relacionamento'] = assessor_relacionamento
        
        new_dataframe = dataframe.copy()
        
        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

    def reorder_columns(self, col_name, position, inplace=True):
        dataframe = self.dataframe
        """Reorder a dataframe's column.
        Args:
            dataframe (pd.DataFrame): dataframe to use
            col_name (string): column name to move
            position (0-indexed position): where to relocate column to
        Returns:
            pd.DataFrame: re-assigned dataframe
        """
        temp_col = dataframe[col_name]
        dataframe = dataframe.drop(columns=[col_name])
        dataframe.insert(loc=position, column=col_name, value=temp_col)
        
        new_dataframe = dataframe.copy()
        
        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

    def add_nome_cliente(self, suitability, column_conta='Cliente', inplace=True):
        """Essa função adiciona o nome de um cliente, com base na Suitability"""
        dataframe = self.dataframe
        suitability = suitability [['CodigoBolsa', 'NomeCliente']]
        dataframe = dataframe.merge(suitability, how='left', left_on=column_conta, right_on='CodigoBolsa')
        dataframe = dataframe.drop('CodigoBolsa', axis = 1)

        new_dataframe = dataframe.copy()

        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

    def add_nome_assessor(self, assessores, column_assessor='Assessor', inplace=True):
        """Essa função adiciona o nome do assessor, com base na assessores leal pablo"""
        dataframe = self.dataframe
        assessores = assessores [['Código assessor', 'Nome assessor']]
        dataframe = dataframe.merge(assessores, how='left', left_on=column_assessor, right_on='Código assessor')
        if column_assessor != 'Código assessor':
            dataframe = dataframe.drop('Código assessor', axis = 1)
        dataframe['Nome assessor'].fillna(dataframe[column_assessor], inplace=True)

        new_dataframe = dataframe.copy()

        if inplace: setattr(self, 'dataframe', new_dataframe)
        return new_dataframe

class Data:
    def __init__(self, data):
        self.cod_data = data
        self.day = int(data[:2])
        self.month = int(data[2:4])
        self.year = int(data[4:])
        meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        self.text_month = meses[self.month - 1]

class St:

    def __init__(self) -> None:
        pass
    
    def to_excel(df):
        import pandas as pd
        import io
        buffer = io.BytesIO()
        writer = pd.ExcelWriter(buffer, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.save()
        return buffer

    def espacamento(n, self):
        for i in range(n): self.write('')

    def get_aniversario_dataset(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\aniversario_diario\tabela_aniversariantes.xlsx"
        dataframe = pd.read_excel(path, sheet_name='Tabela_Completa')
        dataframe.drop('Data Aviso: 10 Dias de Antecedência', axis=1, inplace=True)
        dataframe['Data de Aniversário'] = dataframe['Data de Aniversário'].dt.strftime("%d/%m/%Y")

class Get:

    def __init__(self) -> None:
        pass

    def aniversario(username = 'Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\aniversario_diario\tabela_aniversariantes.xlsx"
        dataframe = pd.read_excel(path, sheet_name='Tabela_Completa')
        dataframe.drop('Data Aviso: 10 Dias de Antecedência', axis=1, inplace=True)
        dataframe['Data de Aniversário'] = dataframe['Data de Aniversário'].dt.strftime("%d/%m/%Y")
    
        return dataframe

    def assessores(username = 'Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\bases_dados\Assessores leal_Pablo.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        return dataframe

    def suitability(username='Leonardo'):
        suitability = pd.read_excel (r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\bases_dados\Suitability.xlsx')
        suitability['CodigoBolsa'] = suitability['CodigoBolsa'].astype(str)
        return suitability

    def clientes_rodrigo(mes=None, username = 'Leonardo', ano='2022'):
        if mes == None:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\bases_dados\Clientes do Rodrigo.xlsx'
            dataframe = pd.read_excel(path, sheet_name='Troca')
        else:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\Clientes Rodrigo\\' + ano + r'\Clientes Rodrigo ' + mes + r'.xlsx'
            dataframe = pd.read_excel(path)

        dataframe['Conta'] = dataframe['Conta'].astype(str)

        return dataframe

    def captacao(dia, username = 'Leonardo', sheet_name='Resumo'):
        try:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\\captacao_diario\captacao_' + dia + r'.xlsx'
            dataframe = pd.read_excel(path, sheet_name=sheet_name)
        except FileNotFoundError:
            data_obj = Data(dia)
            mes = data_obj.text_month
            ano = '20' + str(data_obj.year)
            
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\\arquivos\\" + ano + r"\\" + mes + r"\\captacao_" + dia + r".xlsx"
            dataframe = pd.read_excel(path, sheet_name=sheet_name)

        if sheet_name == 'Resumo':
            dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        if sheet_name == 'Positivador M':
            dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def positivador(dia, username = 'Leonardo', skiprows=2):
        try:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\\captacao_diario\positivador_' + dia + r'.xlsx'
            dataframe = pd.read_excel(path, skiprows=skiprows)
        except FileNotFoundError:
            data_obj = Data(dia)
            mes = data_obj.text_month
            ano = '20' + str(data_obj.year)

            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\\arquivos\\" + ano + r"\\" + mes + r"\\positivador_" + dia + r".xlsx"
            dataframe = pd.read_excel(path, skiprows=skiprows)
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        dataframe['Assessor'] = dataframe['Assessor'].astype(str)
        return dataframe

    def transferencias(dia, username = 'Leonardo'):
        try:
            path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\\captacao_diario\transferencias_' + dia + r'.xlsx'
            dataframe = pd.read_excel(path)
        except FileNotFoundError:
            data_obj = Data(dia)
            mes = data_obj.text_month
            ano = '20' + str(data_obj.year)

            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\\arquivos\\" + ano + r"\\" + mes + r"\\transferencias_" + dia + r".xlsx"
            dataframe = pd.read_excel(path)
        dataframe = dataframe.loc[dataframe['Status'] =='CONCLUÍDO', :]
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def receitas(mes, username='Leonardo'):
        path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\Comissões\Receitas\Bases SplitC\dados_comissão_' + (str.lower(mes)).replace('ç', 'c') + r'.csv'
        dataframe = pd.read_csv(path, decimal=',', sep=';')
        dataframe['Assessor Dono'] = dataframe['Assessor Dono'].str.title()
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def diversificador(data, username='Leonardo'):
        path = r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\COE\arquivos\diversificacao_' + data + r'.xlsx'
        diversificador = pd.read_excel(path, skiprows=2)
        return diversificador

    def tags_comissoes(username='Leonardo'):
        return pd.read_excel(r'C:\Users\\' + username + r'\Dropbox\Fatorial\Inteligência\Codigos\Comissões\Bases de dados\Tags x Categorias.xlsx')

    def emails(username = 'Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\bases_dados\Emails.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        return dataframe

    def controle_conta_nova(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\arquivos\2022\Novos_Transf_Acumulado.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Assessor'] = dataframe['Assessor'].astype(str)
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        dataframe['Assessor'].replace('Atendimento Fatorial', '26839', inplace=True)
        return dataframe

    def controle_conta_perdida(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\captacao_diario\arquivos\2022\Perdidos_Acumulado.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Assessor'] = dataframe['Assessor'].astype(str)
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        dataframe['Assessor'].replace('Atendimento Fatorial', '26839', inplace=True)
        return dataframe

    def nps_aniversario(mes = None, username='Leonardo'):
        if mes != None:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Mensal\\" + mes +"\Ranking Assessores.xlsx"
            dataframe = pd.read_excel(path, skiprows=4)
        else:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Ranking Assessores.xlsx"
            dataframe = pd.read_excel(path, skiprows=4)
        return dataframe
    
    def nps_onboarding(mes = None, username='Leonardo'):
        if mes != None:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Mensal\\" + mes +"\Ranking_onboarding.xlsx"
            dataframe = pd.read_excel(path, skiprows=4)
        else:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Ranking_onboarding.xlsx"
            dataframe = pd.read_excel(path, skiprows=4)
        return dataframe

    def envios_nps(mes = None, username='Leonardo'):
        if mes != None:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Mensal\\" + mes +"\Registro de Envios.xlsx"
            dataframe = pd.read_excel(path, skiprows=2)
        else:
            path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\NPS\Registro de Envios.xlsx"
            dataframe = pd.read_excel(path, skiprows=2)
        return dataframe

    def historico_novos(mes, username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\G20\Arquivos\Histórico Novos Perdidos " + mes + ".xlsx"
        dataframe = pd.read_excel(path, sheet_name='Novos')
        dataframe['Cliente'] = dataframe['Cliente'].astype(int).astype(str)
        return dataframe

    def historico_perdidos(mes, username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\G20\Arquivos\Histórico Novos Perdidos " + mes + ".xlsx"
        dataframe = pd.read_excel(path, sheet_name='Perdidos')
        dataframe['Cliente'] = dataframe['Cliente'].astype(int).astype(str)
        return dataframe

    def base_digital(mes, username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\Relatórios Digital\Base Clientes\distribuicao_clientes_digital_" + str.lower(mes) + ".xlsx"
        dataframe = pd.read_excel(path, sheet_name='Ativo')
        dataframe = dataframe[['Cliente', 'Responsável']]
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def base_exclusive(mes, username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\Relatórios Digital\Base Clientes\distribuicao_clientes_exclusive_" + str.lower(mes) + ".xlsx"
        dataframe = pd.read_excel(path, sheet_name='Ativo')
        dataframe = dataframe[['Cliente', 'Responsável']]
        dataframe['Cliente'] = dataframe['Cliente'].astype(str)
        return dataframe

    def receita_acumulada(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\Bases de Performance\Base Dados\receitas_2022.xlsx"
        dataframe = pd.read_excel(path, sheet_name = 'Resumo Tags')
        dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        return dataframe

    def captacao_acumulada(username='Leonardo'):
        path = r"C:\Users\\" + username + r"\Dropbox\Fatorial\Inteligência\Codigos\Bases de Performance\Base Dados\captacao_2022.xlsx"
        dataframe = pd.read_excel(path)
        dataframe['Código assessor'] = dataframe['Código assessor'].astype(str)
        return dataframe

>>>>>>> 6af27889406a2aeb6985966860c67dc148dcd244
