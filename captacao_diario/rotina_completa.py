import sys
sys.path.insert(1, r'.\Funções')
from funcoes import distribution, envia_captacao, rotina_coe, monitora_vencimentos_RF,export_files_captacao, filtro_positivador, captacao, gera_pipeline, ranking_diario, avisos_novos_transf, confere_bases_b2b, envia_avisos_clientes_b2b
from classes import Data
import warnings
warnings.filterwarnings("ignore")

distribution()

data_hoje_obj = Data('241022')
data_ontem_obj = Data('211022')
enviar_emails = True
data_pipe = '070722'

data_hoje = data_hoje_obj.cod_data
data_ontem = data_ontem_obj.cod_data
mes_atual = data_hoje_obj.text_month

'''dataframes = export_files_captacao(data_hoje, data_ontem)

posi_novo = dataframes[0]
posi_d1 = dataframes[1]
posi_velho = dataframes[2]
clientes_rodrigo = dataframes[3]
assessores = dataframes[4]
lista_transf = dataframes[5]
clientes_novos_ontem = dataframes[6]
suitability = dataframes[7]
registro_transf = dataframes[8]

filtro_positivador(posi_novo, posi_d1, posi_velho, clientes_rodrigo, lista_transf, registro_transf, data_hoje)

captacao(posi_novo, posi_velho, clientes_rodrigo, assessores, lista_transf, clientes_novos_ontem, suitability, data_hoje)

pipeline = gera_pipeline(data_pipe, assessores)

ranking_diario(pipeline, data_hoje)

rotina_coe(data_hoje, mes_atual)

base_digital, base_exclusive = confere_bases_b2b(posi_novo, mes_atual, suitability)'''

if enviar_emails: 
    
    #envia_avisos_clientes_b2b(base_digital, base_exclusive)

    envia_captacao(data_hoje, ['rodrigo.cabral@fatorialinvest.com.br', 'obastos@fatorialinvest.com.br', 'pablo.langenbach@fatorialinvest.com.br', 'jansen@fatorialinvest.com.br'])

print('\nRotina Finalizada\n')

print(r'\n C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\ranking_diario\arquivos\2022\ranking_' + data_hoje + '_geral.xlsx')

print(r'\n C:\Users\Leonardo\Dropbox\Fatorial\Inteligência\Codigos\ranking_diario\arquivos\2022\ranking_' + data_hoje + '_filtrado.xlsx')
