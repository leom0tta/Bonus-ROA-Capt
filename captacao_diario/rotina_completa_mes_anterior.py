import sys
sys.path.insert(1, r'.\Funções')
from funcoes import distribution, export_files_captacao_mes_passado, rotina_coe, monitora_vencimentos_RF,export_files_captacao, filtro_positivador, captacao, gera_pipeline, ranking_diario, avisos_novos_transf, confere_bases_b2b, envia_avisos_clientes_b2b
distribution()

data_hoje = '280722'
data_ontem = '260722'
data_pipe = '070722'
mes_atual = 'Julho'
enviar_email_digital_exclusive = True

dataframes = export_files_captacao_mes_passado(data_hoje, data_ontem, mes_atual, '300622')

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

base_digital, base_exclusive = confere_bases_b2b(posi_novo, mes_atual, suitability)

if enviar_email_digital_exclusive: 
    envia_avisos_clientes_b2b(base_digital, base_exclusive)

print('\nRotina Finalizada\n')
