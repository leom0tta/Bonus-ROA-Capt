from webbrowser import get
import pandas as pd
import numpy as np
import sys
sys.path.insert(1, r'.\Funções')
from funcoes import gera_excel, add_assessor_indicador, reorder_columns
from classes import Data, Get
import warnings
warnings.filterwarnings("ignore")

data_obj = Data('300922')
data_hoje = data_obj.cod_data
mes = data_obj.text_month
responsavel_private = '72260'

# bases net e captação

captacao = Get.captacao(data_hoje)

posi = Get.captacao(data_hoje, sheet_name="Positivador M")

mask_total = captacao['Nome assessor'] == 'Total Fatorial'
captacao = captacao[~mask_total]

# bases receitas

comissoes = Get.receitas(mes)

# bases nps

nps_aniversario = Get.nps_aniversario(mes)
nps_onboarding = Get.nps_onboarding(mes)
envios = Get.envios_nps(mes)

# bases gerais

clientes_rodrigo = Get.clientes_rodrigo(mes)

assessores = Get.assessores()

times = assessores[['Código assessor', 'Time']]

# contas novas e perdidas

contas_novas = Get.controle_conta_nova()
contas_perdidas = Get.controle_conta_perdida()

hist_novos = Get.historico_novos(mes)
hist_perdidos = Get.historico_perdidos(mes)

# distribuição dos times

base_digital = Get.base_digital(mes)

base_exclusive = Get.base_exclusive(mes)

# transferencias

transferencias = Get.transferencias(data_hoje)

caminho_excel = r"G20\Relatorio\Relatório G20 " + mes + ".xlsx"

## divide a captacao e a receita do digital e do exclusive, baseado no positivador

# atualiza a captação líquida no relatório do digital e do exclusive

mask_esse_mes = contas_perdidas['Meses'] == mes
perdidos_esse_mes = contas_perdidas.loc[mask_esse_mes, 'Cliente']

codigo ={
    'Soraya':'Soraya Digital',
    'Alex': '73003',
    'Pamella': '71326',
    'Private': '26877',
    'Brenno': '22222',
    'Bruna Krivochein': '72836'
}
# digital 

base_digital['Responsável'] = base_digital['Responsável'].str.strip()
for responsavel in base_digital['Responsável'].drop_duplicates():
    try:
        base_digital['Responsável'].replace(responsavel, codigo[responsavel], inplace=True)
    except KeyError:
        pass

mask_digital = posi['Assessor'] == 26839
mask_ativo = posi['Net Em M'] != 0
posi_digital = posi[mask_digital & mask_ativo]

posi_digital = posi_digital.merge(base_digital, how='left', on='Cliente')
posi_digital['Responsável'].fillna('Sem Responsável Digital', inplace=True)

contas_perdidas_digital = contas_perdidas[(contas_perdidas['Assessor'] == '26839') & (contas_perdidas['Meses'] == mes)]
contas_perdidas_digital = contas_perdidas_digital[['Assessor', 'Cliente', 'Net Em M']]
contas_perdidas_digital = contas_perdidas_digital.merge(base_digital, how='left', on='Cliente')
contas_perdidas_digital['Responsável'].fillna('Sem Responsável Digital', inplace=True)
contas_perdidas_digital['Captação Líquida em M'] = (-1)*(contas_perdidas_digital['Net Em M'])
contas_perdidas_digital['Assessor'] = contas_perdidas_digital['Responsável']
contas_perdidas_digital['Assessor correto'] = contas_perdidas_digital['Responsável']
del contas_perdidas_digital['Responsável']

posi_digital = pd.concat([posi_digital, contas_perdidas_digital], axis=0)

is_transf = posi_digital['Cliente'].isin(transferencias['Cliente']).to_numpy()

for i, cliente in enumerate(posi_digital['Cliente']):
    if is_transf[i]:
        mask_cliente = posi_digital['Cliente'] == cliente
        net_transferido = posi_digital.loc[mask_cliente, 'Net em M-1'].fillna(0)
        posi_digital.loc[mask_cliente, 'Captação Líquida em M'] += net_transferido

posi_digital['Assessor'] = posi_digital['Responsável'].fillna(posi_digital['Assessor'])

net_digital = posi_digital.groupby('Assessor').sum()[['Captação Líquida em M', 'Net Em M']]
clientes_digital = posi_digital.groupby('Assessor').count()['Cliente']

captacao_digital = net_digital.merge(clientes_digital, how='outer', left_index=True, right_index=True)

captacao_digital.columns = ['Captação Líquida', 'NET XP', 'Qtd Clientes XP']

for responsavel in captacao_digital.index:
    capt, net, num_clientes = captacao_digital.loc[responsavel]
    if responsavel in captacao['Código assessor'].to_numpy():
        captacao.loc[ captacao['Código assessor'] == responsavel, 'Captação Líquida'] += capt
        captacao.loc[ captacao['Código assessor'] == responsavel, 'NET XP'] += net
        captacao.loc[ captacao['Código assessor'] == responsavel, 'Qtd Clientes XP'] += num_clientes
    else:
        dados = captacao_digital.loc[captacao_digital.index == responsavel]
        dados.reset_index(drop=False, inplace=True)
        dados.rename(columns={'Assessor': 'Código assessor'}, inplace=True)
        captacao = pd.concat([captacao, dados], axis=0, ignore_index=False)

codigo['Soraya'] = 'Soraya Exclusive'
# exclusive

base_exclusive['Responsável'] = base_exclusive['Responsável'].str.strip()
for responsavel in base_exclusive['Responsável'].drop_duplicates():
    try:
        base_exclusive['Responsável'].replace(responsavel, codigo[responsavel], inplace=True)
    except KeyError:
        pass

mask_exclusive = posi['Assessor'] == 26994
mask_ativo = posi['Net Em M'] != 0
posi_exclusive = posi[mask_exclusive & mask_ativo]

posi_exclusive = posi_exclusive.merge(base_exclusive, how='left', on='Cliente')
posi_exclusive['Responsável'].fillna('Sem Responsável Exclusive', inplace=True)

contas_perdidas_exclusive = contas_perdidas[(contas_perdidas['Assessor'] == '26994') & (contas_perdidas['Meses'] == mes)]
contas_perdidas_exclusive = contas_perdidas_exclusive[['Assessor', 'Cliente', 'Net Em M']]
contas_perdidas_exclusive = contas_perdidas_exclusive.merge(base_exclusive, how='left', on='Cliente')
contas_perdidas_exclusive['Responsável'].fillna('Sem Responsável Exclusive', inplace=True)
contas_perdidas_exclusive['Captação Líquida em M'] = (-1)*(contas_perdidas_exclusive['Net Em M'])
contas_perdidas_exclusive['Assessor'] = contas_perdidas_exclusive['Responsável']
contas_perdidas_exclusive['Assessor correto'] = contas_perdidas_exclusive['Responsável']
del contas_perdidas_exclusive['Responsável']

posi_exclusive = pd.concat([posi_exclusive, contas_perdidas_exclusive], axis=0)

is_transf = posi_exclusive['Cliente'].isin(transferencias['Cliente']).to_numpy()

for i, cliente in enumerate(posi_exclusive['Cliente']):
    if is_transf[i]:
        mask_cliente = posi_exclusive['Cliente'] == cliente
        net_transferido = posi_exclusive.loc[mask_cliente, 'Net em M-1'].fillna(0)
        posi_exclusive.loc[mask_cliente, 'Captação Líquida em M'] += net_transferido

posi_exclusive['Assessor'] = posi_exclusive['Responsável'].fillna(posi_exclusive['Assessor'])

net_exclusive = posi_exclusive.groupby('Assessor').sum()[['Captação Líquida em M', 'Net Em M']]
clientes_exclusive = posi_exclusive.groupby('Assessor').count()['Cliente']

captacao_exclusive = net_exclusive.merge(clientes_exclusive, how='outer', left_index=True, right_index=True)

captacao_exclusive.columns = ['Captação Líquida', 'NET XP', 'Qtd Clientes XP']

for responsavel in captacao_exclusive.index:
    capt, net, num_clientes = captacao_exclusive.loc[responsavel]
    if responsavel in captacao['Código assessor'].to_numpy():
        captacao.loc[ captacao['Código assessor'] == responsavel, 'Captação Líquida'] += capt
        captacao.loc[ captacao['Código assessor'] == responsavel, 'NET XP'] += net
        captacao.loc[ captacao['Código assessor'] == responsavel, 'Qtd Clientes XP'] += num_clientes
    else:
        dados = captacao_exclusive.loc[captacao_exclusive.index == responsavel]
        dados.reset_index(drop=False, inplace=True)
        dados.rename(columns={'Assessor': 'Código assessor'}, inplace=True)
        captacao = pd.concat([captacao, dados], axis=0, ignore_index=False)

# private

mask_private = captacao['Código assessor'] == '26877'
mask_responsavel = captacao['Código assessor'] == responsavel_private

if responsavel_private in captacao['Código assessor'].to_numpy():
    captacao.loc[mask_responsavel, 'Captação Líquida'] += captacao.loc[mask_private, 'Captação Líquida'].iloc[0]
    captacao.loc[mask_responsavel, 'Qtd Clientes XP'] += captacao.loc[mask_private, 'Qtd Clientes XP'].iloc[0]
    captacao.loc[mask_responsavel, 'NET XP'] += captacao.loc[mask_private, 'NET XP'].iloc[0]


else:
    captacao['Código assessor'].replace('26877', responsavel_private, inplace=True)
    nome_resp_private = assessores.loc[assessores['Código assessor'] == responsavel_private, 'Nome assessor'].iloc[0]

captacao['Ticket Médio'] = captacao['NET XP']/captacao['Qtd Clientes XP']
captacao['Ticket Médio'].fillna(0, inplace=True)

## calcula e coloca o NET indicador 

positivador = posi[['Assessor', 'Cliente', 'Net Em M']]

positivador = add_assessor_indicador(positivador, clientes_rodrigo)

del positivador['Assessor']
del positivador['Cliente']

positivador['Assessor Indicador'] = positivador['Assessor Indicador'].astype(str)

# tira os valores das células e manda para base fatorial
positivador['Assessor Indicador'].replace('26839', 'Base Fatorial', inplace=True)
positivador['Assessor Indicador'].replace('26877', 'Base Fatorial', inplace=True)
positivador['Assessor Indicador'].replace('26994', 'Base Fatorial', inplace=True)

net_indicador = positivador.groupby('Assessor Indicador').sum()
net_indicador = net_indicador.reset_index()
net_indicador['Código assessor'] = net_indicador['Assessor Indicador'].str.title()
del net_indicador['Assessor Indicador']

net_indicador.rename(columns={'Net Em M': 'Net Indicador'}, inplace=True)

relatorio = captacao.merge(net_indicador, how='outer', on='Código assessor')

relatorio['Nome assessor'].fillna(relatorio['Código assessor'], inplace=True)

relatorio = reorder_columns(relatorio, 'Net Indicador', 12)

## extrai as receitas do dataframe de comissões

# substitui o A do atendimento

comissoes['Assessor Dono'].replace('Atendimento Fatorial', 'A26839', inplace=True)

# retira ajustes e erros

lista_categorias_erros = ['Erro Operacional','Incentivo Comercial', 'Ajuste', 'Complemento de Comissão Mínima']
mask_erros = comissoes['Categoria'].isin(lista_categorias_erros)
comissoes = comissoes[~mask_erros]

# corrige o A dos códigos

comissoes['Assessor Dono'] = comissoes['Assessor Dono'].str.lstrip('A')
comissoes['Assessor Dono'].replace('liny Manzieri', 'Aliny Manzieri', inplace=True)
comissoes['Assessor Dono'].replace('lexandre Pessanha', 'Alexandre Pessanha', inplace=True)
comissoes['Assessor Dono'].replace('driano Meneguite', 'Adriano Meneguite', inplace=True)
comissoes['Assessor Dono'].replace('ldeir Dovales', 'Aldeir Dovales', inplace=True)

# separa os clientes do atendimento, do digital e do private

comissoes = comissoes.merge(base_digital, how='left', on='Cliente')
comissoes['Assessor Dono'].replace('26839', np.NaN, inplace=True)
comissoes['Assessor Dono'] = comissoes['Assessor Dono'].fillna(comissoes['Responsável'].fillna('Sem Responsável Digital'))

del comissoes['Responsável']

comissoes = comissoes.merge(base_exclusive, how='left', on='Cliente')
comissoes['Assessor Dono'].replace('26994', np.NaN, inplace=True)
comissoes['Assessor Dono'] = comissoes['Assessor Dono'].fillna(comissoes['Responsável'].fillna('Sem Responsável Exclusive'))

del comissoes['Responsável']

comissoes['Assessor Dono'].replace('26877', responsavel_private, inplace=True)

# agrupa para gerar o dataset de receitas

receitas = comissoes.groupby('Assessor Dono').sum()['Valor Bruto Recebido']
receitas = receitas.reset_index()

## une captação e receita, para calcular o ROA

relatorio = relatorio.merge(receitas, how='outer', left_on='Código assessor', right_on='Assessor Dono')
del relatorio['Time']

relatorio = relatorio.merge(times, how='left', on='Código assessor')
relatorio = reorder_columns(relatorio, 'Time', 2)

relatorio['Código assessor'].fillna(relatorio['Assessor Dono'], inplace=True)
relatorio['Nome assessor'].fillna(relatorio['Assessor Dono'], inplace=True)
relatorio['Time'].fillna('Outros', inplace=True)

relatorio = relatorio[
    ['Código assessor',
    'Nome assessor',
    'Captação Líquida',
    'Ticket Médio',
    'Qtd Clientes XP',
    'Net Indicador',
    'NET XP',
    'Valor Bruto Recebido',
    'Time']
    ]

relatorio['Ticket Médio'] = relatorio['NET XP'] / relatorio['Qtd Clientes XP']

relatorio.rename(columns={'Valor Bruto Recebido': 'Faturamento'}, inplace=True)

relatorio['ROA'] = relatorio['Faturamento'] * 12 / relatorio['NET XP'] 
relatorio['ROA'].replace(np.inf, 0, inplace=True)
relatorio['ROA'].replace(np.NaN, 0, inplace=True)

## compila os clientes novos e perdidos

meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
idx_mes = meses.index(mes)
mes_1 = meses[idx_mes - 1]
mes_2 = meses[idx_mes - 2]
prox_mes = meses[idx_mes + 1]

# clientes novos

# separa os clientes do atendimento, do exclusive e do private

contas_novas = contas_novas.merge(base_digital, how='left', on='Cliente')
contas_novas['Assessor'] = contas_novas['Responsável'].fillna(contas_novas['Assessor'])
contas_novas['Assessor'].replace('26839', 'Sem Responsável Digital')

del contas_novas['Responsável']

contas_novas = contas_novas.merge(base_exclusive, how='left', on='Cliente')
contas_novas['Assessor'] = contas_novas['Responsável'].fillna(contas_novas['Assessor'])
contas_novas['Assessor'].replace('26994', 'Sem Responsável Exclusive')

del contas_novas['Responsável']

contas_novas['Assessor'].replace('26877', responsavel_private, inplace=True)

# segue

mask_ultimos_meses = contas_novas['Meses'].isin([mes, mes_1, mes_2])

contas_novas = contas_novas[mask_ultimos_meses].drop('Net Em M', axis=1)

mask_ja_contou = contas_novas['Cliente'].isin(hist_novos['Cliente'])

contas_novas = contas_novas[~mask_ja_contou]

patrimonio_atualizado = posi[['Cliente', 'Net Em M']]

contas_novas = contas_novas.merge(patrimonio_atualizado, on='Cliente')

contas_1M = contas_novas['Net Em M'] >= 1e6
contas_300k = contas_novas['Net Em M'] >= 3e5

contas_novas['Conta Nova +300k'] = [0 for i in contas_novas.index]
contas_novas.loc[ contas_300k, 'Conta Nova +300k' ] = 1
contas_novas.loc[contas_1M, 'Conta Nova +300k'] = 0

contas_novas['Conta Nova +1M'] = [0 for i in contas_novas.index]
contas_novas['Conta Nova +1M'][contas_1M] = 1

clientes_adicionados = contas_novas[contas_1M | contas_300k]

hist_novos = pd.concat([hist_novos['Cliente'], clientes_adicionados['Cliente']], axis=0).drop_duplicates()

contas_novas = contas_novas.groupby('Assessor').sum()[['Conta Nova +300k', 'Conta Nova +1M']].reset_index()

relatorio = relatorio.merge(contas_novas, how='left', left_on='Código assessor', right_on='Assessor')
del relatorio['Assessor']

# clientes perdidos

# separa os clientes do exclusive e do digital

contas_perdidas = contas_perdidas.merge(base_digital, how='left', on='Cliente')
contas_perdidas['Assessor'] = contas_perdidas['Responsável'].fillna(contas_perdidas['Assessor'])
contas_perdidas['Assessor'].replace('26839', 'Sem Responsável Digital')

del contas_perdidas['Responsável']

contas_perdidas = contas_perdidas.merge(base_exclusive, how='left', on='Cliente')
contas_perdidas['Assessor'] = contas_perdidas['Responsável'].fillna(contas_perdidas['Assessor'])
contas_perdidas['Assessor'].replace('26994', 'Sem Responsável Exclusive')

contas_perdidas['Assessor'].replace('26877', responsavel_private, inplace=True)

# segue

mask_ultimos_meses = contas_perdidas['Meses'].isin([mes, mes_1, mes_2])

contas_perdidas = contas_perdidas[mask_ultimos_meses]

mask_ja_contou = contas_perdidas['Cliente'].isin(hist_perdidos['Cliente'])

contas_perdidas = contas_perdidas[~mask_ja_contou]

contas_1M = contas_perdidas['Net Em M'] >= 1e6
contas_300k = (contas_perdidas['Net Em M'] >= 3e5) & ~contas_1M

contas_perdidas['Conta Perdida +300k'] = [0 for i in contas_perdidas.index]
contas_perdidas['Conta Perdida +300k'][contas_300k] = 1

contas_perdidas['Conta Perdida +1M'] = [0 for i in contas_perdidas.index]
contas_perdidas['Conta Perdida +1M'][contas_1M] = 1

clientes_adicionados = contas_perdidas[contas_1M | contas_300k]

hist_perdidos = pd.concat([hist_perdidos['Cliente'], clientes_adicionados['Cliente']], axis=0).drop_duplicates()

contas_perdidas = contas_perdidas.groupby('Assessor').sum()[['Conta Perdida +300k', 'Conta Perdida +1M']].reset_index()

relatorio = relatorio.merge(contas_perdidas, how='left', left_on='Código assessor', right_on='Assessor')
del relatorio['Assessor']

col_aberturas = ['Conta Nova +300k', 'Conta Nova +1M', 'Conta Perdida +300k', 'Conta Perdida +1M']

relatorio[col_aberturas] = relatorio[col_aberturas].fillna(0)

# atualiza o histórico

writer = pd.ExcelWriter(r'G20\Arquivos\Histórico Novos Perdidos ' + prox_mes + '.xlsx' , engine='xlsxwriter', datetime_format = 'dd/mm/yyyy')
hist_novos.to_excel(writer , sheet_name= 'Novos' , index=False)
hist_perdidos.to_excel(writer , sheet_name= 'Perdidos' , index=False)
writer.save()

## coloca o NPS do mês

# nps aniversario

nps_aniversario['Assessor'] = nps_aniversario['Assessor'].str.lstrip('A')
respostas_aniversario = nps_aniversario[['Assessor', 'Tamanho da amostra']]
nps_aniversario = nps_aniversario[['Assessor', 'XP - Relacionamento - Aniversário - NPS Assessor']]


relatorio = relatorio.merge(nps_aniversario, how='outer', left_on='Código assessor', right_on='Assessor')
del relatorio['Assessor']
relatorio.rename(columns = {'XP - Relacionamento - Aniversário - NPS Assessor' : 'NPS Aniversário'}, inplace=True)

# nps onboarding

nps_onboarding['Assessor'] = nps_onboarding['Assessor'].str.lstrip('A')
respostas_onboarding = nps_onboarding[['Assessor', 'Tamanho da amostra']]
nps_onboarding = nps_onboarding[['Assessor', 'XP - Relacionamento - Onboarding - NPS']]

relatorio = relatorio.merge(nps_onboarding, how='outer', left_on='Código assessor', right_on='Assessor')
del relatorio['Assessor']
relatorio.rename(columns = {'XP - Relacionamento - Onboarding - NPS' : 'NPS Onboarding'}, inplace=True)

# registro de envios

nao_enviados = envios['Survey status'] == 'NOT_SAMPLED'
envios = envios[~nao_enviados].rename(columns = {'Survey status' : 'Número de Envios'})
envios['Código Assessor'] = envios['Código Assessor'].str.lstrip('A')
envios_totais = envios.groupby('Código Assessor').count()['Número de Envios']
envios_totais = envios_totais.reset_index(drop=False)


respostas_totais = respostas_aniversario.merge(respostas_onboarding, how='outer', on='Assessor').fillna(0)
respostas_totais['Tamanho da amostra'] = respostas_totais['Tamanho da amostra_x'] + respostas_totais['Tamanho da amostra_y']

del respostas_totais['Tamanho da amostra_x']
del respostas_totais['Tamanho da amostra_y']

perc_resp = respostas_totais.merge(envios_totais, how='outer', right_on='Código Assessor', left_on='Assessor').fillna(0)
del perc_resp['Assessor']

relatorio = relatorio.merge(perc_resp, how='outer', left_on='Código assessor', right_on='Código Assessor')
del relatorio['Código Assessor']

# atribui as notas do digital aos membros do digital

# digital

mask_digital = relatorio['Código assessor'] == '26839'
aniversario = relatorio.loc[mask_digital, 'NPS Aniversário'].iloc[0]
onboarding = relatorio.loc[mask_digital, 'NPS Onboarding'].iloc[0]
respostas = relatorio.loc[mask_digital, 'Tamanho da amostra'].iloc[0]
envios_registrados = relatorio.loc[mask_digital, 'Número de Envios'].iloc[0]


for membro_digital in ['73981', '73003', 'Sem Responsável Digital']:
    mask_membro = relatorio['Código assessor'] == membro_digital
    relatorio.loc[mask_membro, 'NPS Onboarding'] = onboarding
    relatorio.loc[mask_membro, 'NPS Aniversário'] = aniversario
    relatorio.loc[mask_membro, 'Tamanho da amostra'] = respostas
    relatorio.loc[mask_membro, 'Número de Envios'] = envios_registrados

# exclusive

mask_exclusive = relatorio['Código assessor'] == '26994'
onboarding = relatorio.loc[mask_exclusive, 'NPS Onboarding'].iloc[0]
aniversario = relatorio.loc[mask_exclusive, 'NPS Aniversário'].iloc[0]
respostas = relatorio.loc[mask_exclusive, 'Tamanho da amostra'].iloc[0]
envios_registrados = relatorio.loc[mask_exclusive, 'Número de Envios'].iloc[0]

for membro_exclusive in ['73329', '71326', 'Sem Responsável Exclusive']:
    mask_membro = relatorio['Código assessor'] == membro_exclusive
    relatorio.loc[mask_membro, 'NPS Onboarding'] = onboarding
    relatorio.loc[mask_membro, 'NPS Aniversário'] = aniversario
    relatorio.loc[mask_membro, 'Tamanho da amostra'] = respostas
    relatorio.loc[mask_membro, 'Número de Envios'] = envios_registrados

# private
try:
    mask_private = relatorio['Código assessor'] == '26877'
    onboarding = relatorio.loc[mask_private, 'NPS Onboarding'].iloc[0]
    aniversario = relatorio.loc[mask_private, 'NPS Aniversário'].iloc[0]
    respostas = relatorio.loc[mask_private, 'Tamanho da amostra'].iloc[0]
    envios_registrados = relatorio.loc[mask_private, 'Número de Envios'].iloc[0]
except IndexError: #não consta o private no relatório, pois não foi enviado NPS para ele
    pass

mask_responsavel = relatorio['Código assessor'] == responsavel_private
relatorio.loc[mask_responsavel, 'NPS Onboarding'] = onboarding
relatorio.loc[mask_responsavel, 'NPS Aniversário'] = aniversario
relatorio.loc[mask_responsavel, 'Tamanho da amostra'] = respostas
relatorio.loc[mask_responsavel, 'Número de Envios'] = envios_registrados

relatorio['Percentual de Resposta'] = relatorio['Tamanho da amostra']/relatorio['Número de Envios']

# retira o digital, o exclusive e o private

mask_atendimento = relatorio['Código assessor'] == 'Atendimento Fatorial'

relatorio = relatorio[~ (mask_digital | mask_exclusive | mask_private | mask_atendimento) ]

# coloca o nome do assessor

del relatorio['Nome assessor']
del relatorio['Time']

relatorio = relatorio.merge(assessores, how='left', on='Código assessor')

relatorio['Nome assessor'].fillna(relatorio['Código assessor'], inplace=True)

relatorio = reorder_columns(relatorio, 'Nome assessor', 1)
relatorio = reorder_columns(relatorio, 'Time', 2)

## gera excel

gera_excel(relatorio, caminho_excel, mensagem='\nRelatorio do G20 Registrado')



