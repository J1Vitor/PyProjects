'''
----------------INFORMAÇÕES SOBRE O PROJETO----------------

Objetivo: Construir um scripto para determinação da precipitações de janeiro de 2023 por bairros da cidade de Campina Grande-PB
Autor: Gabriel Rafá; João Vitor Dias
contato: grmf@academico.ufpb.br ;jvdslcontato@gmail.com
Base dos dados: Centro Nacional de Monitoramento e Alertas de Desastres Naturais-CEMADEN
Data da última atualização: 16/09/2023

'''
# ----------------IMPORTAÇÃO DAS BIBLIOTECAS----------------
import pandas as pd


# ----------------IMPORTAÇÃO DOS DADOS----------------
path = r'Atts_sig\dataframe_cpgrande.csv'
df_brutojan = pd.read_csv(path, delimiter=';')
# print(df_brutojan)


# ----------------TRATAMENTO DOS DADOS----------------
# Determinando as colunas importante para o projeto
df_precipitacoes_jan_camp_grande = pd.DataFrame(
    df_brutojan, columns=['nomeEstacao', 'datahora', 'valorMedida'])

# Conversão dos valores da coluna 'valorMedida' para float (e troca do divisor decimal) além da mudança do tipo da coluna datahora de object para datetime
df_precipitacoes_jan_camp_grande['valorMedida'] = df_precipitacoes_jan_camp_grande['valorMedida'].str.replace(
    ',', '.').astype(float)
df_precipitacoes_jan_camp_grande['datahora'] = pd.to_datetime(
    df_precipitacoes_jan_camp_grande['datahora'])

# Determinando a coluna datahora como index do df
df_precipitacoes_jan_camp_grande.set_index('datahora', inplace=True)

# Somando os valores a cada 12hrs iniciando as 9hrs
df_precipitacoes_jan_camp_grande = df_precipitacoes_jan_camp_grande.groupby(
    'nomeEstacao').resample('24H',  offset='09h00min')['valorMedida'].sum().reset_index()


# Obtendo a precipitação acomulada de janeiro por bairros; função groupby
df_precipitacoes_jan_camp_grande_bairros = df_precipitacoes_jan_camp_grande.groupby(
    'nomeEstacao')['valorMedida'].sum()

# ----------------EXPORTANDO OS RESULTADOS----------------

# Salvando os DFs em csv
# Qualquer atualização deve ser seguida da retirada do comentário da proxima linha. A mesma foi comentada apenas para evitar o erro de permição quando o arquivo está em utilização
df_precipitacoes_jan_camp_grande.to_csv(
    r'C:\Users\Joao\Desktop\Aulas_Sig\Atts_sig\dataframe_cpgrande_alpha.csv', decimal=',')
df_precipitacoes_jan_camp_grande_bairros.to_csv(
    r'C:\Users\Joao\Desktop\Aulas_Sig\Atts_sig\precipitacao_cpgrande_bairros.csv', decimal=',')
