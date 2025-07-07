'''
--------------------------------------- PROGRAMA PARA... -------------------------------------------
Adicionar aqui um cabecalho, contendo: objetivo do programa(falar sobre a escolha (valor) da p95 e p99), 
autor e orientador (opcional), dados para contato e 
data da última atualização'''
# ----------------IMPORTAÇÃO DAS BIBLIOTECAS----------------
import pandas as pd

# ----------------PROCESSANDO OS DADOS----------------
def calcula_qtd_p95_e_p99(df):
    '''
    Esta função determina a quantidade de evetos p95 e p99 que ocorreram 
    em um conjunto dados
    
    Par:
        Dataframe 
    '''
    # Renomeação das colunas para aplicação dos filtros
    df.columns = ["Ano", "Mes"] + [f"Dia {i}" for i in range(1, df.shape[1] - 1)]

    # Reorganizando o formado do dataframe para organização dos resultados
    df_reorganizado = pd.melt(df, id_vars=["Ano", "Mes"], var_name="Dia", value_name="precipitacao")

    # Aplicando os filtros para determinar os elementos de p95 e p99
    df_reorganizado["p95"] = df_reorganizado["precipitacao"] >= 42
    df_reorganizado["p99"] = df_reorganizado["precipitacao"] >= 89.17

    # Cria um novo dataframe com as contagens dos eventos de p95 e p99 por dia
    df_resultado = df_reorganizado.groupby(["Ano", "Mes", "Dia"]).agg(
        p95=("p95", "sum"),
        p99=("p99", "sum")
    ).reset_index()

    # Lista de meses e dias para ordenar o resultado gerado
    meses_ordem = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    dias_ordem = [f"Dia {i}" for i in range(1, 32)]

    # Organizando as colunas de mes e as dos dias de acordo com a orgem selecionada
    df_resultado["Mes"] = pd.Categorical(df_resultado["Mes"], categories=meses_ordem, ordered=True)
    df_resultado["Dia"] = pd.Categorical(df_resultado["Dia"], categories=dias_ordem, ordered=True)

    # Pivota o dataframe para ter os dias como colunas
    df_pivot = df_resultado.pivot_table(index=["Ano", "Mes"], columns="Dia", values=["p95", "p99"], fill_value=0).reset_index()

    # Renomeia as colunas para o formato desejado
    df_pivot.columns = ["Ano", "Mes"] + [f"{col[1]}({col[0]})" for col in df_pivot.columns[2:]]

    # Determinando os resultados da contagem dos eventos de p95 e p99
    df_resultado_p95 = df_pivot[["Ano", "Mes"] + [col for col in df_pivot.columns if "p95" in col]].copy()
    df_resultado_p99 = df_pivot[["Ano", "Mes"] + [col for col in df_pivot.columns if "p99" in col]].copy()
    # Adiciona a coluna com o somatório dos eventos
    df_resultado_p95["Somatorio"] = df_resultado_p95.iloc[:, 2:].sum(axis=1)
    df_resultado_p99["Somatorio"] = df_resultado_p99.iloc[:, 2:].sum(axis=1)

    return df_resultado_p95, df_resultado_p99


# ----------------IMPORTAÇÃO DOS DADOS----------------
# colocar aqui o nome do arquivo, pois o script já está na pasta do arquivo
path = r"PERCENTIS_LITORAL_(ATUALIZAÇÃO)_Copia.xlsx"

# ----------------DETERMINAÇÃO DOS DADOS----------------
# Criação do dataframe de acordo com as cidades escolhidas do arquivo enviado
df_precipitacao_bruto_jp = pd.read_excel(path, sheet_name = 'João Pessoa "DFAARA"', usecols="A:AG" ,engine='openpyxl', skiprows=5)
df_precipitacao_bruto_CES = pd.read_excel(path, sheet_name = 'Cruz do Espírito Santo', usecols="A:AG" ,engine='openpyxl', skiprows=5)

# Determinando o resultado da quantidade de eventos p95 e p99 para as cidades escolhidas
df_resultado_p95_jp, df_resultado_p99_jp = calcula_qtd_p95_e_p99(df_precipitacao_bruto_jp) # Resultado João Pessoa-PB
df_resultado_p95_CES, df_resultado_p99_CES = calcula_qtd_p95_e_p99(df_precipitacao_bruto_CES) # Resultado Cruz do Espírito Santo-PB

#----------------EXPORTANDO OS RESULTADOS----------------
"""Qualquer atualização deve ser seguida da retirada do comentário da proxima linha. 
A mesma foi comentada apenas para evitar o erro de permição quando o arquivo está em utilização"""

with pd.ExcelWriter('resultado_final.xlsx', engine='xlsxwriter') as writer:
    # Salvar resultados para João Pessoa (JP)
    df_resultado_p95_jp.to_excel(writer, sheet_name='João Pessoa (p95)', index=False)
    df_resultado_p99_jp.to_excel(writer, sheet_name='João Pessoa (p99)', index=False)

    # Salvar resultados para Cruz do Espírito Santo (CES)
    df_resultado_p95_CES.to_excel(writer, sheet_name='Cruz do Espírito Santo (p95)', index=False)
    df_resultado_p99_CES.to_excel(writer, sheet_name='Cruz do Espírito Santo (p99)', index=False)
