"""Cabecalho
Objetivo: obter relatorio quantitativo de atividades realizadas
adicionadas na aplicacao OpenProject
Autor: Gabriel Rafá, Joao Vitor Lima
Ultima modificacao: 26/05/2025
Contato: jvdslcontato@gmail.com; grmf@academico.ufpb.br
"""

# Importacao dos pacotes
import pandas as pd
import glob
import os


def processar_arquivos():
        """
        Le todos os arquivos .xlsx contendo os resultados quantitativos dos extensionistas
        e retorna um relatório quantitativo.
        Par:
            None
        Return:
            Um DataFrame (ou outro tipo, se aplicável) com o relatório quantitativo.
        """

    print(" Procurando arquivos .xlsx na pasta atual...")
    arquivos = glob.glob("*.xlsx")

    if not arquivos:
        print(" Nenhum arquivo .xlsx encontrado na pasta.")
    else:
        print(f" {len(arquivos)} arquivos encontrados: {arquivos}")

    # Inicializar o dicionário para armazenar os dataframes por regional (nome do arquivo)
    regionais = {}

    # Iterar sobre os arquivos encontrados
    for arquivo in arquivos:
        nome_regional = os.path.splitext(os.path.basename(arquivo))[0]
        print(f"\n Processando arquivo: {arquivo}")

        xls = pd.ExcelFile(arquivo)
        print(f" Planilhas encontradas: {xls.sheet_names}")

        dfs = []
        for sheet in xls.sheet_names:
            print(f"Lendo planilha: {sheet}")
            df = pd.read_excel(arquivo, sheet_name=sheet)
            dfs.append(df)

        df_total = pd.concat(dfs, ignore_index=True)
        print(f"Dados concatenados. Linhas totais: {len(df_total)}")

        df_total = df_total[['MUNICÍPIO', 'AUTOR', 'TIPO']]
        print(" Selecionadas colunas: ['MUNICÍPIO', 'AUTOR', 'TIPO']")

        resumo = df_total.groupby(
            ['MUNICÍPIO', 'AUTOR', 'TIPO']).size().reset_index(name='Quantidade')
        print("Dados agrupados por MUNICÍPIO, AUTOR e TIPO")

        tabela_pivot = resumo.pivot_table(
            index=['MUNICÍPIO', 'AUTOR'], columns='TIPO', values='Quantidade', fill_value=0)
        tabela_pivot = tabela_pivot.reset_index()
        print("Tabela pivoteada criada.")

        # Adicionando Total geral e Executado
        tabela_pivot['Total geral'] = tabela_pivot.select_dtypes(
            include='number').sum(axis=1)

        if 'PLANEJAMENTO' in tabela_pivot.columns:
            tabela_pivot['Executado'] = tabela_pivot['Total geral'] - \
                tabela_pivot['PLANEJAMENTO']
            print("Coluna 'Executado' calculada como 'Total geral' - 'PLANEJAMENTO'")

        elif 'PLAN' in tabela_pivot.columns:
            tabela_pivot['Executado'] = tabela_pivot['Total geral'] - \
                tabela_pivot['PLAN']
            print("Coluna 'Executado' calculada como 'Total geral' - 'PLANEJAMENTO'")
        else:
            tabela_pivot['Executado'] = tabela_pivot['Total geral']
            print(
                "Coluna 'PLANEJAMENTO' não encontrada. 'Executado' igual a 'Total geral'")

        linhas = []

        # Organizar por município e inserir linha de total da cidade
        for municipio in tabela_pivot['MUNICÍPIO'].unique():
            print(f" Processando município: {municipio}")
            linhas.append([municipio] + [""] * (tabela_pivot.shape[1] - 1))

            df_cidade = tabela_pivot[tabela_pivot['MUNICÍPIO'] == municipio]

            for _, row in df_cidade.iterrows():
                linhas.append(row.tolist())

            # Calcular totais da cidade
            soma_cidade = df_cidade.select_dtypes(include='number').sum()

            linha_total_cidade = []
            for col in tabela_pivot.columns:
                if col == 'MUNICÍPIO':
                    linha_total_cidade.append(f'Total {municipio}')
                elif col == 'AUTOR':
                    linha_total_cidade.append("")
                elif col in soma_cidade:
                    linha_total_cidade.append(soma_cidade[col])
                else:
                    linha_total_cidade.append("")

            linhas.append(linha_total_cidade)
            print(f" Linha de total para {municipio} adicionada.")

        df_final = pd.DataFrame(linhas, columns=tabela_pivot.columns)
        print(" DataFrame organizado com linhas de cidade e total por cidade.")

        # Calcular somatórios finais das colunas 'Total geral' e 'Executado'
        colunas_soma_final = ['Total geral', 'Executado']
        linha_soma_final = []

        #  Aplicar condição para ignorar linhas onde MUNICÍPIO começa com 'Total '
        filtro_validas = ~df_final['MUNICÍPIO'].astype(
            str).str.startswith('Total ')

        for col in df_final.columns:
            if col == 'MUNICÍPIO':
                linha_soma_final.append('Soma Final')
            elif col in colunas_soma_final:
                soma = pd.to_numeric(
                    df_final.loc[filtro_validas, col], errors='coerce').sum()
                linha_soma_final.append(soma)
            else:
                linha_soma_final.append('')

        df_final.loc[len(df_final)] = linha_soma_final
        print(" Soma final das colunas 'Total geral' e 'Executado' adicionada.")

        regionais[nome_regional] = df_final
        print(f" Dados do arquivo {arquivo} processados com sucesso!")

    # 5. Salvar no Excel (sem formatação condicional)
    try:
        os.makedirs('results', exist_ok=True)
        file_name = r'results\resultados_atingidos_paraiba.xlsx'
        with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
            for nome, df in regionais.items():
                df.to_excel(writer, sheet_name=nome[:31], index=False)

        print(f"\n Arquivo '{file_name}' criado com sucesso!")

    except PermissionError:
        print(
            f" Por favor, feche o arquivo '{file_name}' para sobrescrevê-lo.")
    except Exception as e:
        print(f" Ocorreu um erro ao salvar o arquivo: {e}")


if __name__ == "__main__":
    processar_arquivos()
    print(input('Digite qualquer tecla para contiunar: '))
