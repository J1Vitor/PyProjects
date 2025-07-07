import pandas as pd
import os
import simplekml

# --- Leitura do arquivo base ---

# Ler o arquivo Excel
df_bruto = pd.read_excel('agricultores_tarifa_verde.xlsx')

# Cria novo df com as colunas necessárias
df_dados_princ = df_bruto.drop(columns=['TIPO PESSOA', 'NUMERO', 'COMPLEMENTO', 'LIVRO',
                                        'LOCALIDADE', 'ROTA', 'CONTA'], errors='ignore')

# Separa LATITUDE e LONGITUDE de COORDENADAS
df_dados_princ[['LATITUDE', 'LONGITUDE']] = df_dados_princ['COORDENADAS'].str.split(
    ',', n=1, expand=True).fillna('')

# Remove a coluna COORDENADAS
df_dados_princ = df_dados_princ.drop(columns=['COORDENADAS'], errors='ignore')

# Converte as coordenadas para formato numérico (muda',' por '.') - apenas no caso do arquivo enviado
df_dados_princ['LATITUDE'] = df_dados_princ['LATITUDE'].str.replace(
    ',', '.', regex=False).astype(float, errors='ignore')
df_dados_princ['LONGITUDE'] = df_dados_princ['LONGITUDE'].str.replace(
    ',', '.', regex=False).astype(float, errors='ignore')

# --- Gera pastas contento os arquivos KMZ's das regionais ---

# Criar pasta principal "regionais_kml"
pasta_principal = 'regionais_kml'
os.makedirs(pasta_principal, exist_ok=True)

# Agrupa os dados por REGIONAL e MUNICÍPIO e gera os KMZ
for regional, df_regional in df_dados_princ.groupby('REGIONAL'):
    # Criar uma pasta para cada regional
    pasta_regional = os.path.join(pasta_principal, regional.replace(' ', '_'))
    os.makedirs(pasta_regional, exist_ok=True)

    for municipio, df_municipio in df_regional.groupby('MUNICÍPIO'):
        # Criar o objeto KML
        kml = simplekml.Kml()

        # Adiciona os pontos ao KML com todas as informações do CSV
        for _, row in df_municipio.iterrows():
            if pd.notnull(row['LATITUDE']) and pd.notnull(row['LONGITUDE']):
                # Criar a descrição com todas as colunas
                descricao = "\n".join([f"{col}: {row[col]}" for col in df_municipio.columns if col not in [
                                      'LATITUDE', 'LONGITUDE']])

                # Adiciona o ponto ao KML
                kml.newpoint(name=row.get('NOME', municipio),
                             coords=[(row['LONGITUDE'], row['LATITUDE'])],
                             description=descricao)

        # Nome do arquivo KMZ com espaços substituídos por underline
        nome_arquivo_kml = os.path.join(
            pasta_regional, f"{municipio.replace(' ', '_')}.kml")
        kml.save(nome_arquivo_kml)
        print(f"Arquivo kml gerado: {nome_arquivo_kml}")

print("\nProcesso concluído!")