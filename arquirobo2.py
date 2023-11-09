# Importar o pandas
import pandas as pd

# Criar uma lista com os nomes das planilhas que você quer ler
planilhas = ["Linha7", "Linha10", "Linha11", "Linha12", "Linha13",
             "XX", "ZZ", "2500R", "9500R", "Linha8", "Linha9", "WW"]

# Ler o arquivo Excel e selecionar as planilhas desejadas, convertendo tudo em texto
dfs = pd.read_excel("C:\\Users\\garot\\Downloads\\Documentos na fila para cadastro.xlsx",
                    sheet_name=planilhas, dtype=str)

# Criar uma lista com os nomes das colunas que você quer imprimir
colunas = ["Nº CONTROLE", "REV.", "AREA", "TIPO", ".SIST.", "LINHA", "PACOTE"]

# Imprimir os dados no console, selecionando apenas as colunas desejadas
for planilha, df in dfs.items():
    print(f"Planilha: {planilha}")
    print(df[colunas])
