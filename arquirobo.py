# Importa a biblioteca pandas
import pandas as pd

# Lê o arquivo Excel e armazena em um DataFrame
df = pd.read_excel(
    "C:\\Users\\garot\\Downloads\\Documentos na fila para cadastro.xlsx")

# Seleciona as colunas que você quer pegar
colunas = ["Nº CONTROLE", "REV.", "AREA", "TIPO", ".SIST.", "LINHA", "PACOTE"]

# Imprime os valores das colunas selecionadas
print(df[colunas])
