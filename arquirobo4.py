# Importar o pandas
import pandas as pd

# Criar uma lista com os nomes das planilhas que você quer ler
planilhas = ["Linha7", "Linha10", "Linha11", "Linha12", "Linha13",
             "XX", "ZZ", "2500R", "9500R", "Linha8", "WW"]

# Ler o arquivo Excel e selecionar as planilhas desejadas, convertendo tudo em texto
dfs = pd.read_excel("C:\\Users\\garot\\Downloads\\Documentos na fila para cadastro.xlsx",
                    sheet_name=planilhas, dtype=str)

# Imprimir o texto de cada célula no console, selecionando apenas as colunas desejadas
for planilha, df in dfs.items():
    print(f"Planilha: {planilha}")
    for i, row in df.iterrows():
        # Criar um dicionário com os nomes das colunas como chaves e os valores das células como valores
        colunas = {"Nº CONTROLE": row["Nº CONTROLE"],
                   "REV.": row["REV."],
                   "AREA": row["AREA"],
                   "TIPO": row["TIPO"],
                   ".SIST.": row[".SIST."],
                   "LINHA": row["LINHA"],
                   "PACOTE": row["PACOTE"],
                   "OBSERVAÇÕES GERAIS": row["OBSERVAÇÕES GERAIS"],
                   "SPSP / CI": row["SPSP / CI"],
                   "MÍDIA": row["MÍDIA"],
                   "RECEBIDO EM": row["RECEBIDO EM"],
                   "DESCREVA A DIVERGÊNCIA": row["DESCREVA A DIVERGÊNCIA"]}
        # Separar a string em duas variáveis usando fatiamento
        nControle = colunas["Nº CONTROLE"] # Usar o valor da chave "Nº CONTROLE" do dicionário

        nControle1 = nControle[:8] # Obter os 8 primeiros caracteres
        nControle2 = nControle[8:] # Obter o restante da string

        # Imprimir o valor de cada variável
        for var, val in colunas.items():
            print(f"{var} = {val}")
        # Imprimir o valor das variáveis nControle1 e nControle2
        print(f"nControle1 = {nControle1}")
        if nControle2 !="":
            print(f"nControle2 = {nControle2}")
        print("\n")
