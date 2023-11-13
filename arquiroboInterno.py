import pyautogui
import pyperclip
import time
import pandas as pd


# Função para colar cada informação de atributo
def zeroAEsquerda(linha):
    if linha and len(linha) == 1:
        linha = "0" + linha
        return linha

# Função tab genérica


def tab():
    time.sleep(0.1)
    pyautogui.press("tab")
    time.sleep(0.1)

# Função para copiar e colar


def copy():
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.1)


def paste(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')


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
        # Criar um dicionário com os nomes das colunas como chaves e os valores das células como valores
        colunas = {"Nº CONTROLE": row["Nº CONTROLE"],
                   "REV.": row["REV."],
                   "AREA": row["AREA"],
                   "TIPO": row["TIPO"],
                   "SISTEMA": row["SISTEMA"],
                   "LINHA": row["LINHA"],
                   "PACOTE": row["PACOTE"],
                   "OBSERVAÇÕES GERAIS": row["OBSERVAÇÕES GERAIS"],
                   "SPSP / CI": row["SPSP / CI"],
                   "MÍDIA": row["MÍDIA"],
                   "RECEBIDO EM": row["RECEBIDO EM"],
                   "DESCREVA A DIVERGÊNCIA": row["DESCREVA A DIVERGÊNCIA"]}

        # Separar a string em duas variáveis usando fatiamento
        # Usar o valor da chave "Nº CONTROLE" do dicionário
        nControle = colunas["Nº CONTROLE"]
        nControle1 = nControle[:8]  # Obter os 8 primeiros caracteres
        nControle2 = nControle[8:]  # Obter o restante da string

        # Usar o valor da chave "REV." do dicionário e dos outros
        revisao = colunas.get("REV.", "NaN")
        area = colunas.get("AREA", "NaN")
        tipo = colunas.get("TIPO", "NaN")
        sistema = colunas["SISTEMA"]
        linha = colunas.get("LINHA", "NaN")
        pacote = colunas.get("PACOTE", "NaN")
        observacoes = colunas.get("OBSERVAÇÕES GERAIS", "NaN")
        spsp = colunas.get("SPSP / CI", "NaN")
        midia = colunas.get("MÍDIA", "NaN")
        recebidoEm = colunas.get("RECEBIDO EM", "NaN")
        divergencia = colunas.get("DESCREVA A DIVERGÊNCIA", "NaN")

        # Imprimir o valor de cada variável
        for var, val in colunas.items():
            print(f"{var} = {val}")

        # Imprimir o valor das variáveis nControle1 e nControle2
        print(f"nControle1 = {nControle1}")
        if nControle2 != "":
            print(f"FOLHAS = {nControle2}")
        print("\n")
