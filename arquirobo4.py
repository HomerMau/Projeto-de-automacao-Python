import pyautogui
import pyperclip
import time
import pandas as pd


# Função para colar cada informação de atributo
def atributo(tipo):
    paste(tipo)
    time.sleep(0.1)
    pyautogui.press("enter")

# Função tab genérica


def tab():
    time.sleep(0.11)
    pyautogui.press("tab")
    time.sleep(0.11)

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
        nControle = colunas["Nº CONTROLE"] # Usar o valor da chave "Nº CONTROLE" do dicionário
        nControle1 = nControle[:8] # Obter os 8 primeiros caracteres
        nControle2 = nControle[8:] # Obter o restante da string

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
        if nControle2 !="":
            print(f"FOLHAS = {nControle2}")
        print("\n")



    
    
    # Mudar de aba, colar o nome do n° de controle e pesquisar
    pyautogui.hotkey('alt', 'tab')
    time.sleep(0.2)

    paste(nControle1)
    pyautogui.press("enter")  # Pesquisa o documento
    time.sleep(8.50)  # Espera carregar
    pyautogui.press('tab', presses=29, interval=0.01)  # Clica no documento
    time.sleep(0.1)
    pyautogui.press("enter")  # Clica no documento

    time.sleep(16.00)  # Espera carregar

    pyautogui.hotkey('ctrl', 'a')
    copy()
    nControleComparacao = pyperclip.paste()
    print(nControleComparacao)

    pyautogui.press('tab', presses=7, interval=0.01) #Analiza a Revisão
    pyautogui.hotkey('ctrl', 'a')
    copy()
    revisaoComparacao = pyperclip.paste()
    print(revisaoComparacao)

    pyautogui.press('tab', presses=12, interval=0.01) #Analiza a Revisão
    #NO RESUMO:
    resumo = "ARQUIVOBO 🤖📖 \n\nAREA = {}\nTIPO = {}\nSISTEMA = {}\nLINHA = {}\nPACOTE = {}\nOBSERVAÇÕES GERAIS = {}\nSPSP / CI = {}\nMÍDIA = {}\nRECEBIDO EM = {}\nDESCREVA A DIVERGÊNCIA = {}".format(area, tipo, sistema, linha, pacote, observacoes, spsp, midia, recebidoEm, divergencia)
    paste(resumo)



    #if nControle != nControleComparacao




    time.sleep(0.1)
    pyautogui.alert(
        "ISSO É PARA OS TESTES!!!!!!! Confira se todas as informações dos atributos estão corretas")
