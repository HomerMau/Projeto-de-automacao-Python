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



pyautogui.hotkey('alt', 'tab')
time.sleep(0.2)

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

    paste(nControle1)
    pyautogui.press("enter")  # Pesquisa o documento
    time.sleep(8.50)  # Espera carregar

    pyautogui.click(x=450, y=290, clicks=2)  # Clica no documento


    time.sleep(0.1)
    pyautogui.press("enter")  # Clica no documento

    time.sleep(20.00)  # Espera carregar

    pyautogui.hotkey('ctrl', 'a')
    copy()
    nControleComparacao = pyperclip.paste()
    print(nControleComparacao)

    pyautogui.click(x=735, y=340, clicks=2)  # Analisa o titulo
    pyautogui.hotkey('ctrl', 'a')
    copy()
    titulo = pyperclip.paste()
    print(titulo)

    pyautogui.click(x=1425, y=395, clicks=2)  # Analiza a Revisão




    pyautogui.hotkey('ctrl', 'a')
    copy()
    revisaoComparacao = pyperclip.paste()
    print(revisaoComparacao)

    pyautogui.press('tab', presses=12, interval=0.01) #Vai para o resumo

    
    #Fazer novo arquivo
    if nControle1 != nControleComparacao:
        pyautogui.hotkey('alt', 'f4', interval=0.1)  # Volta para o Sesuite
        time.sleep(0.5)
        pyautogui.click(x=380, y=290, clicks=2)
        time.sleep(0.5)
        categoriaDoDoc = "DT_{}".format(tipo)
        time.sleep(2.0)
        paste(categoriaDoDoc)
        time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.1)
        pyautogui.click(x=550, y=610, clicks=3)
        time.sleep(16.00)  # Espera carregar
        paste(nControle1)
        pyautogui.press('tab', presses=2, interval=0.01)
        pyautogui.write("CADASTRO EM ANDAMENTO")
        pyautogui.press('tab', presses=4, interval=0.01) #Analisa a Revisão
        pyautogui.hotkey('ctrl', 'a')
        copy()
        revisaoComparacao = pyperclip.paste()
        print(revisaoComparacao)

        pyautogui.press('tab', presses=10, interval=0.01) #Analiza a Revisão









    #NO RESUMO:
    # ARQUIVOBO 🤖 

    resumo = "ARQUIVOBO \n\nAREA = {}\nTIPO = {}\nSISTEMA = {}\nLINHA = {}\nPACOTE = {}\nOBSERVAÇÕES GERAIS = {}\nSPSP / CI = {}\nMÍDIA = {}\nRECEBIDO EM = {}\nDESCREVA A DIVERGÊNCIA = {}".format(area, tipo, sistema, linha, pacote, observacoes, spsp, midia, recebidoEm, divergencia)
    paste(resumo)


    if titulo == "" or titulo == "CADASTRO EM ANDAMENTO":

        #Vai para guia atributos
        time.sleep(0.1)
        pyautogui.click(x=95, y=370, clicks=2)  # Clica na aba Atributos
        time.sleep(0.1)

        pyautogui.press('tab', presses=2, interval=0.01)

        sist = "{}0000".format(sistema)
        paste(sist)
        pyautogui.press("enter")

        tab()
        linha = zeroAEsquerda(linha)
        lin = "{}".format(linha)
        paste(linha)
        pyautogui.press("enter")

        tab()
        trechoSubtrecho = "0000"
        paste(trechoSubtrecho)

        tab()
        etapa = "0"
        paste(etapa)
        pyautogui.press("enter")

        tab()
        classe = ""
        paste(classe)
        pyautogui.press("enter")

        tab()
        classificacao = "{}{}{}".format(tipo, sistema, lin)
        paste(classificacao)

        tab()
        naoIdentificado = "Não Identificado"
        paste(naoIdentificado)
        pyautogui.press("enter")

        tab()
        projetista = "CPTM - Cia Pta Trens"
        paste(projetista)
        pyautogui.press("enter")

        tab()
        nContrato = "0"
        paste(nContrato)
        pyautogui.press("enter")


    time.sleep(0.1)
   # pyautogui.click(x=130, y=160, clicks=2)
    pyautogui.hotkey('alt', 'f4', interval=0.1)
    time.sleep(1.0)
    pyautogui.hotkey('f5')
    time.sleep(4.0)





