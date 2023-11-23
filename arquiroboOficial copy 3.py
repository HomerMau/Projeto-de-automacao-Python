import pyautogui
import pyperclip
import time
import pandas as pd


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
planilhas = ["Linha11"]

# Ler o arquivo Excel e selecionar as planilhas desejadas, convertendo tudo em texto
dfs = pd.read_excel("C:\\Users\\garot\\Downloads\\Documentos na fila para cadastro (1).xlsx",
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
        # Usar o valor da chave "Nº CONTROLE" do dicionário
        nControle = colunas["Nº CONTROLE"]
        nControle1 = nControle[:8]  # Obter os 8 primeiros caracteres
        nControle2 = nControle[8:]  # Obter o restante da string

        # Usar o valor da chave "TIPO" do dicionário
        tipo = colunas["TIPO"]
        tip0 = tipo[2:]  # Garante que o tipo só vai ser duas letras

        # Usar o valor da chave "REV." do dicionário e dos outros
        revisao = colunas.get("REV.", "0")
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

    # Mudar de aba, colar o nome do n° de controle e pesquisar

        paste(nControle1)
        pyautogui.press("enter")  # Pesquisa o documento
        time.sleep(8.50)  # Espera carregar

        pyautogui.click(x=840, y=420, clicks=2)  # Click de conferencia
        time.sleep(0.30)
        copy()
        testeDeErro = pyperclip.paste()
        print(testeDeErro)

        if testeDeErro == "Não encontramos nenhum resultado para sua pesquisa." or testeDeErro == "Verifique se os filtros e os termos da pesquisa estão corretos." or testeDeErro == "Não encontramos nenhum resultado para sua pesquisa.Verifique se os filtros e os termos da pesquisa estão corretos." or testeDeErro == "Não ":

            time.sleep(0.5)
            pyautogui.click(x=380, y=290, clicks=2, interval=0.05)
            time.sleep(0.5)
            categoriaDoDoc = "DT_{}".format(tip0)
            categoriaDoDoc = categoriaDoDoc[:5]
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
            # Analisa a Revisão
            pyautogui.press('tab', presses=4, interval=0.01)
            pyautogui.hotkey('ctrl', 'a')
            copy()
            revisaoComparacao = pyperclip.paste()
            print(revisaoComparacao)

            # Analiza a Revisão
            pyautogui.press('tab', presses=10, interval=0.01)

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

        pyautogui.click(x=1045, y=650, clicks=2)  # Vai para o resumo

        # Fazer novo arquivo
        if nControle1 != nControleComparacao:
            pyautogui.hotkey('alt', 'f4', interval=0.1)  # Volta para o Sesuite
            time.sleep(0.8)
            pyautogui.click(x=380, y=290, clicks=2)
            time.sleep(1.7)
            categoriaDoDoc = "DT_{}".format(tipo)
            time.sleep(2.7)
            paste(categoriaDoDoc)
            time.sleep(0.2)
            pyautogui.press("enter")
            time.sleep(0.1)
            pyautogui.click(x=550, y=610, clicks=3)
            time.sleep(18.00)  # Espera carregar
            paste(nControle1)
            pyautogui.press('tab', presses=2, interval=0.01)
            pyautogui.write("CADASTRO EM ANDAMENTO")
            # Analisa a Revisão
            pyautogui.press('tab', presses=4, interval=0.01)
            pyautogui.hotkey('ctrl', 'a')
            copy()
            revisaoComparacao = pyperclip.paste()
            print(revisaoComparacao)

            # Analisa a Revisão
            pyautogui.press('tab', presses=10, interval=0.01)

        # NO RESUMO:
        # ARQUIVOBO

        time.sleep(0.1)

        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)

        pyautogui.write("")
        time.sleep(0.1)

        resumo = "\nARQUIROBO \n\nAREA = {}\nREVISÃO = {}\nTIPO = {}\nSISTEMA = {}\nLINHA = {}\nPACOTE = {}\nOBSERVAÇÕES GERAIS = {}\nFOLHAS = {}\nSPSP / CI = {}\nMÍDIA = {}\nRECEBIDO EM = {}\nDESCREVA A DIVERGÊNCIA = {}".format(
            area, revisao, tipo, sistema, linha, pacote, observacoes, nControle2, spsp, midia, recebidoEm, divergencia)
        paste(resumo)

        time.sleep(0.1)

        pyautogui.click(x=1620, y=390, clicks=3)
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'a')
        copy()
        data = pyperclip.paste()
        print(data)
        time.sleep(0.1)

        revisao = pd.Series(revisao).fillna("0").iloc[0]

        pyautogui.click(x=1450, y=390, clicks=3)
        paste(revisao)

        print(f"Valor de data: {data}")

        time.sleep(0.1)

        pyautogui.click(x=780, y=335, clicks=2)  # Analisa o titulo
        time.sleep(0.1)

        pyautogui.hotkey('ctrl', 'a')
        copy()
        titulo = pyperclip.paste()
        print(titulo)

        time.sleep(0.1)

        if data == "" or data == resumo or revisao != "0" or titulo == "" or titulo == "CADASTRO EM ANDAMENTO":
                # Vai para guia atributos
                time.sleep(0.1)
                # Clica na aba Atributos
                pyautogui.click(x=95, y=370, clicks=2)
                time.sleep(0.1)

                pyautogui.press('tab', presses=2, interval=0.01)

                sist = "{}0000".format(sistema)
                paste(sist)
                pyautogui.press("enter")

                tab()
                # Função para colar cada informação de atributo

                if linha and len(linha) == 1:
                    linha = "0" + linha

                if linha == "07":
                    linha = "07 - Rubi - Luz"

                if linha == "11":
                    linha = "11 - Coral - Barra Funda"

                if linha == "12":
                    linha = "12 - Safira - Br"

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
                time.sleep(0.3)
                pyautogui.click(x=40, y=175, clicks=3)
                time.sleep(0.9)
                pyautogui.press("enter")
                time.sleep(18.0)

        time.sleep(0.1)
        pyautogui.click(x=140, y=160, clicks=3)
        time.sleep(10.0)
        pyautogui.hotkey('f5')
        time.sleep(9.0)
