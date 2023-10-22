import pyautogui
import pyperclip
import time
import pandas as pd


def copy():
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.1)  # Aguarde um curto período para a cópia ser concluída


def paste(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')


i = 0

nDocuments = int(pyautogui.prompt(text='Digite quantos documentos serão cadastrados', title='Cadastro Automático de documentos', default=''))

pyautogui.PAUSE = 0.5

time.sleep(1.0)

# CLick na pasta
pyautogui.click(x=1850, y=45, clicks=2)

time.sleep(1.0)

while i < nDocuments:
    i = i + 1
        


    # Click no arquivo mais recente
    pyautogui.click(x=400, y=230, clicks=1)
    time.sleep(0.2)

    # Click no renomear Renomear e copiar o nome
    pyautogui.hotkey("f2")
    time.sleep(0.2)
    copy()
    time.sleep(0.2)

    nControl = pyperclip.paste()  # Pega o Numero de controle
    time.sleep(0.2)

    # Click no arquivo mais recente
    pyautogui.click(x=400, y=230, clicks=2)
    time.sleep(1.0)


    # Mudar de aba, colar o nome e pesquisar
    pyautogui.hotkey('ctrl', 'tab')
    time.sleep(1.5)
    paste(nControl)


    time.sleep(1.2)
    pyautogui.press("enter")  # Pesquisa o doc
    time.sleep(5.50)  # Espera carregar
    pyautogui.click(x=1020, y=410, clicks=2)  # Clica no doc
    time.sleep(11.60)  # Espera carregar


    pyautogui.hotkey('alt', 'tab', interval=0.1)  # Volta para o Google
    pyautogui.hotkey('ctrl', 'tab')  # Vai para o PDF
    # Caminha para o tamanho do Zoom
    pyautogui.press('tab', presses=4, interval=0.01)
    pyautogui.write("100")  # Define o zoom como 100%
    time.sleep(0.2)
    pyautogui.click(x=200, y=315, clicks=1)
    time.sleep(0.2)


    # PDF principal e alternar entre as informações (titulo)
    pyautogui.click(x=800, y=335, clicks=3)  # Clica no Titulo 1
    copy()  # Copia o Titulo 1
    title1 = pyperclip.paste()  # Manda o doc para a variável title1

    time.sleep(0.5)

    pyautogui.click(x=800, y=360, clicks=3)  # Clica no Titulo 2

    copy()  # Copia o Titulo 2
    title2 = pyperclip.paste()  # Manda o doc para a variável title2


    time.sleep(0.5)

    pyautogui.click(x=800, y=380, clicks=3)  # Clica no Titulo 3

    copy()  # Copia o Titulo 3
    title3 = pyperclip.paste()  # Manda o doc para a variável title3

    time.sleep(0.5)

    if (title1 == title2) :
        title2 = ""



    cTitle = title1 + title2 + title3  # Concatena title1, title2 e title3




    # Variáveis gerais


    pyautogui.click(x=820, y=410, clicks=3)

    copy()
    tipo = pyperclip.paste()  # Manda a linha do doc para a variável tipo do doc
    time.sleep(0.2)


    pyautogui.click(x=825, y=490, clicks=3)

    copy()
    etapa = pyperclip.paste()  # Manda a linha do doc para a variável etapa


    time.sleep(0.2)


    pyautogui.click(x=985, y=500, clicks=3)

    copy()
    classeSub = pyperclip.paste()  # Manda a linha do doc para a variável classeSub


    time.sleep(0.2)


    pyautogui.click(x=1100, y=500, clicks=3)

    copy()
    sequencial = pyperclip.paste()  # Manda a linha do doc para a variável sequencial

    time.sleep(0.2)


    pyautogui.click(x=1330, y=455, clicks=3)

    copy()
    area = pyperclip.paste()  # Manda a linha do doc para a variável area

    time.sleep(0.2)


    pyautogui.click(x=1515, y=455, clicks=3)

    copy()
    nContrato = pyperclip.paste()  # Manda a linha do doc para a variável nContrato

    time.sleep(0.5)

    if (area == nContrato) :
        nContrato = "0"


    pyautogui.hotkey('alt', 'tab', interval=0.1)  # Volta para o SEsuite


    time.sleep(1.0)
    pyautogui.click(x=750, y=340, clicks=3)  # Clica na aba de Titulo
    paste(cTitle)  # Cola o Titulo
    time.sleep(0.5)


    # Atributos
    pyautogui.click(x=170, y=375, clicks=1)  # Clica em Atributos
    time.sleep(0.5)

    pyautogui.hotkey('alt', 'tab', interval=0.1)  # Volta para o PDF

    time.sleep(1.0)

    # Sistema/Subsistema
    pyautogui.click(x=980, y=415, clicks=3)  # Clica na letra do sistema
    copy()  # Copia o Sistema
    sist = pyperclip.paste()  # Manda o digito do doc para a variável sist

    pyautogui.click(x=1150, y=460, clicks=3)

    copy()
    subSist = pyperclip.paste()  # Manda o digito do doc para a variável subSist

    sistSub = sist + subSist


    pyautogui.hotkey('alt', 'tab', interval=0.1)  # Volta para o Sesuite


    pyautogui.click(x=1050, y=350, clicks=1)

    paste(sistSub)

    time.sleep(0.2)
    pyautogui.press("enter")


    pyautogui.hotkey('alt', 'tab', interval=0.1)  # Volta para o PDF



    # O resto do formulário

    pyautogui.click(x=1110, y=415, clicks=3)

    copy()
    linha = pyperclip.paste()  # Manda a linha do doc para a variável linha



    pyautogui.hotkey('alt', 'tab', interval=0.1)  # Volta para o Sesuite


    pyautogui.click(x=1050, y=445, clicks=1)
    time.sleep(0.2)


    paste(linha)
    time.sleep(0.2)

    pyautogui.press("tab")


    # trecho/Subtrecho

    pyautogui.hotkey('alt', 'tab', interval=0.1)  # Volta para o PDF



    pyautogui.click(x=820, y=460, clicks=3)

    copy()
    trecho = pyperclip.paste()

    time.sleep(0.2)


    pyautogui.click(x=970, y=465, clicks=3)

    copy()
    subTrecho = pyperclip.paste()  # Manda a linha do doc para a variável subtrecho

    trechoSubtrecho = trecho + subTrecho








    pyautogui.hotkey('alt', 'tab', interval=0.1)  # Volta para o Sesuite


    pyautogui.click(x=1050, y=520, clicks=3)

    paste(trechoSubtrecho) #Cola o trecho / subtrecho

    time.sleep(0.2)
    pyautogui.press("enter")


    time.sleep(0.2)


    pyautogui.hotkey('tab')

    time.sleep(0.2)


    paste(etapa) #Cola a etapa

    time.sleep(0.2)
    pyautogui.press("enter")

    time.sleep(0.2)


    pyautogui.hotkey('tab')

    time.sleep(0.2)


    paste(classeSub) #Cola a classe / subclasse 

    time.sleep(0.2)
    pyautogui.press("enter")


    time.sleep(0.2)

    pyautogui.hotkey('tab')

    identificação = tipo + sist + linha + trechoSubtrecho + subSist + etapa + classeSub + sequencial # Concatena tudo para a identificação

    paste(identificação)

    time.sleep(0.2)


    pyautogui.hotkey('tab')
    time.sleep(0.2)

    pyautogui.hotkey('tab')
    time.sleep(0.2)


    paste(area) #Cola a area

    time.sleep(0.2)
    pyautogui.hotkey('tab')

    time.sleep(0.2)


    paste(nContrato) 

    time.sleep(0.2)
    pyautogui.hotkey('tab')

    time.sleep(0.2)
    pyautogui.hotkey('tab')

    pyautogui.write("CPTM- CIA PTA TRENS METROP.")  # COloca a empresa de origem


    time.sleep(0.2)
    pyautogui.press("enter")

    time.sleep(0.5)
    pyautogui.click(x=1880, y=15, clicks=1) #Sai da pasta de Edição

    time.sleep(0.5)
    pyautogui.click(x=585, y=20, clicks=1) # Clica no x do PDF


    time.sleep(0.5)
    pyautogui.click(x=180, y=345, clicks=3) # Clica na barra de pesquisa e apaga o conteudo

    time.sleep(0.2)

    pyautogui.write("")

    time.sleep(0.5)
    pyautogui.click(x=1775, y=15, clicks=1) # Sai do Chrome

    time.sleep(0.5)
    pyautogui.click(x=295, y=225, clicks=1)

    time.sleep(0.5)
    pyautogui.click(x=455, y=140, clicks=1)