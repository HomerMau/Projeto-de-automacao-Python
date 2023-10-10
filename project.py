import pyautogui
import pyperclip
import time
import pandas as pd


def copy():
    pyautogui.hotkey('ctrl', 'c')


def paste():

    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('v')
    pyautogui.keyUp('ctrl')


pyautogui.alert(
    "O código vai começar. Não utilize nada do computador até o código finalizar!")

pyautogui.PAUSE = 0.5


# CLick na pasta
pyautogui.click(x=1850, y=45, clicks=2)

time.sleep(0.5)

# Click no arquivo mais recente
pyautogui.click(x=400, y=230, clicks=1)
time.sleep(0.2)

# Click no renomear Renomear e copiar o nome
pyautogui.hotkey("f2")
time.sleep(0.5)
copy()



# Click no arquivo mais recente
pyautogui.click(x=400, y=230, clicks=2)
time.sleep(1.0)


# Mudar de aba, colar o nome e pesquisar
pyautogui.hotkey('ctrl','tab')
time.sleep(1.5)
paste()
time.sleep(1.2)
pyautogui.press("enter")
time.sleep(4.00)
pyautogui.click(x=1020, y=410, clicks=2)
time.sleep(8.60)


pyautogui.hotkey('alt', 'tab', interval=0.1)
pyautogui.hotkey('ctrl','tab')
pyautogui.press('tab', presses=4, interval=0.01)
pyautogui.write("100")
time.sleep(0.2)
pyautogui.click(x=200, y=315, clicks=1)



#PDF principal e alternar entre as informações
pyautogui.click(x=590, y=400, clicks=3)
copy()
time.sleep(1.0)
pyautogui.hotkey('alt', 'tab', interval=0.1)
pyautogui.click(x=750, y=340, clicks=3)
paste()
time.sleep(1.0)

#PDF principal e alternar entre as informações PARTE 2

pyautogui.hotkey('alt', 'tab', interval=0.1)


pyautogui.click(x=590, y=420, clicks=3)
copy()
time.sleep(1.0)
pyautogui.hotkey('alt', 'tab', interval=0.1)
pyautogui.click(x=750, y=340, clicks=3)
pyautogui.press('right')
paste()
time.sleep(1.0)

#Atributos
pyautogui.click(x=170, y=375, clicks=1)
pyautogui.hotkey('alt', 'tab', interval=0.1)
time.sleep(1.0)

#Sistema/Subsitema
pyautogui.click(x=685, y=475, clicks=2)
copy()
pyautogui.hotkey('alt', 'tab', interval=0.1)
pyautogui.click(x=685, y=525, clicks=3)
paste()

time.sleep(1.0)
pyautogui.hotkey('alt', 'tab', interval=0.1)
pyautogui.click(x=850, y=520, clicks=3)
copy()
pyautogui.hotkey('alt', 'tab', interval=0.1)
time.sleep(1.0)

pyautogui.click(x=685, y=525, clicks=1)
pyautogui.press('right')
time.sleep(1.0)
paste()
pyautogui.click(x=685, y=525, clicks=3)
copy()
pyautogui.click(x=600, y=350, clicks=1)
paste()
time.sleep(1.0)
pyautogui.press("enter")



