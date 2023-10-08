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
time.sleep(0.5)
paste()
time.sleep(0.2)
pyautogui.press("enter")
time.sleep(4.00)
pyautogui.click(x=1020, y=410, clicks=2)
time.sleep(4.60)


pyautogui.hotkey('alt', 'tab', interval=0.1)
pyautogui.hotkey('ctrl','tab')
pyautogui.press('tab', presses=4, interval=0.01)
pyautogui.write("100")
time.sleep(0.2)
pyautogui.click(x=200, y=315, clicks=1)
