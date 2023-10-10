import pyautogui
import pyperclip
import time
import pandas as pd

# pressiona o botão esquerdo do mouse no ponto A (100, 100)
pyautogui.mouseDown (100, 100, button='left')

# move o cursor do mouse para o ponto B (200, 200) em 1 segundo
pyautogui.moveTo (200, 200, 1)

# solta o botão esquerdo do mouse no ponto B
pyautogui.mouseUp (200, 200, button='left')