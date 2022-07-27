import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 2
# pyautogui.hotkey("crtl", "t")
# print("Olá Mundo")

#Para entrar no sistema
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")

time.sleep(2)
pyautogui.click(x=500, y=420)
time.sleep(2)
pyautogui.click(x=350, y=50)
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("enter")

#Site Carregando
time.sleep(8)
pyautogui.click(x=365, y=290, clicks=2)


pyautogui.click(x=375, y=375) #Click no arquivo
pyautogui.click(x=1150, y=190) #Click nos 3 pontos
pyautogui.click(x=980, y=600) #Click no download
time.sleep(7)

#Calcular os Indicadores (faturamento, quantidade de produtos)

 




#pyautogui.press("Ieda")chrome