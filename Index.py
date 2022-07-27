import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 2
# pyautogui.hotkey("crtl", "t")
# print("Olá Mundo")
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")


pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("enter")


#pyautogui.press("Ieda")