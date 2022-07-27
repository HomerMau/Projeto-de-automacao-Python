import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 3




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
time.sleep(2)

pyautogui.click(x=375, y=375) #Click no arquivo
pyautogui.click(x=1150, y=190) #Click nos 3 pontos
pyautogui.click(x=980, y=600) #Click no download
time.sleep(7)

#Calcular os Indicadores (faturamento, quantidade de produtos)

tabela = pd.read_excel(r"C:\Users\Cristiano\Downloads\Vendas - Dez.xlsx")
print(tabela)


quantidade = tabela["Quantidade"].sum()  #soma da coluna quantidade
faturamento = tabela["Valor Final"].sum()   #soma da coluna valor final

print(quantidade)
print(faturamento)

#Entrar no E-mail


pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")  #Abrir nova Aba

time.sleep(2)
pyautogui.click(x=500, y=420)
time.sleep(2)
pyautogui.click(x=350, y=50)
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("enter")
time.sleep(8)



#Mandar o E-mail

pyautogui.click(x=50, y=200)
time.sleep(10)#Clicar no botao +

pyautogui.write("garotadosobrado@gmail.com")#Prencher o Destinatario
pyautogui.press("tab") # Selecionar o E-Mail
pyautogui.press("tab") # Passar para o campo de Assunto

pyperclip.copy("Oiii Mamis te amo 😘😘")#Prencher o Assunto
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("tab") # Passar para o campo de Corpo


texto = f"""Oiii mae!
O numero de vezes que eu te amo é maior que esse número {quantidade + faturamento} !!!""" #Prencher o Corpo do E-mail

pyperclip.copy(texto) #Prencher o Corpo do E-mail
pyautogui.hotkey('ctrl', 'v')

pyautogui.hotkey('crtl', 'enter')#Enviar E-mail

#pyautogui.press("Ieda")chrome