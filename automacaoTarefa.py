#Automação de tarefas em python
#---Executado no Jupyter---#

import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

#Entrando no sistema
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(5)

#Navegar e encontrar o arquivo
pyautogui.click(x = 366, y = 298, clicks = 2)
time.sleep(3)

#Selecionar arquivo e fazer download
pyautogui.click(x = 454, y = 487, clicks = 1)
pyautogui.click(x = 1720, y = 186, clicks = 1)
pyautogui.click(x = 1553, y = 584, clicks = 1)
time.sleep(5)

#Importando arquivo pro python
tabela = pd.read_excel(r'C:\Users\ruyfl\Downloads\Vendas - Dez.xlsx')
display(tabela)

#Calculando os indicadores(Faturamento e Quantidade)
faturamento = tabela['Valor Final'].sum()
print(faturamento)
quantidade = tabela['Quantidade'].sum()
print(quantidade)

#Enviando email com os indicadores de venda
pyautogui.hotkey('ctrl', 't')
pyperclip.copy(r'https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(4)

#Preencher o email
pyautogui.click(x = 79, y = 191, clicks = 1)
pyautogui.write('ruyflamengo12@outlook.com')
pyautogui.press('tab')
pyautogui.press('tab')
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

#Preenchendo o texto e enviando o email
texto = f"""
Prezados,

Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
Att.,
Rui Duarte
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
time.sleep(4)
pyautogui.hotkey("ctrl", "enter")