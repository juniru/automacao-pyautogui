import time

import pandas as pd
import pyautogui
import pyperclip

pyautogui.PAUSE = 0.5
pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
pyperclip.copy(
    'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=319, y=389, clicks=2)
time.sleep(2)
pyautogui.click(x=389, y=400)
pyautogui.click(x=1712, y=194)
pyautogui.click(x=1475, y=564)
time.sleep(5)

tabela = pd.read_excel(r'C:\Users\JUNIR\Downloads\Vendas - Dez.xlsx')
print(tabela)

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=35, y=200)
time.sleep(1)
pyautogui.write('juniro@live.com')
time.sleep(1)
pyautogui.press('tab')
pyautogui.press('tab')
pyperclip.copy('Relatório de vendas')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
texto = f'''
Prezados, bom dia

O faturamento de ontem foi de: R$ {faturamento:,.2f}.
A quantidade de produtos foi de: {quantidade:,}.

Att
Wellington.
'''

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
time.sleep(5)
pyautogui.hotkey('ctrl', 'enter')
