import pyautogui as py
import pyperclip
from time import sleep
import pandas
import openpyxl
import numpy

py.PAUSE = 1

# acessar o drive que esta amarzenado a planilha
py.press('win')
py.write('edge')
py.press('enter')
sleep(3)
pyperclip.copy(
    'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
py.hotkey('ctrl', 'v')
py.press('enter')

# acessar a planilha
sleep(5)
py.click(x=579, y=331, clicks=2)
sleep(1)
py.click(x=579, y=331, clicks=1)
sleep(1)
py.click(x=1660, y=195)
sleep(1)
# download
py.click(x=1539, y=697)


sleep(3)
# analisar planilha
tabela = pandas.read_excel(r'C:\Users\lucio\Downloads\Vendas - Dez (3).xlsx')
print(tabela)

# calcular faturamento e quantidade de produtos vendidos
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

# acessar email
sleep(3)
py.hotkey('ctrl', 't')
pyperclip.copy('link do email que fara o envio')
py.hotkey('ctrl', 'v')
py.press('enter')
sleep(5)
py.click(x=143, y=202)
sleep(5)

# enviar email
py.write('email que ira receber o relatorio')
sleep(3)
py.press('tab')
sleep(2)
py.press('tab')
pyperclip.copy('Relatório de Vendas')
py.hotkey('ctrl', 'v')
sleep(3)
py.press('tab')
sleep(3)
texto = f'''
Prezados,Bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidades de produto: {quantidade:,.2f} produtos vendidos

Abs
Guilherme
'''
sleep(2)
pyperclip.copy(texto)
py.hotkey('ctrl', 'v')
sleep(2)
py.hotkey('ctrl', 'enter')