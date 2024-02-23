#Ler dados da planilha
#Inserir cada célula de cada linha em um campo do sistema

import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row = 2):
    #nome
    pyautogui.click(1079,142,duration=1.5)
    pyautogui.write(linha[0].value)
    #produto
    pyautogui.click(1075,170,duration=1.5)
    pyautogui.write(linha[1].value)
    #quantidade
    pyautogui.click(1084,195,duration=1.5)
    pyautogui.write(str(linha[2].value))
    #categoria
    pyautogui.click(1142,221,duration=1.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(1027,252,duration=1.5)
    pyautogui.click(558,411,duration=1.5)
    
