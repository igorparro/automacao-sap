# ler dados da planilha 
# inserir cada celular de cada linha em um campo do sistema

import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheets = workbook['vendas']

for linha in vendas_sheets.iter_rows(min_row=2):
    #Murilo Barros	Cadeira	434	Esportes
    #cliente
    pyautogui.click(971,513, duration=0.5)
    pyautogui.write(linha[0].value)
    #produto
    pyautogui.click(973,542, duration=0.5)
    pyautogui.write(linha[1].value)
    #quantidade
    pyautogui.click(940,567, duration=0.5)
    pyautogui.write(str(linha[2].value))
    #cateogria
    pyautogui.click(1031,592, duration=0.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(868,621, duration=0.5)
    pyautogui.click(945,583, duration=0.5)



