#Ler dados da planilha
#Inserir cada célula de cada linha em um campo do sistema
import openpyxl
import pyautogui
workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    pyautogui.click(1177,341,duration=1.5)
    pyautogui.write(linha[0].value)
    pyautogui.click(1174,366,duration=1.5)
    pyautogui.write(linha[1].value)
    pyautogui.click(1197,392,duration=1.5)
    pyautogui.write(str(linha[2].value))
    pyautogui.click(1251,420,duration=1.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(1136,448,duration=1.5)
    pyautogui.click(671,427,duration=1.5)
