# É necessario instalar as lib
#pip install openpyxl pyautogui

import openpyxl

##para chamar a pasta 
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')

workbook['Produtos']