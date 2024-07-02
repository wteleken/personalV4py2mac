import openpyxl

workbook = openpyxl.load_workbook('..\\data\\arquivo_base.xlsx')
workbook.save('novo_treino.xlsx')