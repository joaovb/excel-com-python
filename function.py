#coding : latin-1

# importando classe Workbook da biblioteca openpyxl
from openpyxl import Workbook

# gerando um objeto Workbook, que contem as informações
# de um arquivo excel, e um objeto Sheet que contém as 
# informações de uma planilha do arquivo excel. 
book = Workbook()
sheet = book.active


# inserir dados na célula, referenciando o objeto
# sheet através de coluna e linha, A1,A2,B1,B2, etc...

sheet['A1'] = 56
sheet['A2'] = 44
sheet['B1'] = 20
sheet['B2'] = 15


# gravando informações na planilha

book.save('sample.xlsx')