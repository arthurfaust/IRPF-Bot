import os, openpyxl
from openpyxl import Workbook

# "Caminho"
directory = r"C:\Users\Desktop\Desktop\IRPF-Bot"

# os.listdir retorna uma list[str] com o nome dos diretórios do "caminho"
files = os.listdir(directory)

# len retorna a quantidade de arquivos no caminho
files_quantity = len(files)

# Lógica para retornar uma mensagem de erro no terminal
if files_quantity == 0:
    raise Exception("No files found in the directory")

wb = Workbook()
ws = wb.active
ws.title = 'Teste PDF'

print(files_quantity)
print(files)
print(ws.title)