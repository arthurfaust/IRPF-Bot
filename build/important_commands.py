# Comandos importantes 
import os, openpyxl
from openpyxl import Workbook

"""
Estabelece o "Caminho" com r" para que o python não interprete
caracteres de escape, como \n, \t, \u
"""
directory = r"C:\Users\Desktop\Desktop\IRPF-Bot"

# os.listdir retorna uma list[str] com o nome dos diretórios do "Caminho"
files = os.listdir(directory)

# len retorna a quantidade de arquivos no caminho
files_quantity = len(files)

# Lógica para retornar uma mensagem de erro no terminal caso o dir estiver vazio
if files_quantity == 0:
    raise Exception("No files found in the directory")

wb = Workbook()
ws = wb.active
ws.title = 'Teste PDF'

print(files_quantity)
print(files)
print(ws.title)