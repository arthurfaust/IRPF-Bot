# Test

import pandas as pd

# Importa a planilha pelo endereço
# A primeira linha é considerada de títulos, por isso usar header=None
df = pd.read_excel("C:\\Users\\Desktop\\Desktop\\IRPF-Bot\\Test\\InformeTeste.xlsx", header=None)

n2 = df.iloc[1,0]
n1 = df.iloc[18,5]

print(n1)
print(n2)