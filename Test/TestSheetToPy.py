# Test

import pandas as pd

# Importa a planilha pelo endereço
# A primeira linha é considerada de títulos, por isso usar header=None
df = pd.read_excel("C:\\Users\\Desktop\\Desktop\\IRPF-Bot\\Test\\dados.xlsx", header=None)

n1 = df.iloc[0,0]

print(n1)