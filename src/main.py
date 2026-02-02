# Fluxo principal

import pywinauto, pyautogui, time, bens_e_direitos
import pandas as pd
from pywinauto.keyboard import send_keys

#df = pd.read_excel(r"C:\Users\Desktop\Desktop\IRPF-Bot\data\InformeTeste.xlsx", header=None)

bens_e_direitos.AbrirIRPF()
bens_e_direitos.AbrirDeclaracaoCliente()
bens_e_direitos.AbrirBenseDireitos()
bens_e_direitos.NovoLancamento()
