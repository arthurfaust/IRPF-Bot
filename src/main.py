# Fluxo principal

import pywinauto, pyautogui, time, bens_e_direitos
import pandas as pd
from pywinauto.keyboard import send_keys

caminho = r"C:\Users\Desktop\Desktop\IRPF-Bot\test\TesteCod-03-01.xlsx"

bens_e_direitos.AbrirIRPF()
bens_e_direitos.AbrirDeclaracaoCliente(caminho)
bens_e_direitos.AbrirBenseDireitos()
bens_e_direitos.NovoLancamento(caminho)