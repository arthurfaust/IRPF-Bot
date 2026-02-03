# Fluxo principal

import pywinauto, pyautogui, time, bens_e_direitos
import pandas as pd
from pywinauto.keyboard import send_keys

bens_e_direitos.AbrirIRPF()
bens_e_direitos.AbrirDeclaracaoCliente()
bens_e_direitos.AbrirBenseDireitos()
bens_e_direitos.NovoLancamento()