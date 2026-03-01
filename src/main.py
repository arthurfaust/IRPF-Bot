import bens_e_direitos, buscar_arquivos
import pywinauto, pyautogui, time
import pandas as pd
from pywinauto.keyboard import send_keys

def lancamento_completo():
    caminho = buscar_arquivos.buscar_arquivo()
    print("Arquivo: {}".format(caminho))
    
    bens_e_direitos.AbrirIRPF()
    bens_e_direitos.AbrirDeclaracaoCliente(caminho)
    bens_e_direitos.AbrirBenseDireitos()
    lancamento_finalizado = bens_e_direitos.NovoLancamento(caminho)
    
    if lancamento_finalizado == True:
        time.sleep(0.5)
        lancamento_completo()
    
lancamento_completo()