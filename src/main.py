import bens_e_direitos, interface
import pywinauto, pyautogui, time
import pandas as pd
from pywinauto.keyboard import send_keys

def lancamento_completo():
    while True:
        caminho = interface.buscar_arquivo()
        print("Arquivo: {}".format(caminho))
        
        if not caminho: break
    
        bens_e_direitos.abrir_irpf()
        bens_e_direitos.abrir_declaracao(caminho)
        bens_e_direitos.abrir_bens_e_direitos()
        lancamento_finalizado = bens_e_direitos.novo_lancamento(caminho)
    
        if lancamento_finalizado != True: break
    
lancamento_completo()