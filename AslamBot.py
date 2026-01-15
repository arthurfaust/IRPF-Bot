# Aslam Bot 1.0 to IRPF 2025

import pywinauto, pyautogui, time
import pandas as pd
from pywinauto.keyboard import send_keys

df = pd.read_excel("C:\\Users\\Desktop\\Desktop\\IRPF-Bot\\Test\\dados.xlsx")

cpf = df.iloc[0,0]

# Functions
def AbrirIRPF():
    # Open Start Menu, search the IRPF program and open it
    time.sleep(0.5)
    pyautogui.press('win')
    time.sleep(0.5)
    send_keys("IRPF 2025")
    time.sleep(0.5)
    pyautogui.press('enter')
    
AbrirIRPF()
# Cor do Trigger para "Em Preenchimento"
cor = (204, 227, 252)

# Logica do trigger correta, mas pixel errado
while True:
    if pyautogui.pixelMatchesColor(215,90,(204, 227, 252)) == True:
        pyautogui.click(x=680,y=340)
        break
    
"""
def AbrirDeclaracaoCliente():
    # Move to "Em Preenchimento" and click it
    # time.sleep(15)
    pyautogui.click(x=680,y=340)

    # Search for the client by the CPF and open 
    time.sleep(0.5)
    pyautogui.click(x=680,y=378)
    time.sleep(0.5)
    send_keys(cpf)
    time.sleep(0.5)
    pyautogui.press('enter')

    # Move to the client found and double click to open it
    time.sleep(0.5)
    pyautogui.doubleClick(x=720,y=435)
    
def PreencherBenseDireitos():
    # Move to the arrow and scroll the menu down
    time.sleep(5)
    pyautogui.moveTo(x=350,y=450)
    pyautogui.click(clicks=16)

    # Move to "Bens e Direitos" and click it
    time.sleep(0.5)
    pyautogui.click(x=85,y=440)

    # Move to "Novo" and click it
    time.sleep(0.5)
    pyautogui.click(x=1105,y=665)
 
AbrirIRPF()
AbrirDeclaracaoCliente()
PreencherBenseDireitos()
"""