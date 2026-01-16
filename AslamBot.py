# Aslam Bot 1.0 to IRPF 2025

import pywinauto, pyautogui, time
import pandas as pd
from pywinauto.keyboard import send_keys

df = pd.read_excel("C:\\Users\\Desktop\\Desktop\\IRPF-Bot\\Test\\dados.xlsx", header=None)

cpf = "770.396.770-08"

GrupoAtual = df.iloc[1,0]
CodigoAtual = df.iloc[1,1]
LocalAtual = df.iloc[1,2]
CNPJAtual = df.iloc[1,3]
DiscAtual = df.iloc[1,4]

def AbrirIRPF():
    # Open Start Menu, search the IRPF program and open it
    time.sleep(0.5)
    pyautogui.press('win')
    time.sleep(0.5)
    send_keys("IRPF2025")
    time.sleep(0.5)
    pyautogui.press('enter')
    
def AbrirDeclaracaoCliente():
    # Trigger para abrir a declaração do cliente
    # Cor do pixel de trigger (680,195) = (31, 47, 101)
    while True:
        if pyautogui.pixelMatchesColor(680,195,(31, 47, 101)) == True:
            
            time.sleep(0.5)
            # Move to "Em preenchimento" and click it
            pyautogui.click(x=680,y=340)
            time.sleep(0.5)
            # Move to the search bar and click it
            pyautogui.click(x=680,y=378)
            time.sleep(0.5)
            send_keys(cpf)
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(0.5)
            # Move to the client found and double click to open it
            pyautogui.doubleClick(x=720,y=435)
        
            break
    
def AbrirBenseDireitos():
    # Trigger para iniciar o preenchimento de "Bens e Direitos"
    # Cor do pixel de trigger (180,90) = (160, 225, 185)
    while True:
        if pyautogui.pixelMatchesColor(180,90,(160, 225, 185)) == True:
            
            # Move to the arrow and scroll the menu down
            pyautogui.moveTo(x=350,y=450)
            pyautogui.click(clicks=16)

            # Move to "Bens e Direitos" and click it
            time.sleep(0.5)
            pyautogui.click(x=85,y=440)

            break
        
def NovoLancamento():
    # Move to "Novo" and click it
    time.sleep(0.5)
    pyautogui.click(x=1105,y=665)
    
    # Move to "Grupo" and write
    time.sleep(0.5)
    pyautogui.click(x=630,y=300)
    time.sleep(0.25)
    send_keys(GrupoAtual)
    time.sleep(0.25)
    pyautogui.press('enter')
    
    # Move to "Código" and write
    time.sleep(0.5)
    pyautogui.click(x=630,y=355)
    time.sleep(0.25)
    send_keys(CodigoAtual)
    time.sleep(0.25)
    pyautogui.press('enter')
    
    # Move to "Localização" and write
    time.sleep(0.5)
    pyautogui.click(x=630,y=455)
    time.sleep(0.25)
    
    for i in range(13):
        pyautogui.press('backspace')
        
    send_keys(LocalAtual)
    time.sleep(0.5)
    pyautogui.press('down')
    pyautogui.press('enter')
    
    # Para teste
    print(LocalAtual)
    
AbrirIRPF()
AbrirDeclaracaoCliente()
AbrirBenseDireitos()
NovoLancamento()
