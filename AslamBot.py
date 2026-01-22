# Aslam Bot 1.0 to IRPF 2025

import pywinauto, pyautogui, time
import pandas as pd
from pywinauto.keyboard import send_keys

df = pd.read_excel(r"C:\Users\Desktop\Desktop\IRPF-Bot\Test\InformeTeste.xlsx", header=None)

def AbrirIRPF():
    # Open Start Menu, search the IRPF program and open it
    time.sleep(0.5)
    pyautogui.press('win')
    time.sleep(0.5)
    send_keys("IRPF2025")
    time.sleep(0.5)
    pyautogui.press('enter')
    
def AbrirDeclaracaoCliente():
    
    nome = df.iloc[1,0]
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
            send_keys(nome, with_spaces=True)
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
    linhaAtual = 18
    print(len(df))
    
    # Enquanto a linha atual for menor que o tamanho do dataframe
    while linhaAtual < len(df):

        # Se a célula atual estiver vazia, para a automação
        if pd.isna(df.iloc[linhaAtual, 0]):
            break
            
        GrupoAtual = df.iloc[linhaAtual,0]
        CodigoAtual = df.iloc[linhaAtual,1]
        LocalAtual = df.iloc[linhaAtual,2]
        DiscAtual = df.iloc[linhaAtual,4]
        Sit24 = df.iloc[linhaAtual,6]

        # Move o cursor para "Novo" e clica
        time.sleep(0.5)
        pyautogui.click(x=1105,y=665)

        # Move o cursor para "Grupo" e digita o grupo atual
        time.sleep(0.5)
        pyautogui.click(x=630,y=300)
        time.sleep(0.25)
        send_keys(GrupoAtual)
        time.sleep(0.25)
        pyautogui.press('enter')

        # Move o cursor para "Código" e digita o código atual
        time.sleep(0.5)
        pyautogui.click(x=630,y=355)
        time.sleep(0.25)
        send_keys(CodigoAtual)
        time.sleep(0.25)
        pyautogui.press('enter')

        # Move o cursor para "Localização" e digita a loc atual
        time.sleep(0.5)
        pyautogui.click(x=630,y=455)
        time.sleep(0.25)

        for i in range(13):
            pyautogui.press('backspace')

        send_keys(LocalAtual, with_spaces=True)
        time.sleep(0.5)
        pyautogui.press('down')
        pyautogui.press('enter')
        
        # Move o cursor para "Discriminação" e digita a disc atual
        time.sleep(0.5)
        pyautogui.click(x=630,y=650)
        time.sleep(0.25)
        send_keys(DiscAtual, with_spaces=True)
        time.sleep(0.25)
        pyautogui.press('enter')

        # Scroll down
        time.sleep(0.25)
        pyautogui.moveTo(1348,380)
        time.sleep(0.25)
        pyautogui.dragTo(1348,520,1,button='left')

        # Move o cursor para "Situação em 31/12/23"
        time.sleep(0.25)
        pyautogui.click(x=480,y=630)
        time.sleep(0.25)
        for i in range(4):
            pyautogui.press('del')
        time.sleep(0.25)
        """
        send_keys(Sit23)
        time.sleep(0.25)
        pyautogui.press('enter')
        """
        
        # Move o cursor para "Situação em 31/12/24"
        time.sleep(0.25)
        pyautogui.click(x=665,y=630)
        time.sleep(0.25)
        for i in range(4):
            pyautogui.press('del')
        time.sleep(0.25)
        send_keys(Sit24)
        time.sleep(0.25)
        pyautogui.press('enter')

        # Move o cursor para "Ok" e clica
        time.sleep(0.25)
        pyautogui.click(x=1100,y=710)

        linhaAtual += 1
        print(linhaAtual)
        

AbrirIRPF()
AbrirDeclaracaoCliente()
AbrirBenseDireitos()
NovoLancamento()
