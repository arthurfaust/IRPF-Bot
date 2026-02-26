import pywinauto, pyautogui, time
import pandas as pd
from pywinauto.keyboard import send_keys

def AbrirIRPF():
    # Open Start Menu, search the IRPF program and open it
    time.sleep(0.5)
    pyautogui.press('win')
    time.sleep(0.5)
    send_keys("IRPF2025")
    time.sleep(0.5)
    pyautogui.press('enter')
    
def AbrirDeclaracaoCliente(caminho):
    
    df = pd.read_excel(caminho, header=None)
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
        
def NovoLancamento(caminho):
    
    df = pd.read_excel(caminho, header=None)
    linhaAtual = 18
    print("Tamanho do dataframe: {}".format(len(df)))
    
    # Enquanto a linha atual for menor que o tamanho do dataframe
    while linhaAtual < len(df):

        # Se a célula ATUAL estiver vazia, para a automação
        if pd.isna(df.iloc[linhaAtual, 0]):
            break
            
        GrupoAtual = df.iloc[linhaAtual,0]
        CodigoAtual = df.iloc[linhaAtual,1]
        LocalAtual = df.iloc[linhaAtual,2]
        DiscAtual = df.iloc[linhaAtual,4]
        Sit24 = df.iloc[linhaAtual,6]
        LucroOuPrejuizo = "100,00"
        ImpostoExt = "200,00"
        ValorRecebido = "300,00"
        ImpostoPagoExt = "400,00"
        CodAltcoin = "ETH"
        
        # Move o cursor para "Novo" e clica
        time.sleep(0.5)
        pyautogui.click(x=1105,y=665)
        
        # Fazer um trigger aqui para a transição de tela
        
        # Preenche o grupo atual
        time.sleep(0.25)
        send_keys(GrupoAtual)
        time.sleep(0.25)
        pyautogui.press('enter')
        
        # Preenche o código atual
        time.sleep(0.25)
        send_keys(CodigoAtual)
        time.sleep(0.25)
        pyautogui.press('enter')
        
        # Move o cursor para "Localização" e digita a loc atual
        time.sleep(0.25)
        send_keys('^A')
        pyautogui.press('backspace')
        time.sleep(0.25)
        send_keys(LocalAtual, with_spaces=True)
        time.sleep(0.25)
        pyautogui.press('enter')
        
        if GrupoAtual == "08" and CodigoAtual == "01":
            if LocalAtual != "105 - Brasil":
                # Dá OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá enter no autocustodiante
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche a discriminação
                time.sleep(0.25)
                send_keys(DiscAtual, with_spaces= True)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                """
                # Scroll down
                time.sleep(0.25)
                pyautogui.moveTo(1348,380)
                time.sleep(0.25)
                pyautogui.dragTo(1348,600,0.5,button='left')
                """
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(0.25)    
                pyautogui.press('backspace')
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Situação em 31/12/24" e vai para "Repetir"
                time.sleep(0.25)    
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(Sit24)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá tab em "Repetir"
                time.sleep(0.25)
                pyautogui.press('tab')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Lucro ou Prejuízo"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(LucroOuPrejuizo)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(LucroOuPrejuizo)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Valor Recebido"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(ValorRecebido)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(ImpostoPagoExt)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá OK e finaliza o lançamento
                time.sleep(0.25)
                pyautogui.press('enter')
                
        linhaAtual += 1
        print("Linha Atual: {}".format(linhaAtual))
                
        """
        if GrupoAtual == "08" and CodigoAtual == "01":
        
            # Vai para "É autocustodiante?"
            time.sleep(0.25)
            pyautogui.press('enter')
        
            # Vai para "CNPJ custodiante"
            time.sleep(0.25)
            pyautogui.press('enter')
        
            # Preenche a discriminação
        
        
            pyautogui.press('enter')
    
        if GrupoAtual == "08" and CodigoAtual == "01":
    
            # Move o cursor para "Discriminação" e digita a disc atual
            time.sleep(0.5)
            pyautogui.click(x=630,y=650)
            time.sleep(0.25)
            send_keys(DiscAtual, with_spaces=True)
            time.sleep(0.25)
            pyautogui.press('enter')

            if LocalAtual != "105 - Brasil":
                # Scroll down
                time.sleep(0.25)
                pyautogui.moveTo(1348,380)
                time.sleep(0.25)
                pyautogui.dragTo(1348,600,0.5,button='left')
            
                # Move o cursor para "Situação em 31/12/23"
                time.sleep(0.25)
                pyautogui.click(x=480,y=510)
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
        
                # Move o cursor para "Situação em 31/12/24"
                time.sleep(0.25)
                pyautogui.click(x=665,y=510)
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(Sit24)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Move o cursor para "Lucro ou Prejuízo" em "Aplicação Financeira"
                time.sleep(0.25)
                pyautogui.click(x=490,y=630)
                time.sleep(0.25)
                pyautogui.press('enter')
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(LouP)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Move o cursor para "Imposto pago no Exterior" em "Aplicação Financeira"
                time.sleep(0.25)
                pyautogui.click(x=665,y=630)
                time.sleep(0.25)
                pyautogui.press('enter')
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(ImpExt)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Move o cursor para "Valor Recebido" em "Lucros e Dividendos"
                time.sleep(0.25)
                pyautogui.click(x=845,y=630)
                time.sleep(0.25)
                pyautogui.press('enter')
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(ValorRecebido)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Move o cursor para "Imposto Pago Exterior" em "Lucros e Dividendos"
                time.sleep(0.25)
                pyautogui.click(x=1015,y=630)
                time.sleep(0.25)
                pyautogui.press('enter')
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(ImpPagoExt)
                time.sleep(0.25)
                pyautogui.press('enter')
                
            else:
                # Scroll down
                time.sleep(0.25)
                pyautogui.moveTo(1348,380)
                time.sleep(0.25)
                pyautogui.dragTo(1348,520,1,button='left')
            
                # Move o cursor para "Situação em 31/12/23"
                time.sleep(0.25)
                pyautogui.click(x=480,y=630)
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
            
                # Move o cursor para "Situação em 31/12/24"
                time.sleep(0.25)
                pyautogui.click(x=665,y=630)
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(Sit24)
                time.sleep(0.25)
                pyautogui.press('enter')
            
            # Move o cursor para "Ok" e clica
            time.sleep(0.25)
            pyautogui.click(x=1100,y=710)
            
        if GrupoAtual == "08" and CodigoAtual == "02":
            
            # Move o cursor para "Código Altcoin" e digita o cod atual
            time.sleep(0.25)
            pyautogui.click(x=630,y=520)
            time.sleep(0.25)
            send_keys(CodAltcoin, with_spaces=True)
            time.sleep(0.25)
            pyautogui.press('down')
            pyautogui.press('enter')
            
            # Scroll down
            time.sleep(0.25)
            pyautogui.moveTo(1348,380)
            time.sleep(0.25)
            pyautogui.dragTo(1348,600,0.5,button='left')
            
            if LocalAtual != "105 - Brasil":
                print("Operação Estrangeira")
                
                # Move o cursor para "Discriminação" e digita a disc atual
                time.sleep(0.5)
                pyautogui.click(x=725,y=375)
                time.sleep(0.25)
                send_keys(DiscAtual, with_spaces=True)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                
                
                
            else:
                print("Operação Nacional")
                
                # Move o cursor para "Discriminação" e digita a disc atual
                time.sleep(0.5)
                pyautogui.click(x=630,y=500)
                time.sleep(0.25)
                send_keys(DiscAtual, with_spaces=True)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Move o cursor para "Situação em 31/12/23"
                time.sleep(0.25)
                pyautogui.click(x=490,y=630)
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                
                # Move o cursor para "Situação em 31/12/24"
                time.sleep(0.25)
                pyautogui.click(x=665,y=630)
                time.sleep(0.25)
                send_keys('^A')
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(Sit24, with_spaces=True)
        
        # Move o cursor para "OK" e clica        
        time.sleep(0.25)
        pyautogui.click(x=1100,y=710)

        linhaAtual += 1
        print(linhaAtual
        """
        