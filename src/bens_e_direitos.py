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
    QuantidadeLinhas = len(df) - linhaAtual
    
    print("Quantidade lançamentos: {}".format(QuantidadeLinhas))
    print("Linha Atual: {}".format(linhaAtual))
    
    # Enquanto a linha atual for menor que o tamanho do dataframe
    while linhaAtual < len(df):

        # Se a célula ATUAL estiver vazia, para a automação
        if pd.isna(df.iloc[linhaAtual, 0]): break
            
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
        time.sleep(0.5)
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
                send_keys(ImpostoExt)
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
            
            else:
                # Dá enter no autocustodiante
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá enter no CNPJ do custodiante
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche a discriminação
                time.sleep(0.25)
                send_keys(DiscAtual, with_spaces= True)
                time.sleep(0.25)
                pyautogui.press('enter')
                
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
                
                # Dá OK e finaliza o lançamento
                time.sleep(0.25)
                pyautogui.press('enter')
        
        if GrupoAtual == "08" and CodigoAtual == "02":
            if LocalAtual != "105 - Brasil":
                
                # Dá OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Código Altcoin"
                time.sleep(0.25)
                send_keys(CodAltcoin)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá enter em "autocustodiante?"
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Discriminação"
                time.sleep(0.25)
                send_keys(DiscAtual, with_spaces = True)
                time.sleep(0.25)
                pyautogui.press('enter')
                
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
                send_keys(ImpostoExt)
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
               
            else:
                # Preenche "Código Altcoin"
                time.sleep(0.25)
                send_keys(CodAltcoin)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá enter em "autocustodiante?"
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá enter em "CNPJ Custodiante"
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Discriminação"
                time.sleep(0.25)
                send_keys(DiscAtual, with_spaces = True)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(0.25)    
                pyautogui.press('backspace')
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Situação em 31/12/24"
                time.sleep(0.25)    
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(Sit24)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá tab em "Repetir"
                time.sleep(0.25)
                pyautogui.press('tab')
                
                # Dá OK e finaliza o lançamento
                time.sleep(0.25)
                pyautogui.press('enter')
                
        linhaAtual += 1
        print("Linha Atual: {}".format(linhaAtual))
    print("Lançamento finalizado.")