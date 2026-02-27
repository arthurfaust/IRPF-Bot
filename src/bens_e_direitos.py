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
    linha_atual = 18
    QuantidadeLinhas = len(df) - linha_atual
    
    print("Quantidade lançamentos: {}".format(QuantidadeLinhas))
    print("Linha Atual: {}".format(linha_atual))
    
    # Enquanto a linha atual for menor que o tamanho do dataframe
    while linha_atual < len(df):

        # Se a célula ATUAL estiver vazia, para a automação
        if pd.isna(df.iloc[linha_atual, 0]): break
            
        grupo_atual = df.iloc[linha_atual,0]
        codigo_atual = df.iloc[linha_atual,1]
        local_atual = df.iloc[linha_atual,2]
        disc_atual = df.iloc[linha_atual,4]
        sit_24 = df.iloc[linha_atual,6]
        lucro_ou_prejuizo = "100,00"
        imposto_ext = "200,00"
        valor_recebido = "300,00"
        imposto_pago_ext = "400,00"
        codigo_altcoin = "ETH"
        codigo_stablecoin = "USDT"
        
        # Move o cursor para "Novo" e clica
        time.sleep(0.5)
        pyautogui.click(x=1105,y=665)
        
        # Fazer um trigger aqui para a transição de tela
        
        # Preenche o grupo atual
        time.sleep(0.5)
        send_keys(grupo_atual)
        time.sleep(0.25)
        pyautogui.press('enter')
        
        # Preenche o código atual
        time.sleep(0.25)
        send_keys(codigo_atual)
        time.sleep(0.25)
        pyautogui.press('enter')
        
        # Move o cursor para "Localização" e digita a loc atual
        time.sleep(0.25)
        send_keys('^A')
        pyautogui.press('backspace')
        time.sleep(0.25)
        send_keys(local_atual, with_spaces=True)
        time.sleep(0.25)
        pyautogui.press('enter')
        
        if (grupo_atual == "08" and codigo_atual == "01") or (grupo_atual == "08" and codigo_atual == "10"):
            if local_atual != "105 - Brasil":
                # Dá OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá enter no autocustodiante
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche a discriminação
                time.sleep(0.25)
                send_keys(disc_atual, with_spaces= True)
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
                send_keys(sit_24)
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
                send_keys(lucro_ou_prejuizo)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(imposto_ext)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Valor Recebido"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(valor_recebido)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(imposto_pago_ext)
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
                send_keys(disc_atual, with_spaces= True)
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
                send_keys(sit_24)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá tab em "Repetir"
                time.sleep(0.25)
                pyautogui.press('tab')
                
                # Dá OK e finaliza o lançamento
                time.sleep(0.25)
                pyautogui.press('enter')
        
        if grupo_atual == "08" and codigo_atual == "02":
            if local_atual != "105 - Brasil":
                
                # Dá OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Código Altcoin"
                time.sleep(0.25)
                send_keys(codigo_altcoin)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá enter em "autocustodiante?"
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Discriminação"
                time.sleep(0.25)
                send_keys(disc_atual, with_spaces = True)
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
                send_keys(sit_24)
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
                send_keys(lucro_ou_prejuizo)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(imposto_ext)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Valor Recebido"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(valor_recebido)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(imposto_pago_ext)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá OK e finaliza o lançamento
                time.sleep(0.25)        
                pyautogui.press('enter')
               
            else:
                # Preenche "Código Altcoin"
                time.sleep(0.25)
                send_keys(codigo_altcoin)
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
                send_keys(disc_atual, with_spaces = True)
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
                send_keys(sit_24)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá tab em "Repetir"
                time.sleep(0.25)
                pyautogui.press('tab')
                
                # Dá OK e finaliza o lançamento
                time.sleep(0.25)
                pyautogui.press('enter')
                
        if grupo_atual == "08" and codigo_atual == "03":
            if local_atual != "105 - Brasil":
                
                # Dá OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Código Altcoin"
                time.sleep(0.25)
                send_keys(codigo_stablecoin)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá enter em "autocustodiante?"
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Discriminação"
                time.sleep(0.25)
                send_keys(disc_atual, with_spaces = True)
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
                send_keys(sit_24)
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
                send_keys(lucro_ou_prejuizo)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(imposto_ext)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Valor Recebido"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(valor_recebido)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # OK no aviso
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(0.25)
                pyautogui.press('backspace')
                time.sleep(0.25)
                send_keys(imposto_pago_ext)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá OK e finaliza o lançamento
                time.sleep(0.25)        
                pyautogui.press('enter')
               
            else:
                # Preenche "Código Altcoin"
                time.sleep(0.25)
                send_keys(codigo_stablecoin)
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
                send_keys(disc_atual, with_spaces = True)
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
                send_keys(sit_24)
                time.sleep(0.25)
                pyautogui.press('enter')
                
                # Dá tab em "Repetir"
                time.sleep(0.25)
                pyautogui.press('tab')
                
                # Dá OK e finaliza o lançamento
                time.sleep(0.25)
                pyautogui.press('enter')
                
        linha_atual += 1
        print("Linha Atual: {}".format(linha_atual))
    print("Lançamento finalizado.")