import buscar_arquivos
import pywinauto, pyautogui, time
import pandas as pd
from pywinauto.keyboard import send_keys

def AbrirIRPF():
    # Open Start Menu, search the IRPF program and open it
    time.sleep(0.5)
    send_keys('{VK_LWIN}')
    time.sleep(0.5)
    send_keys('IRPF2025')
    time.sleep(0.5)
    send_keys('{ENTER}')
    
def AbrirDeclaracaoCliente(caminho):
    
    df = pd.read_excel(caminho, header=None)
    nome = df.iloc[1,0]
    # Trigger para abrir a declaração do cliente
    # Cor do pixel de trigger (680,195) = (31, 47, 101)
    while True:
        if pyautogui.pixelMatchesColor(680,195,(31, 47, 101)) == True:
            
            # Move para "Em Preenchimento"
            time.sleep(0.5)
            send_keys('%e')
            
            # Move para a barra de pesquisa e procura o nome do cliente
            time.sleep(0.5)
            send_keys('{ENTER}')
            time.sleep(0.5)
            send_keys(nome, with_spaces=True)
            time.sleep(0.5)
            send_keys('{ENTER}')
            time.sleep(0.5)
            send_keys("{TAB}")
            
            # Move to the client found and double click to open it
            time.sleep(0.5)
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
        cnpj_atual = "28640024000124"
        tempo_menor = 0.5
        tempo_maior = 0.5
        
        # Move o cursor para "Novo" e clica
        time.sleep(tempo_maior)
        send_keys('%n')
        
        # Fazer um trigger aqui para a transição de tela
        
        # Preenche o grupo atual
        time.sleep(tempo_maior)
        send_keys(grupo_atual)
        time.sleep(tempo_menor)
        send_keys('{ENTER}')
        
        if grupo_atual == "03" and codigo_atual == "01":
            # Preenche o código atual
            time.sleep(tempo_menor)
            send_keys(codigo_atual)
            time.sleep(tempo_menor)
            send_keys('{DOWN}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a localização atual
            time.sleep(tempo_menor)
            send_keys('^A')
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(local_atual, with_spaces=True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            if local_atual != "105 - Brasil":
                # Dá OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche a discriminação
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces= True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter em "Negociados na bolsa?"
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Lucro ou Prejuízo"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(lucro_ou_prejuizo)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Valor Recebido"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(valor_recebido)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_pago_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)        
                send_keys('{ENTER}')
                
            else:
                # Preenche o CNPJ
                time.sleep(tempo_menor)
                send_keys(cnpj_atual)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche a discriminação
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces= True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter em "Negociados na bolsa?"
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # Dá tab em "Informar Rendimento Isento"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # Dá tab em "Informar Rendimento Exclusivo"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
         
        if grupo_atual == "07" and codigo_atual == "03":
            # Preenche o código atual
            time.sleep(tempo_menor)
            send_keys(codigo_atual)
            time.sleep(tempo_menor)
            send_keys('{DOWN}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche o CNPJ do fundo
            time.sleep(tempo_menor)
            send_keys(cnpj_atual)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a discriminação
            time.sleep(tempo_menor)
            send_keys(disc_atual, with_spaces= True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Dá enter em "Negociados na bolsa?"
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Apaga "Situação em 31/12/23"
            time.sleep(tempo_menor)    
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Preenche "Situação em 31/12/24"
            time.sleep(tempo_menor)    
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(sit_24)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Dá tab em "Repetir"
            time.sleep(tempo_menor)
            send_keys('{TAB}')
                
            # Dá tab em "Informar Rendimento Isento"
            time.sleep(tempo_menor)
            send_keys('{TAB}')
                
            # Dá tab em "Informar Rendimento Exclusivo"
            time.sleep(tempo_menor)
            send_keys('{TAB}')
                
            # Dá OK e finaliza o lançamento
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
        if (grupo_atual == "07" and codigo_atual == "07") or (grupo_atual == "07" and codigo_atual == "08") :
            # Preenche o código atual
            time.sleep(tempo_menor)
            send_keys(codigo_atual)
            time.sleep(tempo_menor)
            send_keys('{DOWN}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche o CNPJ do fundo
            time.sleep(tempo_menor)
            send_keys(cnpj_atual)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a discriminação
            time.sleep(tempo_menor)
            send_keys(disc_atual, with_spaces= True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Dá enter em "Negociados na bolsa?"
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Apaga "Situação em 31/12/23"
            time.sleep(tempo_menor)    
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Preenche "Situação em 31/12/24"
            time.sleep(tempo_menor)    
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(sit_24)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Dá tab em "Repetir"
            time.sleep(tempo_menor)
            send_keys('{TAB}')
                
            # Dá tab em "Informar Rendimento Exclusivo"
            time.sleep(tempo_menor)
            send_keys('{TAB}')
                
            # Dá OK e finaliza o lançamento
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
        
        if grupo_atual == "07" and codigo_atual == "99":
            # Preenche o código atual
            time.sleep(tempo_menor)
            send_keys(codigo_atual)
            time.sleep(tempo_menor)
            
            for i in range(5):
                send_keys('{DOWN}')
                
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a localização atual
            time.sleep(tempo_menor)
            send_keys('^A')
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(local_atual, with_spaces=True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # OK no aviso
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a discriminação
            time.sleep(tempo_menor)
            send_keys(disc_atual, with_spaces= True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Dá enter em "Negociados na bolsa?"
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Apaga "Situação em 31/12/23"
            time.sleep(tempo_menor)    
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Preenche "Situação em 31/12/24"
            time.sleep(tempo_menor)    
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(sit_24)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Dá tab em "Repetir"
            time.sleep(tempo_menor)
            send_keys('{TAB}')
            
            # OK no aviso
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Preenche "Lucro ou Prejuízo"
            time.sleep(tempo_menor)
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(lucro_ou_prejuizo)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # OK no aviso
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Preenche "Imposto pago no Exterior"
            time.sleep(tempo_menor)
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(imposto_ext)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # OK no aviso
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Preenche "Valor Recebido"
            time.sleep(tempo_menor)
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(valor_recebido)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # OK no aviso
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Preenche "Imposto Pago no Exterior"
            time.sleep(tempo_menor)
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(imposto_pago_ext)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
                
            # Dá OK e finaliza o lançamento
            time.sleep(tempo_menor)        
            send_keys('{ENTER}')
        
        if (grupo_atual == "08" and codigo_atual == "01") or (grupo_atual == "08" and codigo_atual == "10"):
            # Preenche o código atual
            time.sleep(tempo_menor)
            send_keys(codigo_atual)
            time.sleep(tempo_menor)
            send_keys('{DOWN}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a localização atual
            time.sleep(tempo_menor)
            send_keys('^A')
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(local_atual, with_spaces=True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            if local_atual != "105 - Brasil":
                # Dá OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter no autocustodiante
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche a discriminação
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces= True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24" e vai para "Repetir"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Lucro ou Prejuízo"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(lucro_ou_prejuizo)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Valor Recebido"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(valor_recebido)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_pago_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)        
                send_keys('{ENTER}')
            
            else:
                # Dá enter no autocustodiante
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter no CNPJ do custodiante
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche a discriminação
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces= True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24" e vai para "Repetir"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
        
        if grupo_atual == "08" and codigo_atual == "02":
            # Preenche o código atual
            time.sleep(tempo_menor)
            send_keys(codigo_atual)
            time.sleep(tempo_menor)
            send_keys('{DOWN}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a localização atual
            time.sleep(tempo_menor)
            send_keys('^A')
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(local_atual, with_spaces=True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            if local_atual != "105 - Brasil":
                
                # Dá OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Código Altcoin"
                time.sleep(tempo_menor)
                send_keys(codigo_altcoin)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter em "autocustodiante?"
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Discriminação"
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces = True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24" e vai para "Repetir"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Lucro ou Prejuízo"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(lucro_ou_prejuizo)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Valor Recebido"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(valor_recebido)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_pago_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)        
                send_keys('{ENTER}')
               
            else:
                # Preenche "Código Altcoin"
                time.sleep(tempo_menor)
                send_keys(codigo_altcoin)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter em "autocustodiante?"
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter em "CNPJ Custodiante"
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Discriminação"
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces = True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
        if grupo_atual == "08" and codigo_atual == "03":
            # Preenche o código atual
            time.sleep(tempo_menor)
            send_keys(codigo_atual)
            time.sleep(tempo_menor)
            send_keys('{DOWN}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a localização atual
            time.sleep(tempo_menor)
            send_keys('^A')
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(local_atual, with_spaces=True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            if local_atual != "105 - Brasil":
                
                # Dá OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Código Altcoin"
                time.sleep(tempo_menor)
                send_keys(codigo_stablecoin)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter em "autocustodiante?"
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Discriminação"
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces = True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24" e vai para "Repetir"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Lucro ou Prejuízo"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(lucro_ou_prejuizo)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Valor Recebido"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(valor_recebido)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_pago_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)        
                send_keys('{ENTER}')
               
            else:
                # Preenche "Código Altcoin"
                time.sleep(tempo_menor)
                send_keys(codigo_stablecoin)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter em "autocustodiante?"
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá enter em "CNPJ Custodiante"
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Discriminação"
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces = True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
        
        if grupo_atual == "08" and codigo_atual == "99":
            # Preenche o código atual
            time.sleep(tempo_menor)
            send_keys(codigo_atual)
            time.sleep(tempo_menor)
            send_keys('{DOWN}')
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            # Preenche a localização atual
            time.sleep(tempo_menor)
            send_keys('^A')
            send_keys('{BACKSPACE}')
            time.sleep(tempo_menor)
            send_keys(local_atual, with_spaces=True)
            time.sleep(tempo_menor)
            send_keys('{ENTER}')
            
            if local_atual != "105 - Brasil":
                # Dá OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche a discriminação
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces= True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24" e vai para "Repetir"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Lucro ou Prejuízo"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(lucro_ou_prejuizo)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Valor Recebido"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(valor_recebido)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # OK no aviso
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Imposto Pago no Exterior"
                time.sleep(tempo_menor)
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(imposto_pago_ext)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)        
                send_keys('{ENTER}')
            
            else:
                
                # Preenche a discriminação
                time.sleep(tempo_menor)
                send_keys(disc_atual, with_spaces= True)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Apaga "Situação em 31/12/23" e vai para "Situação em 31/12/24"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Preenche "Situação em 31/12/24" e vai para "Repetir"
                time.sleep(tempo_menor)    
                send_keys('{BACKSPACE}')
                time.sleep(tempo_menor)
                send_keys(sit_24)
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
                
                # Dá tab em "Repetir"
                time.sleep(tempo_menor)
                send_keys('{TAB}')
                
                # Dá OK e finaliza o lançamento
                time.sleep(tempo_menor)
                send_keys('{ENTER}')
        
        linha_atual += 1
        print("Linha Atual: {}".format(linha_atual))
    print("Lançamento finalizado.")
    lancamento_finalizado = True
    
    return lancamento_finalizado