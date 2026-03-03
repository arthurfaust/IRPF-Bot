import time
import FreeSimpleGUI as sg

def buscar_arquivo():
    sg.theme('Reddit')

    arquivo_escolhido = None
    
    layout = [
        # O 'Input' recebe o caminho e o 'FileBrowse' abre a busca
        [sg.Text('Selecione o arquivo que deseja fazer o lançamento:')],
        [sg.Text('Arquivo:'), sg.Input(key='caminho_arquivo'), sg.FileBrowse('Buscar')],
        [sg.Button('Executar')],
    ]

    janela = sg.Window('IRPF Bot', layout)

    while True:
        eventos, valores = janela.read()
    
        if eventos == sg.WIN_CLOSED: 
            break
        
        if eventos == 'Executar':
            # Agora você também pode acessar o caminho do arquivo selecionado
            arquivo_escolhido = valores['caminho_arquivo']
            print(f'Arquivo selecionado: {arquivo_escolhido}')
            janela.close()
            return arquivo_escolhido

    janela.close()
    return arquivo_escolhido
