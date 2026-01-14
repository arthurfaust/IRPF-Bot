# Algorithm to open IRPF

import pywinauto, pyautogui, time
from pywinauto.keyboard import send_keys

cpf = "770.396.770-08"

# Open Start Menu, search the IRPF program and open it
pyautogui.press('win')
time.sleep(0.5)
send_keys("IRPF")
time.sleep(0.5)
pyautogui.press('enter')

# Move to "Em Preenchimento" and click it
time.sleep(15)
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

# Move to the arrow and scroll the menu down
time.sleep(1)
pyautogui.moveTo(x=350,y=450)
pyautogui.click(clicks=16)

# Move to "Bens e Direitos" and click it
time.sleep(0.5)
pyautogui.click(x=85,y=440)

# Move to "Novo" and click it
time.sleep(0.5)
pyautogui.click(x=1105,y=665)