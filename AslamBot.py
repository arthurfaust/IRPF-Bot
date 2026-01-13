# Algorithm to open IRPF

import pywinauto, pyautogui, time
from pywinauto.keyboard import send_keys

cpf = "77039677008"

# Open Start Menu, search the IRPF program and open it
pyautogui.press('win')
time.sleep(1)
send_keys("IRPF")
time.sleep(1)
pyautogui.press('enter')

# Move to "Em Preenchimento"
time.sleep(15)
pyautogui.click(x=680,y=340)

# Search for the client by the CPF and open 
pyautogui.click(x=680,y=378)
send_keys(cpf)
pyautogui.press('enter')

time.sleep(1)
pyautogui.click(x=720,y=435)


""""
pyautogui.click(x=20,y=750)

or

pyautogui.moveTo(20,750)
pyautogui.click()
"""