# Configura o ambiente para iniciar a automação
import pywinauto, pyautogui, time
import pandas as pd
from pywinauto.keyboard import send_keys

screen_size = pyautogui.size()
print(screen_size)
