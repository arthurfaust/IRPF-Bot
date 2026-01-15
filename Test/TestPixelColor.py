import pywinauto, pyautogui, time

cor = pyautogui.screenshot()
print(cor.getpixel((215,90)))

while True:
    if pyautogui.pixelMatchesColor(215,90,(204, 227, 252)) == True:
        print("oi teste acorda duda")