import pyautogui as gui


for win in gui.getAllWindows():
    if "1억" in win.title:
        print(win.title)
    