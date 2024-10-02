import pyautogui
from time import sleep

sleep(5)

for i in range(1, 22):
    pyautogui.click(x=1760, y=514)
    sleep(1)
    pyautogui.doubleClick(x=282, y=647)
    sleep(0.5)
    pyautogui.hotkey("ctrl","a")
    sleep(0.5)
    pyautogui.hotkey("ctrl","v")
    sleep(0.5)
    pyautogui.click(x=1582, y=984)
    sleep(2)
