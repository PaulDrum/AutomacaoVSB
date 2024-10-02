from selenium import webdriver
from time import sleep
from datetime import date, timedelta
from selenium.webdriver.common.by import By
import keyboard

yesterday = date.today() - timedelta(days=1)
d1 = yesterday.strftime("%d/%m/%Y")

chrome_options = webdriver.EdgeOptions()
chrome_options.add_argument('--start-maximized')
chrome_options.add_argument("--disable-infobars")
chrome_options.add_experimental_option("useAutomationExtension", False)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_argument("--window-position=-1200,0")
chrome_options.add_experimental_option("detach", True)

chrome_options2 = webdriver.EdgeOptions()
chrome_options2.add_argument('--start-maximized')
chrome_options2.add_argument("--disable-infobars")
chrome_options2.add_experimental_option("useAutomationExtension", False)
chrome_options2.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options2.add_argument("--window-position=-2500,0")
chrome_options2.add_experimental_option("detach", True)

navegador = webdriver.Edge(options=chrome_options)
navegador2 = webdriver.Edge(options=chrome_options2)

navegador.get("https://app.powerbi.com/groups/0c85d7de-c021-4a2b-b654-87787120d444/reports/074ff631-9fd2-43dc-9166-1c8095263fbf/ReportSection?experience=power-bi")

sleep(42)

navegador.find_element(By.ID, "collapsePagesPaneBtn").click()

sleep(1)

navegador.fullscreen_window()

sleep(2)

navegador2.get("https://app.powerbi.com/groups/0c85d7de-c021-4a2b-b654-87787120d444/reports/9a63463a-a5a8-499d-b04b-46ebe91d696f/ReportSection?experience=power-bi")

sleep(42)

navegador2.find_element(By.ID, "collapsePagesPaneBtn").click()

sleep(1)

navegador2.fullscreen_window()

while True:
    try:
        if keyboard.is_pressed('esc'):
            break
    except:
        break