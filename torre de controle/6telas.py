from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from Recursos import AcesseMeuSiteTriyax
import keyboard

conta2 = "VALLOURECADM"
conta3 = "VALLOUREC-JECEABA"
usuario = "lucas.saraujo"
senha = "Lucas@2022"

# chrome_options = webdriver.EdgeOptions()
# chrome_options.add_argument('--start-maximized')
# chrome_options.add_argument("--disable-infobars")
# chrome_options.add_experimental_option("useAutomationExtension", False)
# chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
# chrome_options.add_argument("--window-position=0,0")
# chrome_options.add_experimental_option("detach", True)

# chrome_options2 = webdriver.EdgeOptions()
# chrome_options2.add_argument('--start-maximized')
# chrome_options2.add_argument("--disable-infobars")
# chrome_options2.add_experimental_option("useAutomationExtension", False)
# chrome_options2.add_experimental_option("excludeSwitches", ["enable-automation"])
# chrome_options2.add_argument("--window-position=0,700")
# chrome_options2.add_experimental_option("detach", True)

# chrome_options3 = webdriver.EdgeOptions()
# chrome_options3.add_argument('--start-maximized')
# chrome_options3.add_argument("--disable-infobars")
# chrome_options3.add_experimental_option("useAutomationExtension", False)
# chrome_options3.add_experimental_option("excludeSwitches", ["enable-automation"])
# chrome_options3.add_argument("--window-position=-1200,0")
# chrome_options3.add_experimental_option("detach", True)

# chrome_options4 = webdriver.EdgeOptions()
# chrome_options4.add_argument('--start-maximized')
# chrome_options4.add_argument("--disable-infobars")
# chrome_options4.add_experimental_option("useAutomationExtension", False)
# chrome_options4.add_experimental_option("excludeSwitches", ["enable-automation"])
# chrome_options4.add_argument("--window-position=-1200,700")
# chrome_options4.add_experimental_option("detach", True)

# chrome_options5 = webdriver.EdgeOptions()
# chrome_options5.add_argument('--start-maximized')
# chrome_options5.add_argument("--disable-infobars")
# chrome_options5.add_experimental_option("useAutomationExtension", False)
# chrome_options5.add_experimental_option("excludeSwitches", ["enable-automation"])
# chrome_options5.add_argument("--window-position=-2500,0")
# chrome_options5.add_experimental_option("detach", True)

# chrome_options6 = webdriver.EdgeOptions()
# chrome_options6.add_argument('--start-maximized')
# chrome_options6.add_argument("--disable-infobars")
# chrome_options6.add_experimental_option("useAutomationExtension", False)
# chrome_options6.add_experimental_option("excludeSwitches", ["enable-automation"])
# chrome_options6.add_argument("--window-position=-2500,700")
# chrome_options6.add_experimental_option("detach", True)

navegador = webdriver.Edge(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\msedgedriver.exe')
navegador2 = webdriver.Edge(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\msedgedriver.exe')
navegador3 = webdriver.Edge(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\msedgedriver.exe')
navegador4 = webdriver.Edge(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\msedgedriver.exe')
navegador5 = webdriver.Edge(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\msedgedriver.exe')
navegador6 = webdriver.Edge(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\msedgedriver.exe')

navegador.get("https://app.powerbi.com/groups/0c85d7de-c021-4a2b-b654-87787120d444/reports/29ad6040-0b72-41b2-9f40-69cb03ecbad7/ReportSection?experience=power-bi")
navegador2.get("https://app.powerbi.com/groups/0c85d7de-c021-4a2b-b654-87787120d444/reports/29ad6040-0b72-41b2-9f40-69cb03ecbad7/ReportSection?experience=power-bi")
navegador3.get("https://app.powerbi.com/groups/0c85d7de-c021-4a2b-b654-87787120d444/reports/6822ff89-6f77-4f18-b747-41d4cf577874/ReportSectionc94d4414c8870c674b71?experience=power-bi")
navegador4.get("https://app.powerbi.com/groups/0c85d7de-c021-4a2b-b654-87787120d444/reports/6822ff89-6f77-4f18-b747-41d4cf577874/ReportSectionc94d4414c8870c674b71?experience=power-bi")
navegador5.get("https://triyax.com.br/#/login")
navegador6.get("https://triyax.com.br/#/login")

sleep(3)

navegador.set_window_position(0, 0)
navegador2.set_window_position(0, 760)
navegador3.set_window_position(-1200, 0)
navegador4.set_window_position(-1200, 760)
navegador5.set_window_position(-2500, 0)
navegador6.set_window_position(-2500, 760)

sleep(22)

AcesseMeuSiteTriyax.Logar(navegador5, conta2, usuario, senha)
AcesseMeuSiteTriyax.Logar(navegador6, conta3, usuario, senha)

sleep(22)

navegador.find_element(By.ID, "collapsePagesPaneBtn").click()

sleep(5)

navegador2.find_element(By.XPATH, "//*[contains(text(),'JECEABA')]").click()

sleep(5)

navegador2.find_element(By.ID, "collapsePagesPaneBtn").click()

sleep(3)

navegador5.find_element(By.XPATH, "/html/body/fury-root/fury-layout/div/mat-sidenav-container/mat-sidenav-content/div/div/fleetmonitor-wrapper/monitor-mapa-online/mat-sidenav-container/mat-sidenav-content/div/div[2]/div/button[2]").click()

sleep(3)

navegador5.find_element(By.XPATH, "//input[@type='checkbox']").click()

sleep(3)

navegador5.find_element(By.XPATH, "//div//label").click()

sleep(5)

navegador6.find_element(By.XPATH, "/html/body/fury-root/fury-layout/div/mat-sidenav-container/mat-sidenav-content/div/div/fleetmonitor-wrapper/monitor-mapa-online/mat-sidenav-container/mat-sidenav-content/div/div[2]/div/button[2]").click()

sleep(3)

navegador6.find_element(By.XPATH, "//input[@type='checkbox']").click()

sleep(3)

navegador6.find_element(By.XPATH, "//div//label").click()

sleep(1)

navegador.find_element(By.XPATH, "//button[@aria-label='Exibição']").click()
sleep(0.5)
navegador.find_element(By.XPATH, "//button[@title='Abrir no modo de tela inteira']").click()
sleep(0.5)
navegador2.find_element(By.XPATH, "//button[@aria-label='Exibição']").click()
sleep(0.5)
navegador2.find_element(By.XPATH, "//button[@title='Abrir no modo de tela inteira']").click()
sleep(0.5)
navegador3.find_element(By.XPATH, "//button[@aria-label='Exibição']").click()
sleep(0.5)
navegador3.find_element(By.XPATH, "//button[@title='Abrir no modo de tela inteira']").click()
sleep(0.5)
navegador4.find_element(By.XPATH, "//button[@aria-label='Exibição']").click()
sleep(0.5)
navegador4.find_element(By.XPATH, "//button[@title='Abrir no modo de tela inteira']").click()
sleep(0.5)
navegador5.fullscreen_window()
sleep(0.5)
navegador6.fullscreen_window()
sleep(0.5)
navegador5.find_element(By.XPATH, "//button").click()
sleep(0.5)
navegador6.find_element(By.XPATH, "//button").click()

while True:
    try:
        if keyboard.is_pressed('esc'):
            break
    except:
        break