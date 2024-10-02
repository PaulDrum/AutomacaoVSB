from time import sleep
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from suporte.Recursos import TelaPython
from suporte.Recursos import AcesseMeuSite
from suporte.Recursos import AcesseMeuSiteAnalitico
import os, pyperclip, pyautogui

user = os.path.expanduser('~')

tela = TelaPython()
tela.Iniciar()

conta_barreiro = 'VALLOUREC-GESTAO'
conta_jeceaba = 'VALLOUREC-JECEABA'

aux = '1'

if tela.values['relatorio_resumo'] == True and tela.values['relatorio_ocupacao'] == False and tela.values['relatorio_carretinhas'] == False and tela.values['relatorio_permanencia'] == False:
    
    if user == 'C:\\Users\\Vitor.Valadares':
        navegador = webdriver.Chrome(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\vitor\\chromedriver.exe')
    
    else:
        navegador = webdriver.Chrome(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\chromedriver.exe')
    
    navegador.get('https://web.pdo-vallourec.com.br/#/')
    
    AcesseMeuSite.Logar(navegador, conta_barreiro, tela.values['usuario'], tela.values['senha'])
    sleep(30)
    
    navegador.find_element_by_xpath('//*[@id="mat-input-3"]').send_keys(aux, Keys.LEFT_CONTROL + 'a', Keys.DELETE)
    sleep(1)
    navegador.find_element_by_xpath('//*[@id="mat-input-3"]').send_keys(tela.values['data_inicial'])
    sleep(1)
    
    navegador.find_element_by_xpath('//*[@id="mat-input-4"]').send_keys(aux, Keys.LEFT_CONTROL + 'a', Keys.DELETE)
    sleep(1)
    navegador.find_element_by_xpath('//*[@id="mat-input-4"]').send_keys(tela.values['data_final'])
    sleep(1)
   
    navegador.find_element_by_xpath('//*[@id="mat-select-0"]/div/div[1]').click()
    
    navegador.find_element_by_xpath('//*[@id="mat-option-0"]/span').click()
    
    navegador.find_element_by_xpath('/html/body/div[2]/div[1]').click()
    
    navegador.find_element_by_xpath('//*[@id="mat-expansion-panel-header-1"]/span[1]/mat-panel-title/span').click()
    
    navegador.find_element_by_xpath('//*[@id="cdk-accordion-child-1"]/div/pdo-form-content/mat-card-content/div[2]/pdo-button/button').click()
    sleep(20)
    
    navegador.find_element_by_xpath('/html/body/app-root/ms-renderer/mat-sidenav-container/mat-sidenav-content/div/pdo-renderer-toolbar/div/div[3]/ms-toolbar-user-button/div/button').click()
    navegador.find_element_by_xpath('/html/body/app-root/ms-renderer/mat-sidenav-container/mat-sidenav-content/div/pdo-renderer-toolbar/div/div[3]/ms-toolbar-user-button/div/div/div/div/div[3]').click()
    sleep(20)
    
    AcesseMeuSite.Logar(navegador, conta_jeceaba, tela.values['usuario'], tela.values['senha'])
    sleep(15)
    
    pyperclip.copy(tela.values['data_inicial'])
    sleep(1)
    navegador.find_element_by_xpath('//*[@id="mat-input-3"]').send_keys(Keys.LEFT_CONTROL + 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(2)

    pyperclip.copy(tela.values['data_final'])
    sleep(1)
    navegador.find_element_by_xpath('//*[@id="mat-input-4"]').send_keys(Keys.LEFT_CONTROL + 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(2)

    navegador.find_element_by_xpath('//*[@id="mat-expansion-panel-header-1"]').click()
    
    navegador.find_element_by_xpath('//*[@id="cdk-accordion-child-1"]/div/form/button').click()
    sleep(20)
    
    navegador.close()

elif tela.values['relatorio_resumo'] == False and tela.values['relatorio_carretinhas'] == False and tela.values['relatorio_ocupacao'] == True and tela.values['relatorio_permanencia'] == False:
    
    if user == 'C:\\Users\\Vitor.Valadares':
        navegador = webdriver.Chrome(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\vitor\\chromedriver.exe')
    
    else:
        navegador = webdriver.Chrome(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\chromedriver.exe')
    
    navegador.get('https://web.pdo-vallourec.com.br/#/')
    
    AcesseMeuSite.Logar(navegador, conta_jeceaba, tela.values['usuario'], tela.values['senha'])
    sleep(15)
    
    pyperclip.copy(tela.values['data_inicial'])
    sleep(1)
    
    navegador.find_element_by_xpath('//*[@id="mat-input-3"]').send_keys(Keys.LEFT_CONTROL + 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(2)

    pyperclip.copy(tela.values['data_final'])
    sleep(1)
    navegador.find_element_by_xpath('//*[@id="mat-input-4"]').send_keys(Keys.LEFT_CONTROL + 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(2)

    navegador.find_element_by_xpath('//*[@id="mat-expansion-panel-header-10"]').click()
    navegador.find_element_by_xpath('//*[@id="cdk-accordion-child-10"]/div/form/mat-form-field/div/div[1]/div').click()
    navegador.find_element_by_xpath('//*[@id="mat-option-20"]').click()
    navegador.find_element_by_xpath('//*[@id="cdk-accordion-child-10"]/div/form/button').click()
    sleep(15)
                
    navegador.get('https://analitico-vallourec.com.br/login/')
    AcesseMeuSiteAnalitico.Logar(navegador, tela.values['usuario'].lower(), tela.values['senha'])
    sleep(10)
    
    navegador.find_element_by_xpath('/html/body/nav/div[2]/div/ul/li[1]').click()
    sleep(5)
                
    navegador.find_element_by_xpath('/html/body/nav/div[2]/div/ul/li[1]').click()
    sleep(5)
                
    pyperclip.copy(tela.values['data_inicial'])
    sleep(1)
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/div/p[1]/input').send_keys(Keys.LEFT_CONTROL + 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(2)
                
    pyperclip.copy(tela.values['data_final'])
    sleep(1)
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/div/p[2]/input').send_keys(Keys.LEFT_CONTROL + 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(2)
                
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[20]/div[1]/h4/a').click()
    sleep(3)
    
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[20]/div[2]/div/div/select').click()
    pyautogui.hotkey('down', 'enter')
    sleep(1)
                
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[20]/div[2]/div/button').click()
    sleep(5)
                
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[20]/div[2]/div/div/select').click()
    pyautogui.hotkey('down', 'enter')
    sleep(1)
                
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[20]/div[2]/div/button').click()
    sleep(5)
                
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[20]/div[2]/div/div/select').click()
    pyautogui.hotkey('down', 'enter')
    sleep(1)
    
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[20]/div[2]/div/button').click()
    sleep(15)
    
    navegador.close()

elif tela.values['relatorio_carretinhas'] == True and tela.values['relatorio_resumo'] == False and tela.values['relatorio_ocupacao'] == False and tela.values['relatorio_permanencia'] == False:
    
    if user == 'C:\\Users\\Vitor.Valadares':
        navegador = webdriver.Chrome(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\vitor\\chromedriver.exe')
    
    else:
        navegador = webdriver.Chrome(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\chromedriver.exe')
    
    navegador.get('https://analitico-vallourec.com.br/login/')
    
    AcesseMeuSiteAnalitico.Logar(navegador, tela.values['usuario'].lower(), tela.values['senha'])
    sleep(20)
    
    navegador.find_element_by_xpath('/html/body/nav/div[2]/div/ul/li[1]').click()   
    sleep(5)
        
    navegador.find_element_by_xpath('/html/body/nav/div[2]/div/ul/li[1]').click()
    sleep(5)
            
    pyperclip.copy(tela.values['data_inicial'])
    sleep(1)
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/div/p[1]/input').send_keys(Keys.LEFT_CONTROL + 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(2)
    
    pyperclip.copy(tela.values['data_final'])
    sleep(1)
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/div/p[2]/input').send_keys(Keys.LEFT_CONTROL + 'a')
    sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    sleep(2)
    
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[14]/div[1]').click()
    sleep(3)
    
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[14]/div[2]/div/button').click()
    sleep(99)
    
    navegador.close()

elif tela.values['relatorio_carretinhas'] == False and tela.values['relatorio_resumo'] == False and tela.values['relatorio_ocupacao'] == False and tela.values['relatorio_permanencia'] == True:
    
    user = os.path.expanduser('~')


    usuario = "VITOR.VALADARES"
    senha = "Vitor@2022"
    conta_barreiro = "VALLOUREC-GESTAO"
    conta_jeceaba = "VALLOUREC-JECEABA"

    navegador = webdriver.Chrome()

    navegador.get('https://analitico-vallourec.com.br/login/')

    #yesterday = date.today() - timedelta(days=1)
    #d1 = yesterday.strftime("%d/%m/%Y")

    action = ActionChains(navegador)

    AcesseMeuSiteAnalitico.Logar(navegador, usuario.lower(), senha)

    sleep(20)

    navegador.find_element_by_xpath('/html/body/nav/div[2]/div/ul/li[1]').click()

    sleep(5)

    navegador.find_element_by_xpath('/html/body/nav/div[2]/div/ul/li[1]').click()

    sleep(5)

    pyperclip.copy(tela.values['data_inicial'])

    sleep(1)

    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/div/p[1]/input').send_keys(Keys.LEFT_CONTROL + "a")

    sleep(1)

    pyautogui.hotkey("ctrl","v")

    sleep(5)

    pyperclip.copy(tela.values['data_final'])

    sleep(1)

    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/div/p[2]/input').send_keys(Keys.LEFT_CONTROL + "a")

    sleep(1)

    pyautogui.hotkey("ctrl","v")

    sleep(5)

    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[1]').click()

    sleep(3)

    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[21]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[22]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[23]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[24]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[25]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[33]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[34]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[35]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[36]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[37]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[38]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[39]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[40]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[41]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[42]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[43]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[44]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[45]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[46]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[47]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[48]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[49]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[50]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[51]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[52]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[53]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[54]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[55]/input').click()
    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[56]/input').click()

    sleep(5)

    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/button[3]').click()

    sleep(120)

    source = navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/button[2]')

    sleep(1)

    action.double_click(source).perform()

    sleep(1)

    for i in range(128, 153):
        navegador.find_element_by_xpath(f'/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/div/div[{i}]/input').click()

    sleep(1)

    navegador.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div/uib-accordion/div/div[27]/div[2]/div/button[3]').click()

    sleep(120)

    navegador.close()

elif tela.values ['relatorio_manutencao'] == True:

    conta_barreiro = 'VALLOURECADM'
    conta_jeceaba = 'VALLOUREC-JECEABA'
    usuario = 'lucas.saraujo'
    senha = 'Lucas@2022'

    navegador = webdriver.Chrome(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\chromedriver.exe')

    navegador.get('https://fleetmonitor.ddmx.com.br/#/')

    navegador.find_element_by_xpath('//*[@id="mat-input-0"]').send_keys(conta_barreiro)
    navegador.find_element_by_xpath('//*[@id="mat-input-1"]').send_keys(usuario)
    navegador.find_element_by_xpath('//*[@id="mat-input-2"]').send_keys(senha)
    navegador.find_element_by_xpath('/html/body/fury-root/monitor-login/div[2]/form/div/button').click()
    sleep(15)

    navegador.find_element_by_xpath('/html/body/fury-root/fury-layout/div/mat-sidenav-container/mat-sidenav[1]/div/fury-sidenav/div/fury-scrollbar/div[1]/div[2]/div/div/div/fury-sidenav-item[9]/div/a').click()
    sleep(5)

    navegador.find_element_by_xpath('//*[@id="mat-tab-content-0-0"]/div/mat-card/monitor-data-tree/div/map-tree/data-tree/data-group/div[1]/input').click()
    sleep(3)

    navegador.find_element_by_xpath('//*[@id="mat-tab-label-0-1"]').click()
    sleep(2)

    navegador.find_element_by_xpath("//*[contains(text(),\'Relatório de permanência em áreas e pontos de interesse')]").click()
    sleep(2)

    navegador.find_element_by_xpath("//*[contains(text(),\' Selecionar áreas de interesse')]").click()
    sleep(2)

    navegador.find_element_by_xpath('//*[@id="mat-checkbox-21"]/label/div').click()
    navegador.find_element_by_xpath('//*[@id="mat-checkbox-22"]/label/div').click()

    navegador.find_element_by_xpath('//*[@id="mat-dialog-0"]/monitor-selecionar-areas-dialog/mat-dialog-actions/button[1]').click()
    sleep(2)

    navegador.find_element_by_xpath("//*[contains(text(),\' Gerar relatório')]").click()
    sleep(99)

    navegador.find_element_by_xpath('//*[@id="mat-tab-content-0-2"]/div/mat-card/button[1]').click()
    sleep(5)

    navegador.find_element_by_xpath('//*[@id="mat-tab-content-0-2"]/div/table/tbody/tr[2]/td[6]/div/i/span').click()
    sleep(15)

    navegador.find_element_by_xpath('//*[@id="mat-tab-content-0-2"]/div/mat-card/button[2]').click()
    sleep(2)

    pyautogui.click(x=618, y=187)
    sleep(5)

    navegador.find_element_by_xpath('/html/body/fury-root/fury-layout/div/mat-sidenav-container/mat-sidenav[1]/div/fury-sidenav/div/fury-scrollbar/div[1]/div[2]/div/div/div/div/div/a[2]').click()
    sleep(22)

    navegador.find_element_by_xpath('//*[@id="mat-input-0"]').send_keys(conta_jeceaba)
    navegador.find_element_by_xpath('//*[@id="mat-input-1"]').send_keys(usuario)
    navegador.find_element_by_xpath('//*[@id="mat-input-2"]').send_keys(senha)
    navegador.find_element_by_xpath('/html/body/fury-root/monitor-login/div[2]/form/div/button').click()
    sleep(15)

    navegador.find_element_by_xpath('/html/body/fury-root/fury-layout/div/mat-sidenav-container/mat-sidenav[1]/div/fury-sidenav/div/fury-scrollbar/div[1]/div[2]/div/div/div/fury-sidenav-item[9]/div/a').click()
    sleep(5)

    navegador.find_element_by_xpath('//*[@id="mat-tab-content-0-0"]/div/mat-card/monitor-data-tree/div/map-tree/data-tree/data-group/div[1]/input').click()
    sleep(3)

    navegador.find_element_by_xpath('//*[@id="mat-tab-label-0-1"]').click()
    sleep(2)

    navegador.find_element_by_xpath("//*[contains(text(),\'Relatório de permanência em áreas e pontos de interesse')]").click()
    sleep(2)

    navegador.find_element_by_xpath("//*[contains(text(),\' Selecionar áreas de interesse')]").click()
    sleep(2)

    navegador.find_element_by_xpath('//*[@id="mat-checkbox-57"]/label/div').click()

    navegador.find_element_by_xpath('//*[@id="mat-dialog-0"]/monitor-selecionar-areas-dialog/mat-dialog-actions/button[1]').click()
    sleep(2)

    navegador.find_element_by_xpath("//*[contains(text(),\' Gerar relatório')]").click()
    sleep(99)

    navegador.find_element_by_xpath('//*[@id="mat-tab-content-0-2"]/div/mat-card/button[1]').click()
    sleep(5)

    navegador.find_element_by_xpath('//*[@id="mat-tab-content-0-2"]/div/table/tbody/tr[2]/td[6]/div/i/span').click()
    sleep(15)

    navegador.find_element_by_xpath('//*[@id="mat-tab-content-0-2"]/div/mat-card/button[2]').click()
    sleep(2)

    pyautogui.click(x=618, y=187)
    sleep(5)

    navegador.close()

else:
    pass