from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from Recursos import AcesseMeuSite
import os, pyautogui
from datetime import date, timedelta

user = os.path.expanduser('~')

conta_barreiro = 'VALLOUREC-GESTAO'
usuario = 'lucas.saraujo'
senha = 'Lucas@2022'

aux = 1

yesterday = date.today() - timedelta(days=1)
d1 = yesterday.strftime("%d/%m/%Y")

navegador = webdriver.Chrome(executable_path='\\\\srvoffice\\PL\\PLM\\Coord_Movimentação_Equipamentos\\4. Projetos\\2022\\Automações\\cd versions\\geral\\chromedriver.exe')
    
navegador.get('https://web.pdo-vallourec.com.br/#/')
    
AcesseMeuSite.Logar(navegador, conta_barreiro, usuario, senha)
sleep(30)

navegador.find_element_by_xpath('/html/body/app-root/ms-renderer/mat-sidenav-container/mat-sidenav[2]/ms-sidenav/div/div/mat-nav-list/ms-sidenav-item[3]/a/div').click()
sleep(2)

navegador.find_element_by_xpath('/html/body/app-root/ms-renderer/mat-sidenav-container/mat-sidenav[2]/ms-sidenav/div/div/mat-nav-list/ms-sidenav-item[3]/mat-nav-list/ms-sidenav-item[1]/a/div').click()
sleep(5)

navegador.find_element_by_xpath('//*[@id="mat-expansion-panel-header-8"]').click()
sleep(1)
    
navegador.find_element_by_xpath('//*[@id="mat-input-5"]').send_keys(aux, Keys.LEFT_CONTROL + 'a', Keys.DELETE)
sleep(1)
navegador.find_element_by_xpath('//*[@id="mat-input-5"]').send_keys(d1)
sleep(1)
    
navegador.find_element_by_xpath('//*[@id="mat-input-6"]').send_keys(aux, Keys.LEFT_CONTROL + 'a', Keys.DELETE)
sleep(1)
navegador.find_element_by_xpath('//*[@id="mat-input-6"]').send_keys(d1)
sleep(1)
   
navegador.find_element_by_xpath('//*[@id="mat-select-2"]/div/div[1]').click()
sleep(2)

navegador.find_element_by_xpath('//*[@id="mat-checkbox-2"]/label/div').click()
sleep(3)
pyautogui.click(x=647, y=407)
sleep(1)
navegador.find_element_by_xpath('//*[@id="mat-select-3"]/div/div[1]').click()
sleep(2)

navegador.find_element_by_xpath('//*[@id="mat-checkbox-3"]/label/div').click()
sleep(3)
pyautogui.click(x=647, y=407)
sleep(1)
navegador.find_element_by_xpath('//*[@id="mat-select-4"]/div/div[1]').click()
sleep(2)

navegador.find_element_by_xpath('//*[@id="mat-checkbox-4"]/label').click()
sleep(3)
pyautogui.click(x=647, y=407)
sleep(1)
navegador.find_element_by_xpath('//*[@id="cdk-accordion-child-8"]/div/pdo-form-content/mat-card-content/div[2]/pdo-button[2]/button').click()
sleep(15)
    
navegador.close()