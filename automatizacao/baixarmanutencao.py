from time import sleep
from selenium import webdriver
import pyautogui

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
