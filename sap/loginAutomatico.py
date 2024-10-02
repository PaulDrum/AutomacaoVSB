import win32com.client
import sys
import subprocess
import time
from tkinter import *
from datetime import date, timedelta

class SapGui(object):
    def __init__(self):
        yesterday = date.today() - timedelta(days=1)
        self.d1 = yesterday.strftime("%d.%m.%Y")
        self.archive_dest = r"C:\Users\s-Lucas.SAraujo\ZBWINPL007_GESTÃO_DOS_ESTOQUES_20220623.csv"


        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(5)
        

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")       

        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection("3 SAP BW/BPS", True) #nome do sistema / modulo
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize

    def SapLogin(self):

        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "100" #Nº mandante aqui
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "73926" #Usuario aqui
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "BRAu0006" #Senha aqui
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT" #Idioma
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(3)
            self.session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "0000000025"
            self.session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("0000000025")
          

        except:
            print(sys.exc_info()[0])

    def SapGetBAR(self):

        centro1 = "7501"
        centro2 = "7601"

        self.session.findById("wnd[0]/usr/ctxtP_DATE").text = self.d1
        self.session.findById("wnd[0]/usr/ctxtS_PLANT-LOW").text = centro1
        self.session.findById("wnd[0]/usr/ctxtS_PLANT-HIGH").text = centro2
        self.session.findById("wnd[0]/usr/txtP_FILE").text = self.archive_dest
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(15)
    
    
    def SapGetJCB(self):

        centro1 = "vs01"
        centro2 = ""

        self.session.findById("wnd[0]/usr/ctxtP_DATE").text = self.d1
        self.session.findById("wnd[0]/usr/ctxtS_PLANT-LOW").text = centro1
        self.session.findById("wnd[0]/usr/ctxtS_PLANT-HIGH").text = centro2
        self.session.findById("wnd[0]/usr/txtP_FILE").text = self.archive_dest
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(15)

    def SAPQuit(self):
        self.connection.CloseSession('ses[0]')


# if __name__ == '__main__':
#     window = Tk()
#     window.geometry("200x50")
#     botao = Button(window, text="Login SAP", command= lambda: SapGui().SapLogin())
#     botao.pack()
#     mainloop()

