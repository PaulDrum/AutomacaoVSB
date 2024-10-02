import PySimpleGUI as sg
from variaveis_dnf import Variaveis




class TelaDnf:

    def __init__(self):
        sg.theme('DarkTeal9')

        layout = [
            [sg.Text('NAVIO:',size=(15, 0)), sg.Input(key='transporte', size=(15,0))],

            [sg.Text('Planejamento:',size=(15, 0)), sg.Input(key='plj', size=(15,0))],

            [sg.Text('Armador:',size=(15, 0)), sg.Combo([Variaveis.armador_agente[0], Variaveis.armador_agente[1], Variaveis.armador_agente[2]], key='armador', size=(15, 6))],

            [sg.Text('Porto de Embarque:',size=(15, 0)), sg.Combo([Variaveis.port_embarque[0]], key='port_embarque', size=(15, 6))],
            
            [sg.Text('Porto de Destino:',size=(15, 0)), sg.Combo([Variaveis.port_dest[0],
            Variaveis.port_dest[1],Variaveis.port_dest[2],Variaveis.port_dest[3],Variaveis.port_dest[4],Variaveis.port_dest[5],Variaveis.port_dest[6],Variaveis.port_dest[7],
            Variaveis.port_dest[8],Variaveis.port_dest[9],Variaveis.port_dest[10],Variaveis.port_dest[11],Variaveis.port_dest[12],Variaveis.port_dest[13],Variaveis.port_dest[14],
            Variaveis.port_dest[15],Variaveis.port_dest[16],Variaveis.port_dest[17],Variaveis.port_dest[18],Variaveis.port_dest[19],Variaveis.port_dest[20],Variaveis.port_dest[21],
            Variaveis.port_dest[22],Variaveis.port_dest[23],Variaveis.port_dest[24],Variaveis.port_dest[25],Variaveis.port_dest[26],Variaveis.port_dest[27],Variaveis.port_dest[28]], key='port_destino', size=(15, 6))],
            
            [sg.Text('Incoterm:',size=(15, 0)), sg.Combo([Variaveis.incoterm[0], 
            Variaveis.incoterm[1], Variaveis.incoterm[2], Variaveis.incoterm[3], Variaveis.incoterm[4], Variaveis.incoterm[5], Variaveis.incoterm[6], Variaveis.incoterm[7], Variaveis.incoterm[8], Variaveis.incoterm[9], Variaveis.incoterm[10], Variaveis.incoterm[11], Variaveis.incoterm[12], Variaveis.incoterm[13]
            ], key='incoterm', size=(15, 6))],

            [sg.Text('Data BL:', size=(15,0)), sg.In(key='-CAL-', enable_events=True, visible=False), sg.CalendarButton('Calendário', key='calendario', format=('%d/%m/%Y'))],

            [sg.Text('TAXA USD:',size=(15,0)), sg.Input(key="IN1", size=(15,0))],

            [sg.Text('Bruto:',size=(15,0)), sg.Input(key='bruto', size=(15,0))],

            [sg.Text('DUE:',size=(15,0)), sg.Input(key='due', size=(15,0))],

            [sg.Text('BL:',size=(15,0)), sg.Input(key='bl', size=(15,0))],

            [sg.Text('N de Porões:',size=(15,0)), sg.Input(key='porao', size=(15,0))],

            [sg.Text('Analista:',size=(15, 0)), sg.Combo([Variaveis.analista[0], Variaveis.analista[1], Variaveis.analista[2], Variaveis.analista[3], Variaveis.analista[4], Variaveis.analista[5]], key='analista', size=(15, 6))],

            [sg.Text('FECOMERCIO:',size=(15,0)), sg.Input(key='fecomercio', size=(15,0))],

            [sg.Text('Modal:',size=(15, 0)), sg.Combo([Variaveis.modal[0]], key='modal', size=(15, 6))],

            [sg.Text('Estiva:',size=(15, 0)), sg.Combo([Variaveis.estiva[0], Variaveis.estiva[1], Variaveis.estiva[2], Variaveis.estiva[3], Variaveis.estiva[4]], key='estiva', size=(15, 6))],

            [sg.Text('Usina:',size=(15, 0)), sg.Combo([Variaveis.usina[0], Variaveis.usina[1], Variaveis.usina[2], Variaveis.usina[3], Variaveis.usina[4]], key='usina', size=(15, 6))],

            [sg.Text('Armazenagem:',size=(15, 0)), sg.Combo([Variaveis.armazenagem[0], Variaveis.armazenagem[1], Variaveis.armazenagem[2], Variaveis.armazenagem[3]], key='armazenagem', size=(15, 6))],

            [sg.Text('FEDEX / DHL:',size=(15, 0)), sg.Input(key='fedex', size=(15,0))],

            [sg.Text('Frete / Guerra:',size=(15,0)), sg.Input(key='taxa', size=(15,0))],

            [sg.Text('Observação:', size=(15,0)), sg.Input(key='observacao', size=(15,6))],

            [sg.Text('Folder', size=(15,0)), sg.In(size=(25,1) ,key='-FOLDER-'), sg.FolderBrowse()],

            [sg.Button("Criar DNF"), sg.Quit()]
        ]
        self.janela = sg.Window('Gerar DNF', layout)
        
        while True:

            self.button, self.values = self.janela.read()

            if self.button in (sg.WINDOW_CLOSED, "Quit"):
                break
            elif self.button == "Criar DNF":
                text = self.values['IN1']
                if text == '':
                    sg.popup_error('Campo Vazio')
                else:
                    try:
                        value = float(text)
                        break
                    except:
                        sg.popup_error('Digite um valor numérico e utilize ponto em vez de vírgula.')

    def Iniciar(self):
        transporte = self.values['transporte']        
        armador = self.values['armador']
        port_embarque = self.values['port_embarque']
        port_destino = self.values['port_destino']
        incoterm = self.values['incoterm']
        calendario = self.values['-CAL-']
        taxa_usd = self.values['IN1']
        bruto = self.values['bruto']
        due = self.values['due']
        bl = self.values['bl']
        porao = self.values['porao']
        analista = self.values['analista']
        fecomercio = self.values['fecomercio']
        modal = self.values['modal']
        estiva = self.values['estiva']
        usina = self.values['usina']
        armazenagem = self.values['armazenagem']
        fedex = self.values['fedex']
        caixa_texto = self.values['observacao']
        frete_taxa_de_guerra = self.values['taxa']        
        plj = self.values['plj']
        folder = self.values['-FOLDER-']
                       
        self.janela.close()



# tela = TelaDnf()
# tela.Iniciar()







