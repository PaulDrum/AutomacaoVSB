import PySimpleGUI as sg, os

user = os.path.expanduser('~')

class TelaRelTurno:

    def __init__(self):
        usuarios = [
         'Cleiton Rocha', 'Ronaldo Santos', 'ADMIN']
        user_win = ['C:\\Users\\Cleiton.Rocha', 'C:\\Users\\Ronaldo.SSantos', 'C:\\Users\\s-Lucas.SAraujo']
        user = os.path.expanduser('~')
        usuario = ''
        for i in range(0, 2):
            if user == user_win[i]:
                usuario = usuarios[i]

        turnos = [1, 2, 3]

        sg.theme('Reddit')
        column1 = [
            [sg.Text('Responsável:', size=(15, 0), key="texto1", font=('Arial', 15)), sg.Combo(usuarios, default_value=usuario, key='responsavel', size=(22, 0), expand_x=True, font=('Arial', 15))],
            [sg.Text('Data:', size=(15,0), font=('Arial', 15)), sg.In(key='-CAL-', enable_events=True, visible=False, font=('Arial', 15)), sg.CalendarButton('Calendário', key='calendario', format=('%d/%m/%Y'), size=(20,0), font=('Arial', 15))],
            [sg.Text('Turno:', size=(15, 0), font=('Arial', 15)), sg.Combo(turnos, key='turno', size=(22, 0),font=('Arial', 15))],
            [sg.Text('Informações Gerais:', size=(15, 0), font=('Arial', 15)), sg.Multiline(key='infogeral', size=(22, 4), font=('Arial', 15))],
            [sg.Text('Desvios (Segurança):', size=(15, 0), font=('Arial', 15)), sg.Multiline(key='desvio_seg', size=(22, 4), font=('Arial', 15))],
            [sg.Text('Anexo (Opcional):', size=(15,0), font=('Arial', 15)), sg.In(size=(15,1) ,key='anexo_seg_1', font=('Arial', 15)), sg.FileBrowse(font=('Arial', 15)), sg.Button(" + ", key="botao1", font=('Arial', 15))],
            [sg.Text('Anexo (Opcional):', size=(15,0), key="anexo2", visible=False, font=('Arial', 15)), sg.In(size=(15,1) ,key='anexo_seg_2', visible=False, font=('Arial', 15)), sg.FileBrowse(key="anexo22", visible=False, font=('Arial', 15)), sg.Button(" + ", key="botao2", visible=False, font=('Arial', 15))],
            [sg.Text('Anexo (Opcional):', size=(15,0), key="anexo3", visible=False, font=('Arial', 15)), sg.In(size=(15,1) ,key='anexo_seg_3', visible=False, font=('Arial', 15)), sg.FileBrowse(key="anexo222", visible=False, font=('Arial', 15))],
            [sg.Text('Desvios (Meio-Ambiente):', size=(15, 0), font=('Arial', 15)), sg.Multiline(key='desvio_meio', size=(22, 4), font=('Arial', 15))],
            [sg.Text('Anexo (Opcional):', size=(15,0), font=('Arial', 15)), sg.In(size=(15,1) ,key='anexo_meio_1' ,font=('Arial', 15)), sg.FileBrowse(font=('Arial', 15)), sg.Button(" + ", key="botao3", font=('Arial', 15))],
            [sg.Text('Anexo (Opcional):', size=(15,0), key="anexo4", visible=False, font=('Arial', 15)), sg.In(size=(15,1) ,key='anexo_meio_2', visible=False, font=('Arial', 15)), sg.FileBrowse(key="anexo33", visible=False, font=('Arial', 15)), sg.Button(" + ", key="botao4", visible=False, font=('Arial', 15))],
            [sg.Text('Anexo (Opcional):', size=(15,0), key="anexo5", visible=False, font=('Arial', 15)), sg.In(size=(15,1) ,key='anexo_meio_3', visible=False, font=('Arial', 15)), sg.FileBrowse(key="anexo333", visible=False, font=('Arial', 15))],
            [sg.Text('Indisponibilidades:', size=(15, 0), font=('Arial', 15)), sg.Multiline(key='indisp', size=(22, 4) ,font=('Arial', 15))],
            [sg.Text('Equipamentos no Canteiro:', size=(15,0), font=('Arial', 15)), sg.In(size=(15,1) ,key='equip_canteiro', font=('Arial', 15)), sg.FileBrowse(font=('Arial', 15))],
            [sg.Text('Ver e Agir:', size=(15,0), font=('Arial', 15)), sg.In(size=(15,1) ,key='vereagir', font=('Arial', 15)), sg.FileBrowse(font=('Arial', 15))],
            [sg.Button("Gerar Relatório" ,font=('Arial', 15)), sg.Quit(font=('Arial', 15))]
        ]

        layout = [
            [
            sg.Column(column1, scrollable=True, vertical_scroll_only = True, expand_x=True, expand_y=True)
            ]
        ]
        self.janela = sg.Window('Relatório de Turno', layout, resizable = True).Finalize()

        while True:

            self.button, self.values = self.janela.read()

            if self.button in (sg.WINDOW_CLOSED, "Quit"):
                break
            elif self.button == "Gerar Relatório":
                text = self.values['responsavel']
                if text == '':
                    sg.popup_error('Preencha todos os Campos!')
                else:
                    break
            if self.button == "botao1":
                self.janela['botao1'].update(visible=False)
                self.janela['anexo2'].update(visible=True)
                self.janela['anexo_seg_2'].update(visible=True)
                self.janela['anexo22'].update(visible=True)
                self.janela['botao2'].update(visible=True)
            if self.button == "botao2":
                self.janela['botao2'].update(visible=False)
                self.janela['anexo3'].update(visible=True)
                self.janela['anexo_seg_3'].update(visible=True)
                self.janela['anexo222'].update(visible=True)
            if self.button == "botao3":
                self.janela['botao3'].update(visible=False)
                self.janela['anexo4'].update(visible=True)
                self.janela['anexo_meio_2'].update(visible=True)
                self.janela['anexo33'].update(visible=True)
                self.janela['botao4'].update(visible=True)
            if self.button == "botao4":
                self.janela['botao4'].update(visible=False)
                self.janela['anexo5'].update(visible=True)
                self.janela['anexo_meio_3'].update(visible=True)
                self.janela['anexo333'].update(visible=True)

    def Iniciar(self):         
        self.janela.close()

