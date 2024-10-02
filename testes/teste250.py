import PySimpleGUI as sg

class TelaAutoJornada:

    def __init__(self):
        sg.theme('Reddit')
        layout = [
            [sg.Text('----- INFORMAÇÕES DE ACESSO -----', auto_size_text=True)],
            [sg.Text('Usuário:', size=(17, 0)), sg.Input(key='usuario', size=(20, 0))],
            [sg.Text('Senha:', size=(17, 0)), sg.Input(key='senha', password_char='*', size=(20, 0))],
            [sg.Text('Tempo:'), sg.InputCombo(['', ':'], size=(8, 1), key='-TEMPO-')],
            [sg.Text("Data Inicial:",size=(17,0)), sg.In(key="-CAL-", enable_events=True, font=('Arial', 11), size=(17,0)), sg.CalendarButton('Data', key='calendario', format=('%d/%m/%Y'), size=(6,0), font=('Arial', 10))],
            [sg.Text("Data Final:",size=(17,0)), sg.In(key="-CAL2-", enable_events=True, font=('Arial', 11), size=(17,0)), sg.CalendarButton('Data', key='calendario2', format=('%d/%m/%Y'), size=(6,0), font=('Arial', 10))],
            [sg.Text('----- HORÁRIO DOS TURNOS -----', auto_size_text=True)],
            [sg.Text('Turno 1 Fim', size=(17, 0)), sg.Input(key='t1fim', size=(20, 0))],
            [sg.Text('Turno 2 Inicio:', size=(17, 0)), sg.Input(key='t2inicio', size=(20, 0))],
            [sg.Text('Turno 2 Fim:', size=(17, 0)), sg.Input(key='t2fim', size=(20, 0))],
            [sg.Text('Turno 3 Inicio:', size=(17, 0)), sg.Input(key='t3inicio', size=(20, 0))],
            [sg.Text('12x36 Dia Inicio:', size=(17, 0)), sg.Input(key='12x36dinicio', size=(20, 0))],
            [sg.Text('12x36 Dia Fim:', size=(17, 0)), sg.Input(key='12x36dfim', size=(20, 0))],
            [sg.Text('12x36 Noite Inicio:', size=(17, 0)), sg.Input(key='12x36ninicio', size=(20, 0))],
            [sg.Text('12x36 Noite Fim:', size=(17, 0)), sg.Input(key='12x36nfim', size=(20, 0))],
            [sg.Text('----- ESCALAS -----', auto_size_text=True)],
            [sg.CB('6x1 e 6x2', font=('Arial', 12),key='6x1'), sg.CB('12x36', font=('Arial', 12),key='12x36')],
            [sg.Text('----- BASE DE DADOS -----', auto_size_text=True)],
            [sg.Text('Programação de Jornada.xlsx:', size=(22,0)), sg.In(size=(20,1) ,key='doc'), sg.FileBrowse(font=('Arial', 12))],
            [sg.Button('Programar', font=('Arial', 12)), sg.Quit(font=('Arial', 12))]]
        
        
        self.janela = sg.Window('Programação de Jornada', layout, resizable = True).Finalize()

        while True:

                self.button, self.values = self.janela.read()

                if self.button in (sg.WINDOW_CLOSED, "Quit"):
                    break
                elif self.button == "Programar":
                    if self.values["-CAL-"] and self.values["-CAL2-"] != '':
                         if self.values["usuario"] and self.values["senha"] != '':
                              if self.values["6x1"] or self.values["12x36"] != False:
                                    if self.values["doc"] != '':
                                        break
                                    else:
                                        sg.popup_error('Escolha uma base de dados!')
                              else:
                                sg.popup_error('Escolha uma escala!')
                         else:
                            sg.popup_error('Preencha as informações de acesso!')
                    else:
                         sg.popup_error('Preencha as datas!')

TelaAutoJornada()