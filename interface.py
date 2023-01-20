from tkinter.filedialog import askdirectory
from PySimpleGUI import PySimpleGUI as sg
import xml_to_excel

sg.theme('Reddit')
layout = [
    [sg.Button('Selecionar Pasta'),sg.Text('',key = 'diretorio')],
    [sg.Text('Nome da planilha'),sg.Input(key = 'planilha')],
    [sg.Button('Salvar'),sg.Text('',key = 'terminado')]
]


janela = sg.Window('Planilhador',layout)

while True:
    eventos, valores = janela.read()
    if eventos == sg.WIN_CLOSED:
        break
    if eventos == 'Selecionar Pasta':
        path = askdirectory(title='Selecione a pasta')
        janela.Element('diretorio').update(f'{path}')        
    if eventos == 'Salvar':
        xml_to_excel.xml_planilhar(path,valores['planilha'])
        janela.Element('terminado').update('Feito')
        