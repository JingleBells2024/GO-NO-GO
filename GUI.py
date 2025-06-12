import PySimpleGUI as sg

layout = [
    [sg.Text('Financial Analysis', font=('Arial', 16))],
    [sg.Button('Upload Financial PDF/Excel'), sg.Button('Upload Excel Template')],
    [sg.Multiline(size=(50,5), key='-PROMPT-')],
    [sg.Button('Submit')]
]

window = sg.Window('Financial Analysis', layout)

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED,):
        break
    if event == 'Submit':
        prompt = values['-PROMPT-'].strip()
        print("Prompt:", prompt)

window.close()