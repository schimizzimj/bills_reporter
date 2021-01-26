import PySimpleGUI as sg
from report_generator import main

def gui_main():
    sg.theme('DarkBlue3')

    layout = [
        [
            sg.Text("Please choose the input file: ", size=(26, 1), font=('Helvetica', 14)),
            sg.Input(size=(40,1), font=('Helvetica', 14)),
            sg.FileBrowse(key="BILLS", size=(10, 1), font=('Helvetica', 14))
        ],
        [sg.Submit(font=('Helvetica', 14)), sg.Cancel(font=('Helvetica', 14))]
    ]

    window = sg.Window('Report Generator', layout)

    event, values = window.read()
    window.close()

    text_input = values["BILLS"]

    main(text_input)
    sg.popup(f'The file {"report.docx"} has been generated.', font=('Helvetica', 14))

if __name__ == '__main__':
    gui_main()