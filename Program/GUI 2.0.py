import PySimpleGUI as sg
import os
import sys
from Final import *

sg.theme('DarkAmber')
# Define the layout of the GUI
layout = [
    [sg.Text('Select PDF files to upload')],
    [sg.Text('Bid Lines File:'), sg.Input(key='file1'), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Text('Trips File:'), sg.Input(key='file2'), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Button('Upload'), sg.Button('Cancel')],
    [sg.Text('Terminal Output:')],
    [sg.Output(size=(100, 20), key='output')]
]

# Create the GUI window
window = sg.Window('UPS Line Sorting', layout)

# Loop through events in the GUI
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        break
    elif event == 'Upload':
        file1 = values['file1']
        file2 = values['file2']
        # Check if selected files are PDFs
        if not all([file1.endswith('.pdf'), file2.endswith('.pdf')]):
            sg.popup('Please select two PDF files')
        else:
            # Do something with the selected files
            print(f'File 1: {file1}')
            print(f'File 2: {file2}')
            sg.popup('Files uploaded successfully!')
            program(file1, file2)
            
            print(file1)
            print(file2)
            # Redirect output to the GUI terminal
            with sg.Output(output_tk=window['output'].TKOut):
                print("were in")
                sys.stdout = window['output'].TKOut
                program(file1, file2)
                sys.stdout = sys.__stdout__

window.close()