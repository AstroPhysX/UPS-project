import PySimpleGUI as sg
import os
import sys
import pandas as pd

sg.theme('DarkAmber')
# Define the layout of the GUI
layout = [
    [sg.Text('Select PDF files to upload')],
    [sg.Text('Bid Lines File:'), sg.Input(key='file1'), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Text('Trips File:'), sg.Input(key='file2'), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Button('Upload'), sg.Button('Cancel')],
    [sg.Text('Terminal Output:')],
    [sg.Output(size=(100, 20), key='output')],
    [sg.Text('Select days of the week and dates:')],
    [sg.CalendarButton('Choose Date', target='input', key='date', format='%Y-%m-%d', button_color=('white', 'black')),
     sg.Input(key='input', enable_events=True, visible=False)],
    [sg.Listbox(values=[], size=(30, 6), key='selected_dates', select_mode='multiple')]
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
    elif event == 'date':
        window['input'].update(value=values['date'].strftime('%Y-%m-%d'))
        selected_date = pd.to_datetime(values['date']).strftime('%A, %B %d, %Y')
        selected_dates = window['selected_dates'].get_values()
        if selected_date not in selected_dates:
            selected_dates.append(selected_date)
        window['selected_dates'].update(values=selected_dates)