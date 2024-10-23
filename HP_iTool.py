import pandas as pd
import numpy as np
import datetime
import PySimpleGUI as sg
import sys
import os
from pathlib import Path

pd.options.mode.copy_on_write = False

def Check(window, mg_input_1, mg_input_2):
    if mg_input_1 == '' or mg_input_2 == '':
        window['-OUT2-'].update('Remember to select the files.')
    else:
        df1 = pd.read_excel(mg_input_1, keep_default_na=False, engine='openpyxl')
        df2 = pd.read_excel(mg_input_2, keep_default_na=False, engine='openpyxl')

        headerCheck = len(df2.columns.intersection(df1.columns)) == len(df2.columns) # 看要以哪個df為基準，這裡是以df2
        if headerCheck == True: # 兩個表格的header有不同
            window['-OUT2-'].update('Header of two excels are the same :)')
        else:
            ODM = []
            Pulsar = []
            for col in df1.columns.difference(df2.columns):
                ODM.append(col)
            for col in df2.columns.difference(df1.columns):
                Pulsar.append(col)
            window['-OUT2-'].update(f'Header of two excels are different!, ODM: {ODM}, Pulsar: {Pulsar}')

def add_NA_FreeDOS(row):
    if 'FreeDOS' in row['OS']: # 有FreeDOS就換N/A
        row['Image Size (GB)'] = np.nan # Image Size在Pulsar不接受str的資料格式，另外改值
        row[fill_list] = 'N/A'
    return row

def add_NA(row):
    if 'FreeDOS' in row['OS'] and (row[fill_list]=='').all(): # 有FreeDOS且fill_list欄位皆為空
        row['Image Size (GB)'] = np.nan # Image Size在Pulsar不接受str的資料格式，另外改值
        row[fill_list] = 'N/A'    
    return row

def Merge(window, mg_input_1, mg_input_2, freedos, mg_output):
    if mg_input_1 == '' or mg_input_2 == '':
        window['-OUT2-'].update('Remember to select the files.')
    else:
        df1 = pd.read_excel(mg_input_1, keep_default_na=False, engine='openpyxl') # ODM
        df2 = pd.read_excel(mg_input_2, keep_default_na=False, engine='openpyxl') # Pulsarss
        outputfile = str(Path(mg_output)) + '\ssrm_'+datetime.datetime.now().strftime('%y%m%d')+'.xlsx'

        df1['RTM'].replace('', np.NaN, regex=True, inplace=True)
        df1['EOL Date'].replace('', np.NaN, regex=True, inplace=True)

        df_new = df2.set_index('Image ID')
        df_new.update(df1.set_index('Image ID'))
        df_new.reset_index(inplace=True)

        for i in range(len(df_new['Proposed RTM'])):
            if type(df_new['Proposed RTM'][i]) is datetime.datetime:
                df_new['Proposed RTM'][i] = df_new['Proposed RTM'][i].strftime("%m/%d/%Y")
            if type(df_new['Actual RTM'][i]) is datetime.datetime:
                df_new['Actual RTM'][i] = df_new['Actual RTM'][i].strftime("%m/%d/%Y")
                
        if freedos == True:
            df_new = df_new.apply(add_NA_FreeDOS, axis=1)
        else:
            df_new = df_new.apply(add_NA, axis=1)

        def highlight_difference(s): # highlight cells
            is_difference = s != df2[s.name].values
            return ['background-color: lightyellow' if v else '' for v in is_difference]
        
        df_new.style.apply(highlight_difference).to_excel(outputfile, index=False)
        window['-OUT2-'].update('Hi, your files have been merged to the output folder:)')

def Modification(window, mo_input_1, mo_input_2, mo_input_3):
    if mo_input_1 == '' or mo_input_2 == '' or mo_input_3 == '':
        window['-OUT2-'].update('Remember to select the files.')
    else:
        df = pd.read_excel(mo_input_1, keep_default_na=False, engine='openpyxl')
        df.rename(columns={mo_input_2: mo_input_3}, inplace=True)
        df.to_excel('output_ODM.xlsx', index=False)
        window['-OUT2-'].update('Modification is done.')

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def main():
    # 設置主題及icon
    sg.theme("LightBlue")
    sg.set_global_icon(icon=resource_path('hpicon.ico'))

    layout = [[sg.T("")],
              [sg.Text("Choose a file from ODM:", s=20), sg.Input(key="-IN1-"), sg.FileBrowse(file_types=(("Excel Files", ".xls*"), )), sg.Checkbox('FreeDOS', default=True, key="-IN3-")],
              [sg.Text("Choose a file from Pulsar:", s=20), sg.Input(key="-IN2-"), sg.FileBrowse(file_types=(("Excel Files", ".xls*"), ))],
              [sg.Text("Output Folder:", s=20), sg.Input(key="-OUT1-"), sg.FolderBrowse()],
              [sg.T("")], 
              # [sg.Button("Check"), sg.Button("Modify"), sg.Button("Merge"), sg.Button('Clear'), sg.Button('Exit')],
              [sg.Button("Check"), sg.Button("Merge"), sg.Button('Clear'), sg.Button('Exit')],
              [sg.Text(text_color='red', key='-OUT2-')]]

    # Building Window
    window = sg.Window('HP iTool - Release v1.1.0', layout, size=(670,210), keep_on_top=True, finalize=True)

    # Event Loop
    while True:
        event, values = window.read()
        
        if event in (sg.WIN_CLOSED, "Exit"):
            print('Closing the window!')
            break
        if event == 'Check':
            Check(window,
                  mg_input_1=values["-IN1-"],
                  mg_input_2=values["-IN2-"],)
        # if event == 'Modify':
        #     layout_clicked = [[sg.Text("Choose a file you want to modify:"), sg.Input(key="-IN4-", s=32), sg.FileBrowse(file_types=(("Excel Files", ".xls*"), ))],
        #                       [sg.Text("Orginal header:"), sg.Input(key="-IN5-", s=15), sg.Text("Modified header:"), sg.Input(key="-IN6-", s=15)],
        #                       [sg.Ok(), sg.Cancel()]]
        #     clicked, values_clicked = sg.Window('Modification', layout_clicked, disable_close=True, size=(520,100), keep_on_top=True).read(close=True)
        #     if clicked == 'Ok':
        #         Modification(window,
        #                      mo_input_1=values_clicked["-IN4-"],
        #                      mo_input_2=values_clicked["-IN5-"],
        #                      mo_input_3=values_clicked["-IN6-"])
        #         # window['-OUT2-'].update('Modification is done.')
        #     if clicked == 'Cancel':
        #         window['-OUT2-'].update('Not doing the modification...')
        if event == 'Clear':  # clear keys if clear button
            window['-IN1-'].update('')
            window['-IN2-'].update('')
            window['-OUT1-'].update('')
            window['-OUT2-'].update('Clear the inputs...')
        if event == "Merge":
            Merge(window,
                  mg_input_1=values["-IN1-"],
                  mg_input_2=values["-IN2-"],
                  freedos=values["-IN3-"],
                  mg_output=values["-OUT1-"],)
    window.close()    

if __name__ == '__main__':
    fill_list = ['column_name_that_fill_na']
    
    main()