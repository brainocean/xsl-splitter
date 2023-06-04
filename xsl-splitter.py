import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import PySimpleGUI as sg
import os

def get_column_list(df):
    return list(df.columns)

def get_unique_value_list(df, col):
    return list(df[col].unique())

def prepare_target_folder(path):
    if not os.path.isdir(path):
        os.makedirs(path)

def create_main_window():
    sg.theme('Dark Grey 13')
    layout = [[sg.Text('源数据文件')],
            [sg.Input(key='-source-filename-', enable_events=True, expand_x=True, disabled=True),
            sg.FileBrowse(button_text='打开', target='-source-filename-', file_types=(('Excel', '*.xls *.xlsx *.csv'),))],
            [sg.Text('选择用来分组的列')],
            [sg.Combo([], key='-column-combo-', expand_x=True, enable_events=True, readonly=True)],
            [sg.Output(expand_x=True, s=(20,30))],
            [sg.Text('目标文件夹')],
            [sg.Input(key='-target-folder-', expand_x=True, disabled=True),
            sg.FolderBrowse(button_text='浏览')],
            [sg.Button('执行', disabled=True), sg.Button('退出')],
            ]

    return sg.Window('数据文件分割器', layout, finalize=True)

def resetColumnCombo(window, cols):
    window['-column-combo-'].update(values=cols)
    window['执行'].update(disabled=True)

def main_loop(window):
    source_df = None
    source_path = ""

    while True:  # Event Loop
        event, values = window.read()

        if event == "-source-filename-":
            source_path = values['-source-filename-']
            print("读入" + source_path)
            if source_path.endswith('.csv'):
                source_df = pd.read_csv(source_path, index_col=None)
            else:
                source_df = pd.read_excel(source_path, index_col=None)
            resetColumnCombo(window, get_column_list(source_df))
            window["-target-folder-"].update(
                os.path.join(os.path.dirname(source_path), "target"))
            continue
        if event == '-column-combo-':
            if values['-column-combo-'] != None:
                window['执行'].update(disabled=False)
            continue
        if event == sg.WIN_CLOSED or event == '退出':
            window.close()
            break
        if event == '执行':
            target_folder = os.path.join(
                    values['-target-folder-'],
                    os.path.splitext(os.path.basename(source_path))[0])
            print("分割后文件写入文件夹", target_folder)
            prepare_target_folder(target_folder)
            col = values['-column-combo-']
            for key, grp in source_df.groupby(col):
                print("正在写入", key , ".xlsx文件.")
                grp.to_excel(os.path.join(target_folder, str(key) + '.xlsx'), index=False)
            print("文件分割完成.")

main_loop(create_main_window())
