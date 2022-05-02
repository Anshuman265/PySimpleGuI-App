import PySimpleGUI as sg
import pandas as pd
from datetime import datetime
import os

# add color to the window
sg.theme('darkteal9')
#EXCEL_FILE_2 = "excel_file_2.xlsx"
#df = pd.read_excel(EXCEL_FILE_2)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Sr.No', size=(15, 1)), sg.InputText(key='Sr.No')],
    [sg.Text('Name of Student', size=(15, 1)),
     sg.InputText(key='Name of Student')],
    [sg.Text('Roll No', size=(15, 1)), sg.InputText(key='Roll No')],
    [sg.Text('Year of Joining', size=(15, 1)),
     sg.InputText(key='Year of Joining')],
    [sg.Text('Category', size=(15, 1)), sg.Combo(
        ['RA', 'TA', 'TAP', 'COLTEA', 'SF', 'EX'], key='Category')],
    [sg.Text('Dept', size=(15, 1)), sg.Combo(['chE', 'CHEMISTRY', 'PHYSICS',
                                              'MATH', 'IEOR', 'CTARA', 'CMIND', 'CSE', 'DESE', 'ESED'], key='Dept')],
    [sg.Text("Guide's Name", size=(15, 1)), sg.InputText(key="Guide's Name")],
    [sg.Text('Date of Receiving Application', size=(15, 1)),
     sg.InputText(key='Date of Receiving Application')],
    [sg.Text('Days available for processing of Application for ITCommittee', size=(
        15, 1)), sg.InputText(key='Days available for processing of Application for ITCommittee')],
    [sg.Text('Name of Conference', size=(15, 1)),
     sg.InputText(key='Name of Conference')],
    [sg.Text('Place of Conference', size=(15, 1)),
     sg.InputText(key='Place of Conference')],
    [sg.Text('Region', size=(15, 1)), sg.Combo(
        ['EUROPE', 'USA', 'ASIA'], key='Region')],
    [sg.Input(key='From Date', size=(15, 1)), sg.CalendarButton(
        'From Date', format='%d:%m:%Y', close_when_date_chosen=True, target='From Date')],
    [sg.Input(key='To Date', size=(15, 1)), sg.CalendarButton(
        'To Date', format='%d:%m:%Y', close_when_date_chosen=True, target='To Date')],
    [sg.Text('Amount eligible', size=(15, 1)),
     sg.InputText(key='Amount eligible')],
    [sg.Text('Request for', size=(15, 1)), sg.InputText(key='Request for')],
    [sg.Text('Amount Requested', size=(15, 1)),
     sg.InputText(key='Amount Requested')],
    [sg.Text('KM Clearance', size=(15, 1)), sg.InputText(key='KM Clearance')],
    [sg.Text('IFC Clearance', size=(15, 1)),
     sg.InputText(key='IFC Clearance')],
    [sg.Text('Reason or Remark', size=(15, 1)),
     sg.InputText(key='Reason or Remark')],
    [sg.Submit(), sg.Exit()]
]
window = sg.Window(
    'Financial support for attending international conference', layout)
while True:
    event, values = window.read()
    # Data type of values is a dictionary
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        # df = df.append(values, ignore_index=True)
        df = pd.DataFrame()
        df = df.append(values, ignore_index=True)
        excelfilepath = os.getcwd()
        excelfilepath += '\\excel_file_2.xlsx'
        # print(df)
        writer = pd.ExcelWriter(
            excelfilepath, engine='openpyxl', mode='a', if_sheet_exists='overlay')
        df.to_excel(writer, sheet_name="excel_file_2.xlsx",
                    startrow=writer.sheets['excel_file_2.xlsx'].max_row, header=None, index=False)
        writer.save()
        for key in values:
            window.find_element(key).Update('')
window.close()
