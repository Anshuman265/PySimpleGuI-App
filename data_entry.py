import PySimpleGUI as sg
import pandas as pd
from datetime import datetime

# add color to the window
sg.theme('darkteal9')
EXCEL_FILE_1 = "Book2.xlsx"
EXCEL_FILE_2 = "Book1.xlsx"
df = pd.read_excel(EXCEL_FILE_1)
df = pd.read_excel(EXCEL_FILE_2)
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
    print(values['Sr.No'])
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        #df.drop('No of years',axis=1,inplace=True)
        #df.drop('No of Days',axis=1,inplace=True)
        print(df)
        # if df.values['KM Clearance'=='Yes']
        df.to_excel(EXCEL_FILE_2, index=False)
        #df.drop('Name of Student',axis=1,inplace=True)
        #df.drop('Roll No',axis=1,inplace=True)
        #df.drop('Year of Joining',axis=1,inplace=True)
        # df.drop('Category',axis=1,inplace=True)
        # df.drop('Dept',axis=1,inplace=True)
        #df.drop("Guide's Name",axis=1,inplace=True)
        #df.drop('Date of Receiving Application',axis=1,inplace=True)
        #df.drop('Days available for processing of Application for ITCommittee',axis=1,inplace=True)
        #df.drop('Name of Conference',axis=1,inplace=True)
        # df.drop('Region',axis=1,inplace=True)
        #df.drop('From Date',axis=1,inplace=True)
        #df.drop('To Date',axis=1,inplace=True)
        #df.drop('Amount eligible',axis=1,inplace=True)
        #df.drop('Request for',axis=1,inplace=True)
        #df.drop('Amount Requested',axis=1,inplace=True)
        #df.drop('KM Clearance',axis=1,inplace=True)
        #df.drop('IFC Clearance',axis=1,inplace=True)
        #df.drop('Reason or Remark',axis=1,inplace=True)
        #df.drop('Place of Conference',axis=1,inplace=True)
        #a2=df.at['To Date']
        # #a=datetime(int(a1))
        # b=datetime(int(a2))
        # datetime.timedelta(7)
        # d=(b-a).days
        #df=df.append(d,key='No of Days',ignore_index=True)
        window.find_element('Sr.No').Update('')
        window.find_element('Name of Student').Update('')
        window.find_element('Roll No').Update('')
        window.find_element('Year of Joining').Update('')
        window.find_element('Category').Update('')
        window.find_element('Dept').Update('')
        window.find_element("Guide's Name").Update('')
        window.find_element('Date of Receiving Application').Update('')
        window.find_element(
            'Days available for processing of Application for ITCommittee').Update('')
        window.find_element('Name of Conference').Update('')
        window.find_element('Place of Conference').Update('')
        window.find_element('Region').Update('')
        window.find_element('From Date').Update('')
        window.find_element('To Date').Update('')
        window.find_element('Amount eligible').Update('')
        window.find_element('Request for').Update('')
        window.find_element('Amount Requested').Update('')
        window.find_element('KM Clearance').Update('')
        window.find_element('IFC Clearance').Update('')
        window.find_element('Reason or Remark').Update('')


window.close()
