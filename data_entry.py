import PySimpleGUI as sg
import pandas as pd
from datetime import datetime
import os

# Frontend Application
# add color to the window
sg.theme('SandyBeach')
EXCEL_FILE_2 = "excel_file_2.xlsx"
# df = pd.read_excel(EXCEL_FILE_2)

layout = [
    # Heading
    [sg.Text('Please fill out the following fields:',
             justification='c', font='_ 15', expand_x=True)],
    # Serial Number Input
    [sg.Text('Sr.No', size=(15, 1)), sg.InputText(key='Sr.No')],
    # Name of the student
    [sg.Text('Name of Student', size=(15, 1)),
     sg.InputText(key='Name of Student')],
    # Roll number input
    [sg.Text('Roll No', size=(15, 1)), sg.InputText(key='Roll No')],
    # Year of joining
    [sg.Text('Year of Joining', size=(15, 1)),
     sg.InputText(key='Year of Joining')],
    # Category Input
    [sg.Text('Category', size=(15, 1)), sg.Combo(
        ['TA', 'TAP', 'MONASH', 'RA', 'CT', 'SW', 'SF'], key='Category')],
    # Department Input
    [sg.Text('Dept', size=(15, 1)), sg.Combo(['ChE', 'CHEMISTRY', 'PHYSICS',
                                              'MATH', 'IEOR', 'CTARA', 'CMIND', 'CSE', 'DESE', 'ESED'], key='Dept')],
    # Guide Name input
    [sg.Text("Guide's Name", size=(15, 1)), sg.InputText(key="Guide's Name")],
    # Date of Receiving Application
    [sg.Text('Date of Receiving Application', size=(15, 1)),
     sg.InputText(key='Date of Receiving Application')],
    # Days for processing of Application for ITCommittee
    [sg.Text('Days available for processing of Application for ITCommittee', size=(
        15, 1)), sg.InputText(key='Days available for processing of Application for ITCommittee')],
    # Name of the conference
    [sg.Text('Name of Conference', size=(15, 1)),
     sg.InputText(key='Name of Conference')],
    # Place of Conference
    [sg.Text('Place of Conference', size=(15, 1)),
     sg.InputText(key='Place of Conference')],
    # Region
    [sg.Text('Region', size=(15, 1)), sg.Combo(
        ['EUROPE', 'USA', 'ASIA'], key='Region')],
    # From date
    [sg.Input(key='From Date', size=(15, 1)), sg.CalendarButton(
        'From Date', format='%d:%m:%Y', close_when_date_chosen=True, target='From Date')],
    # To Date
    [sg.Input(key='To Date', size=(15, 1)), sg.CalendarButton(
        'To Date', format='%d:%m:%Y', close_when_date_chosen=True, target='To Date')],
    # Amount Eligible
    [sg.Text('Amount eligible', size=(15, 1)),
     sg.InputText(key='Amount eligible')],
    [sg.Text('Request for', size=(15, 1)), sg.InputText(key='Request for')],
    # Amount Requested
    [sg.Text('Amount Requested', size=(15, 1)),
     sg.InputText(key='Amount Requested')],
    # KM Clearance
    [sg.Text('KM Clearance', size=(15, 1)), sg.InputText(key='KM Clearance')],
    [sg.Text('IFC Clearance', size=(15, 1)),
     # IFC Clearance
     sg.InputText(key='IFC Clearance')],
    # Reason or Remark
    [sg.Text('Reason or Remark', size=(15, 1)),
     sg.InputText(key='Reason or Remark')],
    [sg.Submit(), sg.Exit()]
]
window = sg.Window(
    'Financial support for attending international conference', layout, size=(600, 600))
while True:
    event, values = window.read()
    # Data type of values is a dictionary
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        excel_file = pd.read_excel(EXCEL_FILE_2)
        roll_no = excel_file['Roll No'].map(int).unique()
        try:
            # Checking if the Roll Number is an Integers
            student_roll_no = int(values['Roll No'])
            student_exists = student_roll_no in roll_no
            # df = df.append(values, ignore_index=True)
            if not student_exists:
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
                sg.popup(
                    f'{values["Roll No"]} has been added to the records successfully!')
                for key in values:
                    window.find_element(key).Update('')
            else:
                sg.popup(
                    f'The roll number {student_roll_no} already exists in the record', 'Please Try Again!')
                window.find_element('Roll No').Update('')
        except:
            sg.popup('Roll Number should be an integer!')
            window.find_element('Roll No').Update('')

window.close()
