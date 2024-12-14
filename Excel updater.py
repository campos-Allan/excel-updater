"""updating a spreadsheet data
pick updated data in two sheets and format into another sheet
"""
from datetime import datetime
from datetime import timedelta
import os
import win32com.client

values = {'285': {
    'P5_E': [0, 0]},
    '342': {
    'P6_R': [0, 0]},
    '340': {
    'P1_R': [0, 0],
    'P2_E': [0, 0]},
    '346': {
    'P3_E': [0, 0],
    'P4_E': [0, 0]}
}

# (datetime.now()-timedelta(days=30)).strftime('%y') fixed variable for sake of example
ano = 2024
meses = {'01': 'janeiro', '02': 'fevereiro', '03': 'março', '04': 'abril', '05': 'maio', '06': 'junho',
         '07': 'julho', '08': 'agosto', '09': 'setembro', '10': 'outubro', '11': 'novembro', '12': 'dezembro'}
mes = '11'  # (datetime.now()-timedelta(days=30)).strftime('%m')
dir_path = os.path.dirname(os.path.realpath(__file__))
FILE_1 = dir_path+'\\'
FILE_1 += [f for f in os.listdir() if (('.xlsx' in f)
                                       # file with planned values
                                       and (f'{ano}{mes}' in f))][0]
mes_escrito = meses[mes]
EXCEL_1 = win32com.client.Dispatch("EXCEL.Application")
EXCEL_1.Visible = False
PLANILHA_1 = EXCEL_1.Workbooks.Open(FILE_1)
SHEETS_1 = PLANILHA_1.Worksheets(['sheet1', 'sheet2'])
for sheet in SHEETS_1:
    if sheet.Name == 'sheet1':
        values['340']['P1_R'][0] = sheet.Cells(1, 2).Value
        values['340']['P2_E'][0] = sheet.Cells(2, 2).Value
        values['346']['P3_E'][0] = sheet.Cells(3, 2).Value
        values['346']['P4_E'][0] = sheet.Cells(4, 2).Value
    elif sheet.Name == 'sheet2':
        sheet = SHEETS_1[1]
        values['342']['P6_R'][0] = sheet.Cells(1, 2).Value
        values['285']['P5_E'][0] = sheet.Cells(2, 2).Value
PLANILHA_1.Close()
EXCEL_1.Quit()
# file with query to get realized values, and cross reference graphic
FILE_2 = dir_path+'\\sheet.xlsx'
EXCEL_2 = win32com.client.Dispatch("EXCEL.Application")
EXCEL_2.Visible = False
PLANILHA_2 = EXCEL_2.Workbooks.Open(FILE_2)
PLANILHA_2.RefreshAll()
EXCEL_2.CalculateUntilAsyncQueriesDone()
SHEETS_2 = PLANILHA_2.Worksheets(['sheet1', 'sheet2'])
for sheet in SHEETS_2:
    # commented lines would change the filter of the updated data
    # sheet.Cells(1, 2).Value = (datetime.now()-timedelta(days=30)).year
    # sheet.Cells(2, 2).Value = mes_escrito
    lastRow = sheet.UsedRange.Rows.Count
    for j in range(5, lastRow+1):
        try:
            ID = int(sheet.Cells(j, 1).Value)
        except ValueError:
            pass
        if len(str(ID)) == 3:
            TYPE = sheet.Cells(j, 3).Value
            E_R = sheet.Cells(j, 2).Value
            try:
                values[str(ID)][f'{TYPE}_{E_R}'][1] = sheet.Cells(j, 4).Value
            except KeyError as e:
                print(e)
PLANILHA_2 = EXCEL_2.Workbooks.Open(FILE_2)
SHEET_3 = PLANILHA_2.Worksheets(f'{ano}-{mes}')
for k in range(5, 11):
    ID = str(int(SHEET_3.Cells(k, 1).Value))
    if len(ID) == 3:
        TYPE = SHEET_3.Cells(k, 3).Value
        E_R = SHEET_3.Cells(k, 2).Value
        try:
            SHEET_3.Cells(k, 4).Value = values[ID][f'{TYPE}_{E_R}'][0]
            SHEET_3.Cells(k, 5).Value = values[ID][f'{TYPE}_{E_R}'][1]
        except KeyError as e:
            print(e)
PLANILHA_2.Save()
PLANILHA_2.Close()
EXCEL_2.Quit()
