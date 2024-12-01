"""updating a spreadsheet data
pick updated data in two sheets and format into another sheet
"""
from datetime import datetime
from datetime import timedelta
from os import listdir
import win32com.client

values = {'285': {
    'DSL_E': [0, 0],
    'GLN_E': [0, 0]},
    '342': {
    'GLN_E': [0, 0]},
    '340': {
    'DSL_E': [0, 0],
    'DSL_R': [0, 0],
    'GLN_E': [0, 0],
    'NFT_R': [0, 0],
    'NFT_E': [0, 0],
    'QRS_R': [0, 0],
    'QRS_E': [0, 0]},
    '346': {
    'DSL_E': [0, 0],
    'GLN_E': [0, 0],
    'OC_E': [0, 0],
    'GLP_R': [0, 0],
    'GOL_R': [0, 0]}
}

ano = (datetime.now()-timedelta(days=30)).strftime('%y')
meses = {'01': 'janeiro', '02': 'fevereiro', '03': 'março', '04': 'abril', '05': 'maio', '06': 'junho',
         '07': 'julho', '08': 'agosto', '09': 'setembro', '10': 'outubro', '11': 'novembro', '12': 'dezembro'}
mes = (datetime.now()-timedelta(days=30)).strftime('%m')
FILE_1 = 'path/'
FILE_1 += [f for f in listdir() if (('.xlsx' in f)
                                    and (f'{ano}{mes}' in f))][0]
mes_escrito = meses[mes]
EXCEL_1 = win32com.client.Dispatch("EXCEL.Application")
EXCEL_1.Visible = False
PLANILHA_1 = EXCEL_1.Workbooks.Open(FILE_1)
SHEETS_1 = PLANILHA_1.WorkSHEETs(['sheetName1', 'sheetName2'])
for sheet in SHEETS_1:
    if sheet.Name == 'sheetName1':
        values['340']['P1_R'][0] = format(sheet.Cells(19, 2).Value, '.0f')
        values['340']['P2_E'][0] = format(sheet.Cells(
            20, 2).Value+sheet.Cells(21, 2).Value, '.0f')
        values['340']['P3_R'][0] = format(sheet.Cells(22, 2).Value, '.0f')
        values['340']['P4_E'][0] = format(sheet.Cells(
            24, 2).Value+sheet.Cells(25, 2).Value, '.0f')
        values['340']['P5_R'][0] = format(sheet.Cells(23, 2).Value, '.0f')
        values['346']['P6_E'][0] = format(sheet.Cells(
            34, 2).Value+sheet.Cells(35, 2).Value+sheet.Cells(36, 2).Value, '.0f')
        values['346']['P7_E'][0] = format(sheet.Cells(
            32, 2).Value+sheet.Cells(33, 2).Value, '.0f')
        values['346']['P8_E'][0] = format(sheet.Cells(
            37, 2).Value+sheet.Cells(38, 2).Value, '.0f')
        values['346']['P9_R'][0] = format(sheet.Cells(
            30, 2).Value+sheet.Cells(31, 2).Value, '.0f')
    elif sheet.Name == 'sheetName2':
        sheet = SHEETS_1[1]
        values['342']['P10_E'][0] = format(sheet.Cells(
            24, 3).Value+sheet.Cells(26, 3).Value, '.0f')
        values['285']['P11_E'][0] = format(sheet.Cells(28, 3).Value+sheet.Cells(29, 3).Value-(
            sheet.Cells(24, 3).Value+sheet.Cells(26, 3).Value), '.0f')
        values['285']['P12_E'][0] = format(sheet.Cells(
            27, 3).Value+sheet.Cells(31, 3).Value, '.0f')
PLANILHA_1.Close()
EXCEL_1.Quit()
FILE_2 = 'path/sheet.xlsx'
EXCEL_2 = win32com.client.Dispatch("EXCEL.Application")
EXCEL_2.Visible = False
PLANILHA_2 = EXCEL_2.Workbooks.Open(FILE_2)
PLANILHA_2.RefreshAll()
EXCEL_2.CalculateUntilAsyncQueriesDone()
SHEETS_2 = PLANILHA_2.WorkSHEETs(['sheetName1', 'sheetName2'])
for sheet in SHEETS_2:
    sheet.Cells(1, 2).Value = (datetime.now()-timedelta(days=30)).year
    sheet.Cells(2, 2).Value = mes_escrito
    lastRow = sheet.UsedRange.Rows.Count
    for j in range(5, lastRow+1):
        try:
            ID = int(sheet.Cells(j, 1).Value)
        except ValueError:
            pass
        if len(str(ID)) == 3:
            TYPE = sheet.Cells(j, 3).Value
            E_R = sheet.Cells(j, 2).Value
            if str(ID) == '340' and TYPE == 'P1':
                pass
            else:
                try:
                    values[str(ID)][f'{TYPE}_{E_R}'][1] = format(
                        sheet.Cells(j, 4).Value, '.0f')
                except KeyError as e:
                    print(e)
PLANILHA_2 = EXCEL_2.Workbooks.Open(FILE_2)
ano_4 = (datetime.now()-timedelta(days=30)).year
SHEET_3 = PLANILHA_2.WorkSHEETs(f'{ano_4}-{mes}')
for k in range(4, 19):
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
