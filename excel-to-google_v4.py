from sheet_service_v2 import Create_Service
import os.path
import win32com.client as win32

#source excel file path
excel_path = 'C:\\Users\\usenadu\\Documents\\dumidu-pearson\\google api\\1\\excel_files\\'

f = open ("excel.txt", "r")
f1 = f.readlines()

g = open ("gsheet.txt", "r")
g1 = g.readlines()

for excel_file, gsheet in zip(f1, g1):
    excel_file = excel_file.rstrip()
    gsheet = gsheet.rstrip()

    excel_file_path = os.path.join(excel_path, excel_file)

    xlApp = win32.Dispatch('Excel.Application')
    wb = xlApp.Workbooks.Open(excel_file_path)      # wb --> workbook
    ws1 = wb.Worksheets('tab1')                     # ws --> worksheet
    rngData1 = ws1.Range('A1').CurrentRegion()
    ws2 = wb.Worksheets('tab2')
    rngData2 = ws2.Range('A1').CurrentRegion()


    # define Google Sheet and api
    gsheet_id = gsheet
    CLIENT_SECRET_FILE = 'client_secret.json'
    API_SERVICE_NAME = 'sheets'
    API_VERSION = 'v4'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)


    # code for excel tab 1
    response = service.spreadsheets().values().clear(
        spreadsheetId=gsheet_id,
        range='tab1!$A:$ZZ',
    ).execute()

    response = service.spreadsheets().values().append(
        spreadsheetId=gsheet_id,
        valueInputOption='RAW',
        range='tab1!A1',
        body=dict(
            majorDimension='ROWS',
            values=rngData1
        )
    ).execute()


    # code for excel tab 2
    response = service.spreadsheets().values().clear(
        spreadsheetId=gsheet_id,
        range='tab2!$A:$ZZ',
    ).execute()

    response = service.spreadsheets().values().append(
        spreadsheetId=gsheet_id,
        valueInputOption='RAW',
        range='tab2!A1',
        body=dict(
            majorDimension='ROWS',
            values=rngData2
        )
    ).execute()
    
    print (excel_file, "done")
