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
    wb = xlApp.Workbooks.Open(r"%s" % excel_file_path)          # wb --> workbook
    ws1 = wb.Worksheets('Scan Information')                     # ws --> worksheet
    rngData1 = ws1.Range('A1').CurrentRegion()
    ws2 = wb.Worksheets('Vulnerability Data')
    rngData2 = ws2.Range('A1').CurrentRegion()
    wb.Close(True) 

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
        range='Scan Information!$A:$ZZ',
    ).execute()

    response = service.spreadsheets().values().append(
        spreadsheetId=gsheet_id,
        valueInputOption='RAW',
        range='Scan Information!A1',
        body=dict(
            majorDimension='ROWS',
            values=rngData1
        )
    ).execute()


    # code for excel tab 2
    response = service.spreadsheets().values().clear(
        spreadsheetId=gsheet_id,
        range='Vulnerability Data!$A:$ZZ',
    ).execute()

    response = service.spreadsheets().values().append(
        spreadsheetId=gsheet_id,
        valueInputOption='RAW',
        range='Vulnerability Data!A1',
        body=dict(
            majorDimension='ROWS',
            values=rngData2
        )
    ).execute()
    
    print (excel_file, "done")
