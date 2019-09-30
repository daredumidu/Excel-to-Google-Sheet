from sheet_service_v2 import Create_Service
import win32com.client as win32


xlApp = win32.Dispatch('Excel.Application')
wb = xlApp.Workbooks.Open(r"C:\Users\usenadu\Documents\dumidu-pearson\google api\1\july.xlsx")         # wb --> workbook
ws = wb.Worksheets('tab1')                                                                             # ws --> worksheet
rngData = ws.Range('A1').CurrentRegion()


# Google Sheet Id
gsheet_id = '1LaVWGtYFJCXakvh0HDVk9nI_B1LnoJ2dQ5pR-I94smw'      # july
CLIENT_SECRET_FILE = 'client_secret.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)


response = service.spreadsheets().values().clear(
    spreadsheetId=gsheet_id,
    range='tab1!A1:C40',
).execute()


response = service.spreadsheets().values().append(
    spreadsheetId=gsheet_id,
    valueInputOption='RAW',
    range='tab1!A1',
    body=dict(
        majorDimension='ROWS',
        values=rngData
    )
).execute()
