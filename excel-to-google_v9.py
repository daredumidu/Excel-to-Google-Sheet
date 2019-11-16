from sheet_service_v2 import Create_Service
import os.path
import os
import re
import win32com.client as win32
import logging
import urllib.request, urllib.error
import glob

logging.basicConfig(level=logging.INFO, format='%(asctime)s :: %(levelname)s :: %(message)s', filename='excel-to-google.log')

# - - - - - - - - - - - - - - - file handling. - - - - - - - - - - - - - - - 
# source file path for excel and google sheet.
master_path = 'C:\\VM_Reports\\CTI_WEEKLY_REPORTS\\scheduled\\'
baseurl = 'https://docs.google.com/spreadsheets/d/'

f = open ("master_file.csv", "r")                                   # open source excel folder and gogole sheet id list.
f1 = f.readlines()

for line in f1:
    line = line.rstrip()                                            # rstrip - remove the new line.
    a = line.split(",")                                             # split the line from "comma".
    excel_folder,excel_name,gsheetid = a[0],a[1],a[2]

    excel_folder_path = os.path.join(master_path, excel_folder)     # print folder path.
    #print (excel_folder_path)

    excel_files = os.listdir(excel_folder_path)                     # print files in folder.
    #print (excel_files)

    if len(excel_files) == 0:
        print("%s is Empty" % excel_folder_path)
        logging.error("- - - - - - - - - - %s is Empty - - - - - - - - - -" % excel_folder_path)
    else:
        #print('Folder is Not Empty')

        # finds the specific file only.
        excel_files = glob.glob('%s/**/%s*.xlsx' % (excel_folder_path, excel_name), recursive=True)
        #print (excel_files)

        def extract_number(f):                                      # finds the latest file. 
            s = re.findall("\d+$",f)
            return (int(s[0]) if s else -1,f)
    
        latest_excel_file = max(excel_files,key=extract_number)     # print the latest file.
        print (latest_excel_file)
        logging.info(latest_excel_file)
    
        #excel_file_path = os.path.join(master_path, excel_folder, latest_excel_file) # print latest file with full path.
        #print (excel_file_path)
        #logging.info(excel_file_path)
        # - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

        # - - - - - reading data from excel file. - - - - -

        xlApp = win32.Dispatch('Excel.Application')                  
        wb = xlApp.Workbooks.Open(r"%s" % latest_excel_file)        # wb --> workbook
        ws1 = wb.Worksheets('Scan Information')                     # ws --> worksheet
        rngData1 = ws1.Range('A1').CurrentRegion()
        ws2 = wb.Worksheets('Vulnerability Data')
        rngData2 = ws2.Range('A1').CurrentRegion()
        wb.Close(True) 
        # - - - - - - - - - - - - - 

        # - - - - - google sheet operations. - - - - -
        print (baseurl + gsheetid)                                  # google sheet url.
        logging.info(baseurl + gsheetid)
        
        try:
            response = urllib.request.urlopen(baseurl + gsheetid).getcode()     # check if google sheet is available.
            #print (response)
            logging.info(response)
        
        except urllib.error.HTTPError as e:
            print ("Google sheet not found")
            logging.error("- - - - - - - - - - Google sheet not found - - - - - - - - - -")

        else:
            # define Google Sheet and api.
            gsheet_id = gsheetid
            CLIENT_SECRET_FILE = 'client_secret.json'
            API_SERVICE_NAME = 'sheets'
            API_VERSION = 'v4'
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
            service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)


            # code for excel and gsheet tab 1.
            # clear the data in the google sheet tab.
            response = service.spreadsheets().values().clear(
                spreadsheetId=gsheet_id,
                range='Scan Information!$A:$ZZ',                    # define the area to clear in google sheet.
            ).execute()

            # populate google sheet with new data.
            response = service.spreadsheets().values().append(
                spreadsheetId=gsheet_id,
                valueInputOption='RAW',
                range='Scan Information!A1',                        # define the tab and starting cell.
                body=dict(
                    majorDimension='ROWS',
                    values=rngData1
                )
            ).execute()


            # code for excel and gsheet tab 2.
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
            # - - - - - - - - - - - - - 
    
            print (latest_excel_file, "done")
            logging.info("- - - - - - - - - - %s done - - - - - - - - - -" % latest_excel_file)

logging.info("= = = = = = = = = = End of Logging = = = = = = = = = =")
