import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

spreadsheet_name = 'Reconciliation 2020-04-07 #3'
#Note: to work, share the spreadsheet to edit with editor@gigo-274322.iam.gserviceaccount.com

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('editor_secret.json', scope) #private key information
client = gspread.authorize(creds)

sheet = client.open(spreadsheet_name)

original_classification_task = sheet.get_worksheet(0)
#Note: get_worksheet(16) is the notes worksheet.

finalClassifications = original_classification_task.col_values(8)
listNA = [] #List of articles which are not original classifications.
for i in range(len(finalClassifications)):
    if finalClassifications[i] == 'no':
        listNA.append(i + 1)

#Indexing starts at 1 and ends at 15 so original_classification_task and notes are untouched
for worksheet in sheet.worksheets()[14:15]:
    for i in listNA:
        worksheet.update_cell(i, 8, "N/A") #column 8 is the Final column for all worksheets
        print("Cell", (i, 8), "set to N/A in worksheet", worksheet.title)
