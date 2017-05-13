import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import datetime

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("Viethope Email Address").sheet1

# Extract and print all of the values
list_of_hashes = sheet.get_all_records()

n = len(list_of_hashes)

for i in xrange(2, n + 2):
    print(" {} with {} \n".format(
        sheet.cell(i, 1).value, sheet.cell(i, 3).value))
    now_ = datetime.datetime.fromtimestamp(
        time.time()).strftime('%Y-%m-%d %H:%M:%S')
    sheet.update_cell(i, 4, now_)
