# Read SQLite file, then calculate the contrast straight to the patient, output to .xlsx

#  import csv
import pprint
import sqlite3 as sqlite
import openpyxl
from tkinter import filedialog
from openpyxl.styles import Alignment
import clinical_data

# Main Method of project


fileName = filedialog.askopenfilename(initialdir='C:\\', title='Select database file', filetypes=(('sqlite files','*.sqlite'),('all files','*.*')))
CMSW = fileName[-23:-20]
if CMSW[0] == '/':
    CMSW = CMSW.replace('/', '')
con = sqlite.connect(fileName)

with con:

    cur = con.cursor()
    cur.execute('SELECT * FROM CMSWCases')

    col_names = [cn[0] for cn in cur.description]

    rows = cur.fetchall()

    checkCases = [('Case ID', 'Case Number', '% of contrast saved', 'Contrast injected straight to patient', 'Volume-% mismatch'), ]

    for row in rows:
        # if row[16]<35.:    # check if diverted contrast is under a threshold
        to_patient = clinical_data.straight_to_patient(row[0], fileName)
        if to_patient[1]:
            checkCases.append((row[1], row[0], round(row[16], 1), round(to_patient[0], 1), to_patient[1]))
        else:
            checkCases.append((row[1], row[0], round(row[16], 1), round(to_patient[0], 1)))
    pprint.pprint(checkCases)


#  CsvName = CMSW + 'directinjected.csv'
XlsxName = CMSW + 'directinjected.xlsx'
''' replaced by .xlsx
with open(CsvName, 'w', newline='') as csvFile:
    writer = csv.writer(csvFile)
    writer.writerows(checkCases)

csvFile.close()
'''
wb = openpyxl.Workbook()
dataSheet = wb.active
for row in range(0, len(checkCases)):
    for col in range(0, len(checkCases[row])):
        dataSheet.cell(row=row+1, column=col+1, value=checkCases[row][col]).alignment = Alignment(wrapText=True)

wb.save(XlsxName)
