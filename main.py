# Read SQLite file, then calculate the contrast straight to the patient, output to .xlsx

#  import csv
import pprint
import sqlite3 as sqlite
import openpyxl
from tkinter import filedialog
from openpyxl.styles import Alignment


# Reads from the injection table to sum up the injections
def straight_to_patient(case_number: int):

    _con = sqlite.connect(fileName)

    with _con:
        contrast_inj = 0.
        mismatch = False
        _cur = _con.cursor()
        _cur.execute('SELECT * FROM CMSWInjections')

        _cols = _cur.fetchall()

        for _col in _cols:
            # _col[18](%) matches Alex's data, _col[17](volume) goes by volume diverted
            if _col[1] == case_number and _col[5] == 1 and _col[18] == 0:
                contrast_inj += _col[20]
                if _col[17] != 0:
                    print('Case', _col[1], 'contains a mismatch between % and volume diverted')
                    mismatch = True
                # _col[12] = total injection _col[16] = diverted volume _col[19] = total volume to patient
                # _col[30] = pressure _col[32] = pause _col[29] = flow rate to patient
                if round(_col[12], 4) != round(_col[16] + _col[19], 4) and _col[30] == 0 and _col[32] == 0:
                    if _col[29] != 0:
                        print('Injection', _col[0], 'suspicious', _col[12], '!=', _col[16] + _col[19])

        return [contrast_inj, mismatch]


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
        to_patient = straight_to_patient(row[0])
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
