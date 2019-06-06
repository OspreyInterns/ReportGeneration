
#  import csv
#  import pprint
import sqlite3 as sqlite
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles import PatternFill

# Reads from the injection table to sum up the injections


def straight_to_patient(case_number: int, file_name):

    _con = sqlite.connect(file_name)

    with _con:
        contrast_inj = 0.
        alt_contrast_inj = 0
        _cur = _con.cursor()
        _cur.execute('SELECT * FROM CMSWInjections')

        _cols = _cur.fetchall()

        for _col in _cols:
            # _col[18](%) matches Alex's data, _col[17](volume) goes by volume diverted
            if _col[1] == case_number and _col[5] == 1 and _col[18] == 0:
                contrast_inj += _col[20]
                if _col[17] != 0:
                    print('Case', _col[1], 'contains a mismatch between % and volume diverted')
            if _col[1] == case_number and _col[5] == 1 and _col[17] == 0:
                alt_contrast_inj += _col[20]
                # _col[12] = total injection _col[16] = diverted volume _col[19] = total volume to patient
                # _col[30] = pressure _col[32] = pause _col[29] = flow rate to patient
                if round(_col[12], 4) != round(_col[16] + _col[19], 4) and _col[30] == 0 and _col[32] == 0:
                    if _col[29] != 0:
                        print('Injection', _col[0], 'suspicious', _col[12], '!=', _col[16] + _col[19])

        return [contrast_inj, alt_contrast_inj]


def excel_write(file_name, CMSW):
    con = sqlite.connect(file_name)

    with con:

        cur = con.cursor()
        cur.execute('SELECT * FROM CMSWCases')

        col_names = [cn[0] for cn in cur.description]

        rows = cur.fetchall()

        checkCases = [('Case ID/Patient ID Field #', )]

        for row in rows:
            # if row[16]<35.:    # check if diverted contrast is under a threshold
            to_patient = straight_to_patient(row[0], file_name)
            if row[8] == 0:
                perc_threshold = 'N/A'
            else:
                perc_threshold = row[15]/row[8]*100
            if row[2] == '2.1.24':
                checkCases.append(('', '', '', row[5][0:10], row[1][-12:-1], '', row[19], row[8], row[13],
                                   row[14], row[15], row[16], perc_threshold, '', to_patient[0], '', '', to_patient[1],
                                   to_patient[0] - to_patient[1]))
            else:
                checkCases.append(('', '', '', row[5][0:10], row[1][-12:-1], row[20][-8:-1], row[19], row[8], row[13],
                                  row[14], row[15], row[16], perc_threshold, '', to_patient[0], '', '', to_patient[1],
                                  to_patient[0]-to_patient[1]))
        #  pprint.pprint(checkCases)

    #  CsvName = CMSW + 'directinjected.csv'
    XlsxName = CMSW + 'DyeMinishOutput.xlsx'
    ''' replaced by .xlsx
    with open(CsvName, 'w', newline='') as csvFile:
        writer = csv.writer(csvFile)
        writer.writerows(checkCases)

    csvFile.close()
    '''
    wb = openpyxl.load_workbook('F173-A_template-DyeMINISH Display Data Summary.xlsx')
    dataSheet = wb.active
    for row in range(0, len(checkCases)):
        for col in range(0, len(checkCases[row])):
            if row != 0 and checkCases[row][6] <= 5.:
                dataSheet.cell(row=row + 1, column=col + 1, value=checkCases[row][col]).fill = PatternFill(
                    fill_type = "solid", start_color='FFFF00', end_color='FFFF00')
            elif row != 0 and checkCases[row][8] == 0 and checkCases[row][9] == 0 and checkCases[row][10] == 0 and checkCases[row][11] == 0:
                dataSheet.cell(row=row + 1, column=col + 1, value=checkCases[row][col]).fill = PatternFill(
                    fill_type="solid", start_color='FFFF00', end_color='FFFF00')
            else:
                dataSheet.cell(row=row + 1, column=col + 1, value=checkCases[row][col]).alignment = Alignment(wrapText=True)

    wb.save(XlsxName)
