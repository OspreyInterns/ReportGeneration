
import sqlite3 as sqlite
import openpyxl
from openpyxl.styles import Alignment

# Write data for sales team to appropriate template


def list_builder(file_name, cmsw):

    con = sqlite.connect(file_name)

    with con:

        cur = con.cursor()
        cur.execute('SELECT * FROM CMSWCases')
        rows = cur.fetchall()
        cases = [('Case ID/Patient ID Field #', )]

        for row in rows:
            if row[13] >= row[8]/3 and row[15] <= row[8]/3:
                color = 1
            elif row[13] >= row[8]*2/3 and row[15] <= row[8]*2/3:
                color = 2
            elif row[13] >= row[8] and row[15] <= row[8]:
                color = 3
            elif row[13] >= row[8] and row[15] >= row[8]:
                color = 4
            else:
                color = 0
            cases.append((color, row[5][0:10], row[5][11:22], row[8], row[13], row[15], row[14], row[16]))

def excel_write(file_names, cmsw):

    cases = []
    for file_name in file_names:
        con = sqlite.connect(file_name)

        with con:

            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWCases')
            rows = cur.fetchall()

            for row in rows:
                if row[19] <= 5 or row[13] == row[14] == row[15] == row[16] == 0: #
                    pass
                else:
                    if row[15] <= row[8]/3 and row[13] >= row[8]/3:
                        color = 1
                    elif row[15] <= row[8]*2/3 and row[13] >= row[8]*2/3:
                        color = 2
                    elif row[15] <= row[8] and row[13] >= row[8]:
                        color = 3
                    elif row[15] >= row[8] and row[13] >= row[8]:
                        color = 4
                    else:
                        color = 0
                    cases.append((color, row[5][0:10], row[5][11:22], row[8], row[13], row[15], row[14], row[16]))
    xlsx_name = str(cmsw) + '-data-tables.xlsx'
    wb = openpyxl.load_workbook('Sales-Template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'
    for row in range(len(cases)):
        for col in range(len(cases[row])):
            data_sheet.cell(row=row + 17, column=col + 1, value=cases[row][col])
            data_sheet.cell(row=row + 17, column=col + 1).alignment = Alignment(wrapText=True)

    data_sheet.column_dimensions['A'].hidden = True
    wb.save(xlsx_name)
