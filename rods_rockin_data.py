
import sqlite3 as sqlite
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

# Write data for sales team to appropriate template


def sort_criteria(case):
    return case[1], case[2]


def injection_table(file_names):
    cases = []
    case_number = 0

    for file_name in file_names:
        con = sqlite.connect(file_name)

        with con:
            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWCases')
            rows = cur.fetchall()

            case_id_number = {}

            for row in rows:
                case_id_number[row[0]] = row[1][3:22]

        with con:

            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWInjections')
            rows = cur.fetchall()

            for row in rows:
                if case_number != row[1]:
                    case_number = row[1]
                    _cmsw = str(file_name[-23:-20]).replace('/', '')
                    cases.append(['CMSW', '', '', '', '', int(_cmsw)])
                    cases.append(['Case', '', '', '', '', case_id_number[case_number]])
                if row[5] == 1:
                    inj_asp = 'INJ'
                else:
                    inj_asp = 'ASP'
                if row[6] == 1:
                    contrast_asp = 'Yes'
                else:
                    contrast_asp = ''
                if row[36] == 1:
                    replacement = 'Yes'
                else:
                    replacement = ''
                cases.append([row[2], row[3], row[4], row[34], row[35], inj_asp, contrast_asp, replacement, row[7],
                              row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17],
                              row[18], row[19], row[20], row[21], row[22], row[24], row[33], row[25], row[26], row[27],
                              row[28], row[29], row[30], row[31], row[32], '', row[20]+row[17], ''])

    return cases


def list_builder(file_names):

    cases = []

    for file_name in file_names:
        con = sqlite.connect(file_name)

        with con:

            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWCases')
            rows = cur.fetchall()

            for row in rows:
                uses = dyevert_uses(row[0], file_name)
                if row[19] <= 5 or row[13] == row[14] == row[15] == row[16] == 0:
                    pass
                else:
                    if row[15] <= row[8] / 3 <= row[13]:
                        color = 1
                    elif row[15] <= row[8] * 2 / 3 <= row[13]:
                        color = 2
                    elif row[15] <= row[8] <= row[13]:
                        color = 3
                    elif row[15] >= row[8] <= row[13]:
                        color = 4
                    else:
                        color = 0
                    cases.append((color, row[5][0:10], row[5][11:22], row[8], row[13], row[15], row[14], row[16],
                                  '', uses[0], uses[1]))
    cases.sort(key=sort_criteria)
    return cases


def dyevert_uses(case_number, file_name):

    _con = sqlite.connect(file_name)

    with _con:

        dyevert_used = 0.
        dyevert_not_used = 0
        _cur = _con.cursor()
        _cur.execute('SELECT * FROM CMSWInjections')

        _rows = _cur.fetchall()

        for _row in _rows:
            if _row[1] == case_number and _row[5] == 1 and _row[18] == 0:
                dyevert_not_used += 1
            elif _row[1] == case_number and _row[5] == 1:
                dyevert_used += 1

        return [dyevert_not_used, dyevert_used]


def excel_write(file_names, cmsw):

    cases = list_builder(file_names)
    xlsx1_name = str(cmsw) + 'rods-rocking-data.xlsx'
    wb = openpyxl.load_workbook('Rods-Template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'

    for row in range(len(cases)):
        for col in range(len(cases[row])):
            data_sheet.cell(row=row + 17, column=col + 1, value=cases[row][col])
            data_sheet.cell(row=row + 17, column=col + 1).alignment = Alignment(wrapText=True)
    data_sheet.column_dimensions['A'].hidden = True
    wb.save(xlsx1_name)

    injections = injection_table(file_names)
    xlsx2_name = str(cmsw) + 'rods-radical-data.xlsx'
    wb = openpyxl.load_workbook('Rods-Other-Template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'

    for row in range(len(injections)):
        for col in range(len(injections[row])):
            if len(injections[row]) >= 10:
                if injections[row][35] == 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                    data_sheet.cell(row=row + 4, column=col + 1).alignment = Alignment(wrapText=True)
                elif injections[row][19] == 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                    data_sheet.cell(row=row + 4, column=col + 1).alignment = Alignment(wrapText=True)
                    data_sheet.cell(row=row + 4, column=20).font = Font(bold=True)
                    data_sheet.cell(row=row + 4, column=col + 1).fill = PatternFill(
                        fill_type="solid", start_color='FBE798', end_color='FBE798')
                elif injections[row][19] != 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                    data_sheet.cell(row=row + 4, column=col + 1).alignment = Alignment(wrapText=True)
                    data_sheet.cell(row=row + 4, column=20).font = Font(bold=True)
                    data_sheet.cell(row=row + 4, column=col + 1).fill = PatternFill(
                        fill_type="solid", start_color='C5E1B3', end_color='C5E1B3')
                if (injections[row][5] == 'ASP' and injections[row][6] != 'Yes') or \
                        (injections[row][5] == 'INJ' and injections[row][29] == 0):
                    data_sheet.row_dimensions[row+4].hidden = True
            else:
                data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                data_sheet.cell(row=row + 4, column=col + 1).alignment = Alignment(wrapText=True)
    wb.save(xlsx2_name)
