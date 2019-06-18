
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

                case_id_number[row[0]] = row[1][-23:-4]

        with con:

            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWInjections')
            rows = cur.fetchall()

            for row in rows:
                puff_inj = ''
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
                if inj_asp == 'INJ':
                    if row[20]+row[17] >= 3:
                        puff_inj = 'Injection'
                    elif row[20]+row[17] <= 2:
                        puff_inj = 'Puff'
                    elif row[28] >= 2.5:
                        puff_inj = 'Injection'
                    elif row[28] <= 2:
                        puff_inj = 'Puff'
                cases.append([row[2], row[3], row[4], row[34], row[35], inj_asp, contrast_asp, replacement, row[7],
                              row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16],
                              round(row[17], 2), row[18], row[19], round(row[20], 2), round(row[21], 2), row[22],
                              row[24], row[33], row[25], row[26], row[27], round(row[28], 2), round(row[29], 2),
                              row[30], row[31], row[32], '', round(row[20]+row[17], 2), puff_inj])

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
                #  if row[19] <= 5 or row[13] == row[14] == row[15] == row[16] == 0:
                #      pass
                #  else:
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
                              uses[1], uses[3], uses[0], uses[2], int(row[3])))
    cases.sort(key=sort_criteria)

    return cases


def dyevert_uses(case_number, file_name):

    _con = sqlite.connect(file_name)
    dyevert_used_inj = 0
    dyevert_not_used_inj = 0
    dyevert_used_puff = 0
    dyevert_not_used_puff = 0
    vol_used_inj = 0
    vol_not_used_inj = 0
    vol_used_puff = 0
    vol_not_used_puff = 0

    with _con:

        _cur = _con.cursor()
        _cur.execute('SELECT * FROM CMSWInjections')
        _rows = _cur.fetchall()

        for _row in _rows:
            if _row[20] + _row[17] >= 3:
                puff_inj = 1
            elif _row[20] + _row[17] <= 2:
                puff_inj = 2
            elif _row[28] >= 2.5:
                puff_inj = 1
            elif _row[28] <= 2:
                puff_inj = 2
            if round(_row[28], 2) != 0 and round(_row[20], 2) != 0:
                if _row[1] == case_number and _row[5] == 1 and _row[18] == 0 and puff_inj == 1:
                    dyevert_not_used_inj += 1
                    vol_not_used_inj += _row[20]
                elif _row[1] == case_number and _row[5] == 1 and puff_inj == 1:
                    dyevert_used_inj += 1
                    vol_used_inj += _row[20]
                elif _row[1] == case_number and _row[5] == 1 and _row[18] == 0 and puff_inj == 2:
                    dyevert_not_used_puff += 1
                    vol_not_used_puff += _row[20]
                elif _row[1] == case_number and _row[5] == 1 and puff_inj == 2:
                    dyevert_used_puff += 1
                    vol_used_puff += _row[20]

        return [dyevert_not_used_inj, dyevert_used_inj, dyevert_not_used_puff, dyevert_used_puff]


def excel_write(file_names, cmsw):

    cases = list_builder(file_names)
    xlsx1_name = str(cmsw).replace('s', '') + 'rods-case-data.xlsx'
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
    xlsx2_name = str(cmsw).replace('s', '') + 'rods-detailed-data.xlsx'
    wb = openpyxl.load_workbook('Rods-Other-Template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'

    for row in range(len(injections)):
        for col in range(len(injections[row])):
            if len(injections[row]) >= 10:
                if injections[row][35] == 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                elif injections[row][19] == 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                    data_sheet.cell(row=row + 4, column=20).font = Font(bold=True)
                    data_sheet.cell(row=row + 4, column=col + 1).fill = PatternFill(
                        fill_type="solid", start_color='FBE798', end_color='FBE798')
                elif injections[row][19] != 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                    data_sheet.cell(row=row + 4, column=20).font = Font(bold=True)
                    data_sheet.cell(row=row + 4, column=col + 1).fill = PatternFill(
                        fill_type="solid", start_color='C5E1B3', end_color='C5E1B3')
                if (injections[row][5] == 'ASP' and injections[row][6] != 'Yes') or \
                        (injections[row][5] == 'INJ' and (injections[row][28] == 0 or injections[row][21] == 0)):
                    data_sheet.row_dimensions[row+4].hidden = True
            else:
                data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
            data_sheet.cell(row=row + 4, column=col + 1).alignment = Alignment(wrapText=True)

    wb.save(xlsx2_name)
