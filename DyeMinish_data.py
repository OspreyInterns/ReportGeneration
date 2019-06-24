
import sqlite3 as sqlite
import openpyxl
from openpyxl.styles import Alignment, PatternFill

# Reads from the injection table to sum up the injections


def straight_to_patient(case_number, file_name):

    _con = sqlite.connect(file_name)

    with _con:

        contrast_inj = 0.
        alt_contrast_inj = 0
        _cur = _con.cursor()
        _cur.execute('SELECT * FROM CMSWInjections')

        _rows = _cur.fetchall()

        for _row in _rows:
            if _row[1] == case_number and _row[5] == 1 and (_row[18] == 0 or _row[15] <= 20) and _row[30] == 0:
                contrast_inj += _row[20]
            if _row[1] == case_number and _row[5] == 1 and (_row[17] == 0 or _row[15] <= 20):
                alt_contrast_inj += _row[20]

        return [contrast_inj, alt_contrast_inj]


def would_be_saved(file_name):

    con = sqlite.connect(file_name)

    with con:
        cur = con.cursor()
        cur.execute('SELECT * FROM CMSWCases')
        rows = cur.fetchall()
        case_info = []

        for row in rows:
            case_info.append([row[0], row[13], row[8], row[15]])

    what_if = []

    for case in case_info:
        direct_injected = straight_to_patient(case[0], file_name)
        vol_inj_off = direct_injected[0]
        vol_inj_on = case[3] - vol_inj_off
        vol_att_off = direct_injected[0]
        vol_att_on = case[1] - vol_att_off
        print(vol_att_off, vol_att_on, vol_inj_off, vol_inj_on)
        if vol_att_on != 0:
            perc_savings_on = 1. - (vol_inj_on / vol_att_on)
            print(perc_savings_on)
            would_be_total = vol_inj_on + (vol_inj_off * (1. - perc_savings_on))
            if case[2] != 0:
                would_be_portion = (would_be_total / case[2]) * 100
            else:
                would_be_portion = 0
            what_if.append([would_be_total, would_be_portion])
        else:
            what_if.append([0, 0])

    return what_if


def list_builder(file_name):

    con = sqlite.connect(file_name)

    with con:

        cur = con.cursor()
        cur.execute('SELECT * FROM CMSWCases')
        rows = cur.fetchall()
        check_cases = []
        what_if = would_be_saved(file_name)

        for row in rows:
            to_patient = straight_to_patient(row[0], file_name)
            if row[8] == 0:
                perc_threshold = 'N/A'
            else:
                perc_threshold = row[15] / row[8] * 100
            if row[2] == '2.1.24':
                check_cases.append(('', '', '', row[5][0:10], row[1][-12:-4], '', row[19], row[8], row[13],
                                    row[14], row[15], row[16], perc_threshold, '', to_patient[0], '', '',
                                    what_if[row[0] - 1][0], what_if[row[0] - 1][1]))
            else:
                check_cases.append(('', '', '', row[5][0:10], row[1][-12:-4], row[20][-8:], row[19], row[8], row[13],
                                    row[14], row[15], row[16], perc_threshold, '', to_patient[0], '', '',
                                    what_if[row[0] - 1][0], what_if[row[0] - 1][1]))

        return check_cases


def excel_flag_write(file_name, cmsw):

    check_cases = list_builder(file_name)

    xlsx_name = cmsw + 'DyeMinishFlaggedOutput.xlsx'
    wb = openpyxl.load_workbook('Dyeminish-template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'

    for row in range(len(check_cases)):
        for col in range(len(check_cases[row])):
            if float(check_cases[row][6]) <= 5.:
                data_sheet.cell(row=row + 2, column=col + 1, value=check_cases[row][col]).fill = PatternFill(
                    fill_type="solid", start_color='FFFF00', end_color='FFFF00')
                data_sheet.cell(row=row + 2, column=16, value='Case less than 5 Minutes')
            elif check_cases[row][8] == 0 and check_cases[row][9] == 0 and check_cases[row][10] == 0 \
                    and check_cases[row][11] == 0:
                data_sheet.cell(row=row + 2, column=col + 1, value=check_cases[row][col]).fill = PatternFill(
                    fill_type="solid", start_color='FFFF00', end_color='FFFF00')
                data_sheet.cell(row=row + 2, column=16, value='No contrast injected')
            else:
                data_sheet.cell(row=row + 2, column=col + 1, value=check_cases[row][col])
            data_sheet.cell(row=row + 2, column=col + 1).alignment = Alignment(wrapText=True)

    wb.save(xlsx_name)


def excel_destructive_write(file_name, cmsw):

    check_cases = list_builder(file_name)
    remove_cases = []
    for case in check_cases:
        if case[6] <= 5. or int(case[8]) == int(case[9]) == int(case[10]) == int(case[11]) == 0:
            remove_cases.append(case)

    for case in remove_cases:
        check_cases.remove(case)

    xlsx_name = cmsw + 'DyeMinishFilteredOutput.xlsx'
    wb = openpyxl.load_workbook('Dyeminish-template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'

    for row in range(len(check_cases)):
        for col in range(len(check_cases[row])):
            data_sheet.cell(row=row+2, column=col+1, value=check_cases[row][col])
            data_sheet.cell(row=row+2, column=col+1).alignment = Alignment(wrapText=True)

    wb.save(xlsx_name)
