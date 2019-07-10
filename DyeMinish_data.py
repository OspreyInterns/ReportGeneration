
import sqlite3 as sqlite
import openpyxl
from openpyxl.styles import Alignment, PatternFill

# Reads from the injection table to sum up the injections


def straight_to_patient(file_name):

    con = sqlite.connect(file_name)

    with con:

        contrast_inj = []
        cur = con.cursor()
        cur.execute('SELECT * FROM CMSWInjections')

        rows = cur.fetchall()

        for n in range(len(rows)):
            contrast_inj.append([0, 0])
        for row in rows:
            if row[5] == 1 and (row[18] <= 1 or row[15] <= 20) and row[30] == 0:
                contrast_inj[row[1]-1][0] += row[20]
            if row[5] == 1 and (row[17] == 0 or row[15] <= 20):
                contrast_inj[row[1]-1][1] += row[20]

        return contrast_inj


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
    direct_injected = straight_to_patient(file_name)

    for case in case_info:
        vol_inj_off = direct_injected[case[0]-1][0]
        vol_inj_on = case[3] - vol_inj_off
        vol_att_on = case[1] - vol_inj_off
        if vol_att_on != 0:
            perc_savings_on = 1. - (vol_inj_on / vol_att_on)
            would_be_total = vol_inj_on + (vol_inj_off * (1. - perc_savings_on))
            if case[2] != 0:
                would_be_portion = (would_be_total / case[2]) * 100
            else:
                would_be_portion = 0
            what_if.append([would_be_total, would_be_portion])
        else:
            what_if.append([vol_inj_off, vol_inj_off/case[2]])

    return what_if


def list_builder(file_names):
    """Takes the list of files and builds the list of lists to write"""
    cases = []

    for file_name in file_names:

        con = sqlite.connect(file_name)

        with con:

            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWCases')
            rows = cur.fetchall()
            what_if = would_be_saved(file_name)
            to_patient = straight_to_patient(file_name)

            for row in rows:
                if row[8] == 0:
                    perc_threshold = 'N/A'
                else:
                    perc_threshold = row[15] / row[8] * 100
                if row[2] == '2.1.24':
                    cases.append(('', '', '', row[5][0:10], row[1][-12:-4], '', row[19], row[8], row[13],
                                 row[14], row[15], row[16], perc_threshold, '', to_patient[row[0]-1][0], '', '',
                                 what_if[row[0] - 1][0], what_if[row[0] - 1][1]))
                else:
                    cases.append(('', '', '', row[5][0:10], row[1][-12:-4], row[20][-8:], row[19], row[8], row[13],
                                 row[14], row[15], row[16], perc_threshold, '', to_patient[row[0]-1][0], '', '',
                                 what_if[row[0] - 1][0], what_if[row[0] - 1][1]))

    return cases


def excel_flag_write(file_names, cmsws):
    """Writes data into the an Excel Sheet
    Takes two inputs:
        -the file names of the CMSW databases
        -the serial numbers of the CMSWs
    Generates one files:
        -The flagged table, with data that hits possible removal criteria being highlighted in yellow
    """
    cases = list_builder(file_names)

    xlsx_name = str(cmsws) + 'DyeMinishFlaggedOutput.xlsx'
    wb = openpyxl.load_workbook('Dyeminish-template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'

    for row in range(len(cases)):
        for col in range(len(cases[row])):
            if float(cases[row][6]) <= 5.:
                data_sheet.cell(row=row + 2, column=col + 1, value=cases[row][col]).fill = PatternFill(
                    fill_type="solid", start_color='FFFF00', end_color='FFFF00')
                data_sheet.cell(row=row + 2, column=16, value='Case less than 5 Minutes')
            elif cases[row][8] == 0 and cases[row][9] == 0 and cases[row][10] == 0 \
                    and cases[row][11] == 0:
                data_sheet.cell(row=row + 2, column=col + 1, value=cases[row][col]).fill = PatternFill(
                    fill_type="solid", start_color='FFFF00', end_color='FFFF00')
                data_sheet.cell(row=row + 2, column=16, value='No contrast injected')
            else:
                data_sheet.cell(row=row + 2, column=col + 1, value=cases[row][col])
            data_sheet.cell(row=row + 2, column=col + 1).alignment = Alignment(wrapText=True)

    wb.save(xlsx_name)


def excel_destructive_write(file_names, cmsws):
    """Writes data into the an Excel Sheet
    Takes two inputs:
        -the file names of the CMSW databases
        -the serial numbers of the CMSWs
    Generates one files:
        -The flagged table, with data that hits possible removal criteria being excluded
    """
    cases = list_builder(file_names)
    remove_cases = []
    for case in cases:
        if case[6] <= 5. or int(case[8]) == int(case[9]) == int(case[10]) == int(case[11]) == 0:
            remove_cases.append(case)

    for case in remove_cases:
        cases.remove(case)

    xlsx_name = str(cmsws) + 'DyeMinishFilteredOutput.xlsx'
    wb = openpyxl.load_workbook('Dyeminish-template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'

    for row in range(len(cases)):
        for col in range(len(cases[row])):
            data_sheet.cell(row=row+2, column=col+1, value=cases[row][col])
            data_sheet.cell(row=row+2, column=col+1).alignment = Alignment(wrapText=True)

    wb.save(xlsx_name)
