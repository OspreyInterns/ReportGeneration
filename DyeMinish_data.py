
import sqlite3 as sqlite
import openpyxl
from openpyxl.styles import Alignment, PatternFill

# case column numbers
CMSW_CASE_ID = 0
CASE_ID = 1
DATE_OF_PROCEDURE = 5
DYEVERT_USED = 6
THRESHOLD_VOLUME = 8
ATTEMPTED_CONTRAST_INJECTION_VOLUME = 13
DIVERTED_CONTRAST_VOLUME = 14
CUMULATIVE_VOLUME_TO_PATIENT = 15
PERCENTAGE_CONTRAST_DIVERTED = 16
TOTAL_DURATION = 19
END_TIME = 20

#  injection column numbers
IS_AN_INJECTION = 5
LINEAR_DYEVERT_MOVEMENT = 15
DYEVERT_CONTRAST_VOLUME_DIVERTED = 17
PERCENT_CONTRAST_SAVED = 18
CONTRAST_VOLUME_TO_PATIENT = 20
PREDOMINANT_CONTRAST_LINE_PRESSURE = 30

# Highlight color
YELLOW = 'FFFF00'
# Reads from the injection table to sum up the injections


def straight_to_patient(file_name):

    con = sqlite.connect(file_name)

    with con:

        contrast_inj = []
        cur = con.cursor()
        cur.execute('SELECT * FROM CMSWInjections')

        rows = cur.fetchall()

        for placeholder in range(len(rows)):
            contrast_inj.append([0, 0])
        for row in rows:
            if row[IS_AN_INJECTION] == 1 and (row[PERCENT_CONTRAST_SAVED] <= 1 or row[LINEAR_DYEVERT_MOVEMENT] <= 20)\
                    and row[PREDOMINANT_CONTRAST_LINE_PRESSURE] == 0:
                contrast_inj[row[CASE_ID]-1][0] += row[CONTRAST_VOLUME_TO_PATIENT]
            if row[IS_AN_INJECTION] == 1 and (row[DYEVERT_CONTRAST_VOLUME_DIVERTED] == 0
                                              or row[LINEAR_DYEVERT_MOVEMENT] <= 20):
                contrast_inj[row[CASE_ID]-1][1] += row[CONTRAST_VOLUME_TO_PATIENT]

        return contrast_inj


def would_be_saved(file_name):

    con = sqlite.connect(file_name)

    with con:

        cur = con.cursor()
        cur.execute('SELECT * FROM CMSWCases')
        rows = cur.fetchall()
        case_info = []

        for row in rows:
            case_info.append([row[CMSW_CASE_ID], row[ATTEMPTED_CONTRAST_INJECTION_VOLUME],
                              row[THRESHOLD_VOLUME], row[LINEAR_DYEVERT_MOVEMENT]])

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
                pmdv = row[DYEVERT_USED]
                if row[THRESHOLD_VOLUME] == 0:
                    perc_threshold = 'N/A'
                else:
                    perc_threshold = row[CUMULATIVE_VOLUME_TO_PATIENT] / row[THRESHOLD_VOLUME] * 100
                if row[2] == '2.1.24':
                    cases.append(('', '', '', row[DATE_OF_PROCEDURE][0:10], row[CASE_ID][-12:-4], '',
                                 row[TOTAL_DURATION], row[THRESHOLD_VOLUME], row[ATTEMPTED_CONTRAST_INJECTION_VOLUME],
                                 row[DIVERTED_CONTRAST_VOLUME], row[CUMULATIVE_VOLUME_TO_PATIENT],
                                 row[PERCENTAGE_CONTRAST_DIVERTED], perc_threshold,
                                  '', to_patient[row[CMSW_CASE_ID]-1][0], '', '',
                                 what_if[row[CMSW_CASE_ID] - 1][0], what_if[row[CMSW_CASE_ID] - 1][1], pmdv))
                else:
                    cases.append(('', '', '', row[DATE_OF_PROCEDURE][0:10], row[CASE_ID][-12:-4],
                                 row[END_TIME][-8:], row[TOTAL_DURATION], row[THRESHOLD_VOLUME],
                                 row[ATTEMPTED_CONTRAST_INJECTION_VOLUME], row[DIVERTED_CONTRAST_VOLUME],
                                 row[CUMULATIVE_VOLUME_TO_PATIENT], row[PERCENTAGE_CONTRAST_DIVERTED], perc_threshold,
                                  '', to_patient[row[CMSW_CASE_ID]-1][0], '', '',
                                 what_if[row[CMSW_CASE_ID] - 1][0], what_if[row[CMSW_CASE_ID] - 1][1], pmdv))

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
        for col in range(len(cases[row])-1):
            if cases[row][19] == 0:
                data_sheet.cell(row=row + 2, column=col + 1, value=cases[row][col]).fill = PatternFill(
                    fill_type='solid', start_color=YELLOW, end_color=YELLOW)
            elif float(cases[row][6]) <= 5.:
                data_sheet.cell(row=row + 2, column=col + 1, value=cases[row][col]).fill = PatternFill(
                    fill_type='solid', start_color=YELLOW, end_color=YELLOW)
            elif cases[row][8] == 0 and cases[row][9] == 0 and cases[row][10] == 0 \
                    and cases[row][11] == 0:
                data_sheet.cell(row=row + 2, column=col + 1, value=cases[row][col]).fill = PatternFill(
                    fill_type='solid', start_color=YELLOW, end_color=YELLOW)
            else:
                data_sheet.cell(row=row + 2, column=col + 1, value=cases[row][col])
            data_sheet.cell(row=row + 2, column=col + 1).alignment = Alignment(wrapText=True)
        if cases[row][19] == 0:
            data_sheet.cell(row=row + 2, column=16, value='DyeTect Case')
            data_sheet.cell(row=row + 2, column=14, value='No')
        elif float(cases[row][6]) <= 5.:
            data_sheet.cell(row=row + 2, column=16, value='Case less than 5 Minutes')
            data_sheet.cell(row=row + 2, column=14, value='No')
        elif cases[row][8] == 0 and cases[row][9] == 0 and cases[row][10] == 0 \
                and cases[row][11] == 0:
            data_sheet.cell(row=row + 2, column=16, value='No contrast was injected')
            data_sheet.cell(row=row + 2, column=14, value='No')

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
        if case[6] <= 5. or int(case[8]) == int(case[9]) == int(case[10]) == int(case[11]) == 0 or case[19] == 0:
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
