
import sqlite3 as sqlite
import logging
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import cmsw_read

# case column numbers
CMSW_CASE_ID = 0
CASE_ID = 1
SERIAL_NUMBER = 3
DATE_OF_PROCEDURE = 5
THRESHOLD_VOLUME = 8
ATTEMPTED_CONTRAST_INJECTION_VOLUME = 13
DIVERTED_CONTRAST_VOLUME = 14
CUMULATIVE_VOLUME_TO_PATIENT = 15
PERCENTAGE_CONTRAST_DIVERTED = 16
TOTAL_DURATION = 19

#  injection column numbers
TIME_STAMP = 2
SYRINGE_REVISION = 3
PMDV_REVISION = 4
IS_AN_INJECTION = 5
IS_ASPIRATING_CONTRAST = 6
DYEVERT_DIAMETER = 7
SYRINGE_DIAMETER = 8
STARTING_SYRINGE_POSITION = 9
ENDING_SYRINGE_POSITION = 10
LINEAR_SYRINGE_MOVEMENT = 11
SYRINGE_VOLUME_INJECTED_OR_ASPIRATED = 12
STARTING_DYEVERT_POSITION = 13
ENDING_DYEVERT_POSITION = 14
LINEAR_DYEVERT_MOVEMENT = 15
DIVERT_VOLUME_DIVERTED = 16
DYEVERT_CONTRAST_VOLUME_DIVERTED = 17
PERCENT_CONTRAST_SAVED = 18
INJECTION_VOLUME_TO_PATIENT = 19
CONTRAST_VOLUME_TO_PATIENT = 20
CUMULATIVE_CONTRAST_VOLUME_TO_PATIENT = 21
OTHER_VOLUME_TO_PATIENT = 22
STARTING_CONTRAST_PERCENT_IN_SYRINGE = 24
STARTING_CONTRAST_PERCENT_IN_DYEVERT = 25
ENDING_CONTRAST_PERCENT_IN_DYEVERT = 26
DURATION = 27
FLOW_RATE_TO_FROM_SYRINGE = 28
FLOW_RATE_TO_PATIENT = 29
PREDOMINANT_CONTRAST_LINE_PRESSURE = 30
STARTING_DYEVERT_STOPCOCK_POSITION = 31
IS_SYSTEM_PAUSED = 32
ENDING_CONTRAST_PERCENT_IN_SYRINGE = 33
SYRINGE_ADDRESS = 34
PMDV_ADDRESS = 35
IS_DEVICE_REPLACEMENT = 36

# colors
WHITE = 0
LTGRN = 1
GREEN = 2
YELLOW = 3
RED = 4

# Write data for sales team to appropriate template


def _sort_criteria(case):
    """reads info for a sort"""
    return case[1], case[2]


def injection_table(file_names):
    """Connects to an individual database and determines which injections were puffs, injections,
    and leaves some uncatagorized to be classified by a person looking at the surrounding data
    """
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

                case_id_number[row[CMSW_CASE_ID]] = row[CASE_ID][-23:-4]

        with con:

            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWInjections')
            rows = cur.fetchall()

            for row in rows:
                puff_inj = ''
                if case_number != row[CASE_ID]:
                    case_number = row[CASE_ID]
                    _cmsw = str(file_name[-23:-20]).replace('/', '')
                    cases.append(['CMSW', '', '', '', '', _cmsw])
                    cases.append(['Case', '', '', '', '', case_id_number[case_number]])
                if row[IS_AN_INJECTION] == 1:
                    inj_asp = 'INJ'
                else:
                    inj_asp = 'ASP'
                if row[IS_ASPIRATING_CONTRAST] == 1:
                    contrast_asp = 'Yes'
                else:
                    contrast_asp = ''
                if row[IS_DEVICE_REPLACEMENT] == 1:
                    replacement = 'Yes'
                else:
                    replacement = ''
                if inj_asp == 'INJ':
                    if row[CONTRAST_VOLUME_TO_PATIENT]+row[DYEVERT_CONTRAST_VOLUME_DIVERTED] >= 3:
                        puff_inj = 'Injection'
                    elif row[CONTRAST_VOLUME_TO_PATIENT]+row[DYEVERT_CONTRAST_VOLUME_DIVERTED] <= 2:
                        puff_inj = 'Puff'
                    elif row[FLOW_RATE_TO_FROM_SYRINGE] >= 2.5:
                        puff_inj = 'Injection'
                    elif row[FLOW_RATE_TO_FROM_SYRINGE] <= 2:
                        puff_inj = 'Puff'
                    else:
                        debug_msg = 'Event ' + str(row[0]) + ' in cmsw ' + str(cmsw_read.cmsw_id_read(file_name)) + \
                                    ', case ' + str(row[CASE_ID]) + ' matched neither type'
                        logging.warning(debug_msg)
                cases.append([row[TIME_STAMP], row[SYRINGE_REVISION], row[PMDV_REVISION], row[SYRINGE_ADDRESS],
                              row[PMDV_ADDRESS], inj_asp, contrast_asp, replacement, row[DYEVERT_DIAMETER],
                              row[SYRINGE_DIAMETER], row[STARTING_SYRINGE_POSITION], row[ENDING_SYRINGE_POSITION],
                              row[LINEAR_SYRINGE_MOVEMENT], row[SYRINGE_VOLUME_INJECTED_OR_ASPIRATED],
                              row[STARTING_DYEVERT_POSITION], row[ENDING_DYEVERT_POSITION],
                              row[LINEAR_DYEVERT_MOVEMENT], row[DIVERT_VOLUME_DIVERTED],
                              round(row[DYEVERT_CONTRAST_VOLUME_DIVERTED], 2), row[PERCENT_CONTRAST_SAVED],
                              row[INJECTION_VOLUME_TO_PATIENT], round(row[CONTRAST_VOLUME_TO_PATIENT], 2),
                              round(row[CUMULATIVE_CONTRAST_VOLUME_TO_PATIENT], 2), row[OTHER_VOLUME_TO_PATIENT],
                              row[STARTING_CONTRAST_PERCENT_IN_SYRINGE], row[ENDING_CONTRAST_PERCENT_IN_SYRINGE],
                              row[STARTING_CONTRAST_PERCENT_IN_DYEVERT], row[ENDING_CONTRAST_PERCENT_IN_DYEVERT],
                              row[DURATION], round(row[FLOW_RATE_TO_FROM_SYRINGE], 2),
                              round(row[FLOW_RATE_TO_PATIENT], 2), row[PREDOMINANT_CONTRAST_LINE_PRESSURE],
                              row[STARTING_DYEVERT_STOPCOCK_POSITION], row[IS_SYSTEM_PAUSED], '',
                              round(row[CONTRAST_VOLUME_TO_PATIENT]+row[DYEVERT_CONTRAST_VOLUME_DIVERTED], 2),
                              puff_inj])

    return cases


def list_builder(file_names):
    """Takes the list of files and builds the list of lists to write"""
    cases = []

    for file_name in file_names:

        con = sqlite.connect(file_name)
        with con:

            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWCases')
            rows = cur.fetchall()
            uses = dyevert_uses(file_name)

            for row in rows:
                if not(row[TOTAL_DURATION] <= 5) and not(row[ATTEMPTED_CONTRAST_INJECTION_VOLUME]
                                                         == row[DIVERTED_CONTRAST_VOLUME]
                                                         == row[LINEAR_DYEVERT_MOVEMENT] == 0
                                                         and row[DIVERT_VOLUME_DIVERTED] <= 1):
                    if row[CUMULATIVE_VOLUME_TO_PATIENT] <= row[THRESHOLD_VOLUME] \
                            / 3 <= row[ATTEMPTED_CONTRAST_INJECTION_VOLUME]:
                        color = LTGRN
                    elif row[CUMULATIVE_VOLUME_TO_PATIENT] <= row[THRESHOLD_VOLUME] \
                            * 2 / 3 <= row[ATTEMPTED_CONTRAST_INJECTION_VOLUME]:
                        color = GREEN
                    elif row[CUMULATIVE_VOLUME_TO_PATIENT] <= row[THRESHOLD_VOLUME] \
                            <= row[ATTEMPTED_CONTRAST_INJECTION_VOLUME]:
                        color = YELLOW
                    elif row[CUMULATIVE_VOLUME_TO_PATIENT] >= row[THRESHOLD_VOLUME] \
                            <= row[ATTEMPTED_CONTRAST_INJECTION_VOLUME]:
                        color = RED
                    else:
                        color = WHITE
                    cases.append((color, row[DATE_OF_PROCEDURE][0:10], row[DATE_OF_PROCEDURE][11:22],
                                  row[THRESHOLD_VOLUME], row[ATTEMPTED_CONTRAST_INJECTION_VOLUME],
                                  row[CUMULATIVE_VOLUME_TO_PATIENT], row[DIVERTED_CONTRAST_VOLUME],
                                  row[PERCENTAGE_CONTRAST_DIVERTED], uses[row[CMSW_CASE_ID]][1],
                                  uses[row[CMSW_CASE_ID]][3], uses[row[CMSW_CASE_ID]][0],
                                  uses[row[CMSW_CASE_ID]][2], int(row[SERIAL_NUMBER])))
    cases.sort(key=_sort_criteria)

    return cases


def dyevert_uses(file_name):
    """Connects to an individual database to calculate the volume of contrast injected
    and the number of times contrast was injected both in puffs and injections
    Volume data is currently unused
    """
    con = sqlite.connect(file_name)
    dyevert_used_inj = 0
    dyevert_not_used_inj = 0
    dyevert_used_puff = 0
    dyevert_not_used_puff = 0
    vol_used_inj = 0
    vol_not_used_inj = 0
    vol_used_puff = 0
    vol_not_used_puff = 0
    case_number = 0
    uses = []

    with con:

        cur = con.cursor()
        cur.execute('SELECT * FROM CMSWInjections')
        rows = cur.fetchall()

        for ph in range(rows[-1][1]+1):
            uses.append([0, 0, 0, 0])
        for row in rows:
            if row[CASE_ID] != case_number:
                uses[case_number] = ([dyevert_not_used_inj, dyevert_used_inj, dyevert_not_used_puff, dyevert_used_puff])
                dyevert_used_inj = 0
                dyevert_not_used_inj = 0
                dyevert_used_puff = 0
                dyevert_not_used_puff = 0
                vol_used_inj = 0
                vol_not_used_inj = 0
                vol_used_puff = 0
                vol_not_used_puff = 0
                case_number = row[1]
            if row[CASE_ID] == case_number:
                if row[CONTRAST_VOLUME_TO_PATIENT] + row[DYEVERT_CONTRAST_VOLUME_DIVERTED] >= 3:
                    puff_inj = 1
                elif row[CONTRAST_VOLUME_TO_PATIENT] + row[DYEVERT_CONTRAST_VOLUME_DIVERTED] <= 2:
                    puff_inj = 2
                elif row[FLOW_RATE_TO_FROM_SYRINGE] >= 2.5:
                    puff_inj = 1
                elif row[FLOW_RATE_TO_FROM_SYRINGE] <= 2:
                    puff_inj = 2
                if round(row[FLOW_RATE_TO_FROM_SYRINGE], 2) != 0 and round(row[FLOW_RATE_TO_PATIENT], 2) != 0:
                    if row[IS_AN_INJECTION] == 1 and row[PERCENT_CONTRAST_SAVED] == 0 and puff_inj == 1:
                        dyevert_not_used_inj += 1
                        vol_not_used_inj += row[CONTRAST_VOLUME_TO_PATIENT]
                    elif row[IS_AN_INJECTION] == 1 and puff_inj == 1:
                        dyevert_used_inj += 1
                        vol_used_inj += row[CONTRAST_VOLUME_TO_PATIENT]
                    elif row[IS_AN_INJECTION] == 1 and row[PERCENT_CONTRAST_SAVED] == 0 and puff_inj == 2:
                        dyevert_not_used_puff += 1
                        vol_not_used_puff += row[CONTRAST_VOLUME_TO_PATIENT]
                    elif row[IS_AN_INJECTION] == 1 and puff_inj == 2:
                        dyevert_used_puff += 1
                        vol_used_puff += row[CONTRAST_VOLUME_TO_PATIENT]

        uses[case_number] = ([dyevert_not_used_inj, dyevert_used_inj, dyevert_not_used_puff, dyevert_used_puff])

        return uses


def excel_write(file_names, cmsw):
    """Writes data into the two Excel Sheets as seen in the example
    Takes two inputs:
        -the file names of the CMSW databases
        -the serial numbers of the CMSWs
    Generates two files:
        -The summary table, which as an augmented sales table
        -The in depth table, which details every injection from the databases
    """
    print('Processing Rod\'s summary data')
    cases = list_builder(file_names)
    xlsx1_name = str(cmsw) + 'rods-case-data.xlsx'
    wb = openpyxl.load_workbook('Rods-Template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'
    print('Writing summary data')
    for row in range(len(cases)):
        for col in range(len(cases[row])):
            data_sheet.cell(row=row + 17, column=col + 1, value=cases[row][col])
            data_sheet.cell(row=row + 17, column=col + 1).alignment = Alignment(wrapText=True)

    data_sheet.column_dimensions['A'].hidden = True
    wb.save(xlsx1_name)
    print('Summary data written, processing injection data')
    injections = injection_table(file_names)
    xlsx2_name = str(cmsw).replace('s', '') + 'rods-detailed-data.xlsx'
    wb = openpyxl.load_workbook('Rods-Other-Template.xlsx')
    data_sheet = wb.active
    data_sheet.title = 'Sheet1'
    print('Writing injection data', end='')
    for row in range(len(injections)):
        if row % 5000 == 0:
            print('.', end='')
        if len(injections[row]) >= 10 and injections[row][35] != 0:
            data_sheet.cell(row=row + 4, column=20).font = Font(bold=True)
        if (injections[row][5] == 'ASP' and injections[row][6] != 'Yes') or \
                (injections[row][5] == 'INJ' and (injections[row][28] == 0 or injections[row][21] == 0)):
            data_sheet.row_dimensions[row + 4].hidden = True
        for col in range(len(injections[row])):
            if len(injections[row]) >= 10:
                if injections[row][35] == 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                elif injections[row][19] == 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                    data_sheet.cell(row=row + 4, column=col + 1).fill = PatternFill(
                        fill_type="solid", start_color='FBE798', end_color='FBE798')
                elif injections[row][19] != 0:
                    data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
                    data_sheet.cell(row=row + 4, column=col + 1).fill = PatternFill(
                        fill_type="solid", start_color='C5E1B3', end_color='C5E1B3')
            else:
                data_sheet.cell(row=row + 4, column=col + 1, value=injections[row][col])
            data_sheet.cell(row=row + 4, column=col + 1).alignment = Alignment(wrapText=True)

    print('')
    print('Saving...')
    wb.save(xlsx2_name)
    print('Rod\'s report finished')
