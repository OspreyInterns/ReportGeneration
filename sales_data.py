
import sqlite3 as sqlite
import openpyxl
from openpyxl.styles import Alignment
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

# Write data for sales team to appropriate template


def sort_criteria(case):
    return case[1], case[2]


def list_builder(file_names):

    cases = []

    for file_name in file_names:
        con = sqlite.connect(file_name)

        with con:

            cur = con.cursor()
            cur.execute('SELECT * FROM CMSWCases')
            rows = cur.fetchall()

            for row in rows:
                if row[19] <= 5 or row[13] == row[14] == row[15] == row[16] == 0:  #
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
                    cases.append((color, row[5][0:10], row[5][11:22], row[8], row[13], row[15], row[14], row[16]))
    cases.sort(key=sort_criteria)
    return cases


def write(file_names, cmsw):

    cases = list_builder(file_names)
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

    pptx_name = str(cmsw) + '-slide.pptx'
    colors = [0, 0, 0, 0]
    attempted = 0
    diverted = 0

    for case in cases:
        if case[0] == 4:
            colors[0] += 1
        elif case[0] == 3:
            colors[1] += 1
        elif case[0] == 2:
            colors[2] += 1
        elif case[0] == 1:
            colors[3] += 1
        attempted += case[4]
        diverted += case[6]

    perc_saved = round(diverted/attempted*100)
    prs = Presentation('Slide-Template.pptx')
    total = colors[0] + colors[1] + colors[2] + colors[3]
    title = 'N=' + str(total)
    data = CategoryChartData()
    data.add_series(title, colors)
    data.categories = ['> Threshold, N=', '< Threshold, N=', '< 2/3 Threshold, N=', '< 1/3 Threshold, N=']
    prs.slides[0].shapes[4].chart.replace_data(data)
    text = ['All cases (N=' + str(len(cases)) + ')',
            str(perc_saved) + '% avg \nLess \nContrast',
            str(round(diverted)) + ' mL less total']

    for box in range(5, 8, 1):
        prs.slides[0].shapes[box].text_frame.clear()
        prs.slides[0].shapes[box].text_frame.paragraphs[0].text = text[box-5]
        prs.slides[0].shapes[box].text_frame.paragraphs[0].allignment = PP_ALIGN.CENTER
        prs.slides[0].shapes[box].text_frame.paragraphs[0].font.size = Pt(26-2*box)

    prs.slides[0].shapes[5].text_frame.paragraphs[0].font.bold = True
    prs.slides[0].shapes[7].text_frame.paragraphs[0].font.italic = True
    prs.save(pptx_name)
