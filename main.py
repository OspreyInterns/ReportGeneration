# Read SQLite file, then calculate the contrast straight to the patient, output to .xlsx


from tkinter import filedialog
import clinical_data
import sales_data

# Main Method of project


fileName = filedialog.askopenfilename(initialdir='C:\\', title='Select database file', filetypes=(('sqlite files',
                                                                                '*.sqlite'), ('all files', '*.*')))
CMSW = fileName[-23:-20]
if CMSW[0] == '/':
    CMSW = CMSW.replace('/', '')
data_type = input('Enter \"clinical\" to generate a clinical report or \"sales\" to generate a sales report: ')
if data_type == 'clinical':
    clinical_data.excel_write(fileName, CMSW)
elif data_type == 'sales':
    sales_data.excel_write(fileName, CMSW)
