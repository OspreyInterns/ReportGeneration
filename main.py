# Read SQLite file, then calculate the contrast straight to the patient, output to .xlsx


from tkinter import filedialog
import clinical_data

# Main Method of project


fileName = filedialog.askopenfilename(initialdir='C:\\', title='Select database file', filetypes=(('sqlite files','*.sqlite'),('all files','*.*')))
CMSW = fileName[-23:-20]
if CMSW[0] == '/':
    CMSW = CMSW.replace('/', '')
clinical_data.excel_write(fileName, CMSW)
