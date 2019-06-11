# Read SQLite file, then calculate the contrast straight to the patient, output to .xlsx


from tkinter import filedialog
import tkinter as tk
import clinical_data
import sales_data

# Main Method of project

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.clinical = tk.Button(self)
        self.clinical['text'] = 'Generate Clinical Report'
        self.clinical['command'] = self.clinical_report
        self.clinical.pack(side='top')

        self.sales = tk.Button(self)
        self.sales['text'] = 'Generate Sales Report'
        self.sales['command'] = self.sales_report
        self.sales.pack(side='right')

        self.quit = tk.Button(self, text='QUIT', fg='red',
                              command=self.master.destroy)
        self.quit.pack(side='bottom')

    def clinical_report(self):
        fileName = filedialog.askopenfilename(initialdir='C:\\', title='Select database file',
                                                  filetypes=(('sqlite files',
                                                              '*.sqlite'), ('all files', '*.*')))
        CMSW = fileName[-23:-20]
        if CMSW[0] == '/':
            CMSW = CMSW.replace('/', '')
        clinical_data.excel_write(fileName, CMSW)

    def sales_report(self):
        CMSWs = []
        fileNames = filedialog.askopenfilenames(title='Select database file',
                                                  filetypes=(('sqlite files','*.sqlite'), ('all files', '*.*')))
        fileNames = list(fileNames)
        for file in fileNames:
            if '/' in file:
                file = file.replace('/', '')
            CMSWs.append(file[-23:-20])
        sales_data.excel_write(fileNames, CMSWs)


root = tk.Tk()
app = Application(master=root)
app.mainloop()
