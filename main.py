# Read SQLite file, then calculate the contrast straight to the patient, output to .xlsx

from tkinter import filedialog
import tkinter as tk
import clinical_data
import sales_data
import rods_rockin_data

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
        self.sales.pack(side='top')

        self.rods = tk.Button(self)
        self.rods['text'] = 'Generate Rod\'s\n Radical Report'
        self.rods['command'] = self.rods_report
        self.rods.pack(side='top')

        self.quit = tk.Button(self, text='QUIT', fg='red',
                              command=self.master.destroy)
        self.quit.pack(side='top')

    def clinical_report(self):

        file_name = filedialog.askopenfilename(initialdir='C:\\', title='Select database file',
                                               filetypes=(('sqlite files', '*.sqlite'), ('all files', '*.*')))
        cmsw = file_name[-23:-20]
        if cmsw[0] == '/':
            cmsw = cmsw.replace('/', '')
        clinical_data.excel_write(file_name, cmsw)
        print('Done')

    def sales_report(self):

        cmsws = []
        file_names = filedialog.askopenfilenames(title='Select database file',
                                                 filetypes=(('sqlite files', '*.sqlite'), ('all files', '*.*')))
        file_names = list(file_names)
        for file in file_names:
            if '/' in file:
                file = file.replace('/', '')
            cmsws.append(file[-23:-20])
        sales_data.excel_write(file_names, cmsws)
        print('Done')

    def rods_report(self):

        cmsws = []
        file_names = filedialog.askopenfilenames(title='Select database file',
                                                 filetypes=(('sqlite files', '*.sqlite'), ('all files', '*.*')))
        file_names = list(file_names)
        for file in file_names:
            if '/' in file:
                file = file.replace('/', '')
            cmsws.append(file[-23:-20])
        rods_rockin_data.excel_write(file_names, cmsws)
        print('Done')


root = tk.Tk()
app = Application(master=root)
app.mainloop()
