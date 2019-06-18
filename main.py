# Read SQLite file, then calculate the contrast straight to the patient, output to .xlsx

from tkinter import filedialog
import tkinter as tk
import clinical_data
import sales_data
import rods_rockin_data

# Main Method of project


class Application(tk.Frame):

    def __init__(self, master=None):

        self.delete = tk.BooleanVar()

        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):

        self.dyeminishf = tk.Button(self)
        self.dyeminishf['text'] = 'Generate Dyeminsh Report \n with flagging'
        self.dyeminishf['command'] = self.dyeminish_report
        self.dyeminishf.pack(side='top')

        self.delete_flag = tk.Checkbutton(self)
        self.delete_flag['text'] = 'Check to delete flagged entries'
        self.delete_flag['variable'] = self.delete
        self.delete_flag['onvalue'] = True
        self.delete_flag['offvalue'] = False
        self.delete_flag.pack(side='top')

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

    def dyeminish_report(self):

        file_name = filedialog.askopenfilename(title='Select database file',
                                               filetypes=(('sqlite files', '*.sqlite'), ('all files', '*.*')))
        cmsw = file_name[-23:-20]
        if cmsw[0] == '/':
            cmsw = cmsw.replace('/', '')
        if not self.delete.get():
            clinical_data.excel_flag_write(file_name, cmsw)
        else:
            clinical_data.excel_destructive_write(file_name, cmsw)
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
