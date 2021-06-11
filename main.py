# Read SQLite file, then calculate the contrast straight to the patient, output to .xlsx
from tkinter import filedialog
import tkinter as tk
import logging
import os
import DyeMinish_data
import sales_data
import rods_rockin_data
import cmsw_read

# Main Method of project, creates UI and takes input to pass to functions


class Application(tk.Frame):

    def __init__(self, master=None):

        self.delete = tk.BooleanVar()

        super().__init__(master)
        self.master = master
        self.pack()

        self.dyeminish = tk.Button(self)
        self.dyeminish['text'] = 'Generate Dyeminish Report'
        self.dyeminish['command'] = self.dyeminish_report
        self.dyeminish.pack(side='top')

        # self.delete_flag = tk.Checkbutton(self)
        # self.delete_flag['text'] = 'Check to delete flagged entries'
        # self.delete_flag['variable'] = self.delete
        # self.delete_flag['onvalue'] = True
        # self.delete_flag['offvalue'] = False
        # self.delete_flag.pack(side='top')

        self.sales = tk.Button(self)
        self.sales['text'] = 'Generate Sales Report'
        self.sales['command'] = self.sales_report
        self.sales.pack(side='top')

        self.rods = tk.Button(self)
        self.rods['text'] = 'Generate Rod\'s\n Radical Report'
        self.rods['command'] = self.rods_report
        self.rods.pack(side='top')

        self.all = tk.Button(self)
        self.all['text'] = 'I want it all'
        self.all['command'] = self.all_reports
        self.all.pack(side='top')

        self.quit = tk.Button(self, text='QUIT', fg='red',
                              command=self.master.destroy)
        self.quit.pack(side='top')

    def dyeminish_report(self):
        """Opens file browser, processes chosen files, checks the delete flag,
        then calls the appropriate write method
        """
        file_names = filedialog.askopenfilenames(title='Select database file',
                                                 filetypes=(('sqlite files', '*.sqlite'), ('all files', '*.*')))
        cmsws = []
        print('Processing file selection', end='')
        for file in file_names:
            cmsws.append(str(cmsw_read.cmsw_id_read(file)))
            print('.', end='')

        number = 0

        for cmsw in cmsws:
            cmsws[number] = cmsw
            number += 1
        print('')
        print('Input ready, beginning report...')
        try:
            # if not self.delete.get():
            DyeMinish_data.excel_flag_write(file_names, cmsws)
            # else:
                # DyeMinish_data.excel_destructive_write(file_names, cmsws)
        except Exception:
            logging.exception('Unexpected issue')
        print('Done')

    @staticmethod
    def sales_report():
        """Opens file browser, processes chosen files, then calls the write method"""
        cmsws = []
        number = 0
        file_names = filedialog.askopenfilenames(title='Select database file',
                                                 filetypes=(('sqlite files', '*.sqlite'), ('all files', '*.*')))
        print('Processing file selection', end='')
        file_names = list(file_names)
        for file in file_names:
            cmsws.append(str(cmsw_read.cmsw_id_read(file)))
            print('.', end='')
        for cmsw in cmsws:
            cmsw = cmsw.replace('/', '')
            cmsws[number] = cmsw
            number += 1
        print('')
        print('Input ready, beginning report...')
        try:
            sales_data.write(file_names, cmsws)
        except Exception:
            logging.exception('Unexpected issue')
        print('Done')

    @staticmethod
    def rods_report():
        """Opens file browser, processes chosen files, then calls the excel_write method"""
        cmsws = []
        number = 0
        file_names = filedialog.askopenfilenames(title='Select database file',
                                                 filetypes=(('sqlite files', '*.sqlite'), ('all files', '*.*')))
        print('Processing file selection', end='')
        file_names = list(file_names)
        for file in file_names:
            cmsws.append(str(cmsw_read.cmsw_id_read(file)))
            print('.', end='')
        for cmsw in cmsws:
            cmsw = cmsw.replace('/', '')
            cmsws[number] = cmsw
            number += 1
        print('')
        print('Input ready, beginning report...')
        try:
            rods_rockin_data.excel_write(file_names, cmsws)
        except Exception:
            logging.exception('Unexpected issue')
        print('Done')

    @staticmethod
    def all_reports():
        cmsws = []
        number = 0
        file_names = filedialog.askopenfilenames(title='Select database file',
                                                 filetypes=(('sqlite files', '*.sqlite'), ('all files', '*.*')))
        print('Processing file selection', end='')
        file_names = list(file_names)
        for file in file_names:
            cmsws.append(str(cmsw_read.cmsw_id_read(file)))
            print('.', end='')
        for cmsw in cmsws:
            cmsw = cmsw.replace('/', '')
            cmsws[number] = cmsw
            number += 1
        print('')
        print('Input ready, beginning report...')
        try:
            rods_rockin_data.excel_write(file_names, cmsws)
            sales_data.write(file_names, cmsws)
            DyeMinish_data.excel_flag_write(file_names, cmsws)
            # DyeMinish_data.excel_destructive_write(file_names, cmsws)
        except Exception:
            logging.exception('Unexpected issue')
        print('Done')


if os.path.isfile('Dyeminish-template.xlsx') and os.path.isfile('Rods-Template.xlsx') \
        and os.path.isfile('Sales-Template.xlsx') and os.path.isfile('Slide-Template.pptx'):
    logging.basicConfig(filename='Report.log', filemode='w')
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
else:
    if not os.path.isfile('Dyeminish-template.xlsx'):
        print('Missing Dyeminish-template.xlsx')
    if not os.path.isfile('Rods-Template.xlsx'):
        print('Missing Rods-Template.xlsx')
    if not os.path.isfile('Sales-Template.xlsx'):
        print('Missing Sales-Template.xlsx')
    if not os.path.isfile('Slide-Template.pptx'):
        print('Missing Slide-Template.pptx')
