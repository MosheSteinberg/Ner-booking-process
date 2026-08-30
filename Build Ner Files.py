import json
import pandas as pd
from tkinter import ttk
from tkinter import Tk, StringVar, N, W, E, S, IntVar, filedialog, RIGHT, LEFT, TOP, BOTTOM
import os
from pathlib import Path

class MyWindow():
    def __init__(self, master):
        self.master = master
        master.title = 'Build .ner file'

        self.csv_row = 1
        self.csv_label = ttk.Label(text='CSV file')
        self.csv_label.grid(column=1, row=self.csv_row)

        self.csv_filepath_var = StringVar()
        self.csv_filepath_entry = ttk.Entry(textvariable=self.csv_filepath_var, width=60)
        self.csv_filepath_entry.grid(column=2, row=self.csv_row, sticky=(E, W))

        self.csv_browse_button = ttk.Button(text='Browse', command=self.file_explore_csv)
        self.csv_browse_button.grid(column=3, row=self.csv_row)

    def file_explore_csv(self):
        filename = filedialog.askopenfilename(initialdir=os.path.join(Path.home(), 'Downloads'),
                                            title='Select a file',
                                            filetypes=(('CSV files', '*.csv*'),))
        self.csv_filepath_var.set(filename)
        data = pd.read_csv(filename, on_bad_lines='skip')
        self.options = data.columns.values

        self.column_selection_row = self.csv_row + 1
        self.column_selection_var = StringVar()
        self.column_selection = ttk.OptionMenu(self.master, self.column_selection_var, self.options[10], *self.options)
        self.column_selection.grid(column=2, row=self.column_selection_row)

        self.sheet_name_var = StringVar()
        self.sheet_name = ttk.Entry(textvariable=self.sheet_name_var)
        self.sheet_name.grid(column=1, row=self.column_selection_row)

        self.dot_ner_file_var = StringVar()
        self.dot_ner_file_var.set('')
        self.dot_net_file_label = ttk.Label(textvariable=self.dot_ner_file_var)
        self.dot_net_file_label.grid(column=2, row=self.column_selection_row+100)

        self.add_selection_button = ttk.Button(text='Select column', command=self.add_column)
        self.add_selection_button.grid(column=3, row=self.column_selection_row)

        self.dot_net_file_name_label = ttk.Label(text='Name of .ner file')
        self.dot_net_file_name_label.grid(column=1, row=self.column_selection_row+2)

        self.dot_ner_file_name_var = StringVar()
        self.dot_ner_file_name_entry = ttk.Entry(textvariable=self.dot_ner_file_name_var)
        self.dot_ner_file_name_entry.grid(column=2, row=self.column_selection_row+2)

        self.dot_ner_file_name_button = ttk.Button(text='Create file', command=self.create_dot_ner_file)
        self.dot_ner_file_name_button.grid(column=3, row=self.column_selection_row+2)

    def add_column(self):
        current_json = self.dot_ner_file_var.get()
        if current_json == '':
            current_dic = {}
        else:
            current_dic = json.loads(current_json)
        current_dic[self.sheet_name_var.get()] = {"1":self.column_selection_var.get()}
        new_json = json.dumps(current_dic, indent=2)
        self.dot_ner_file_var.set(new_json)

    def create_dot_ner_file(self):
        filename = self.dot_ner_file_name_var.get()
        f = open(filename, mode='w')
        f.write(self.dot_ner_file_var.get())
        f.close()

root = Tk()
MyWindow(root)
root.mainloop()