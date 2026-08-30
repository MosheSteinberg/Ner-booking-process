import pandas as pd
import numpy as np
import xlsxwriter

from pathlib import Path
import os, sys, json

from tkinter import ttk
from tkinter import Tk, StringVar, N, W, E, S, IntVar
from tkinter import filedialog, PhotoImage, messagebox
from tkcalendar import DateEntry
from datetime import date, timedelta
from pandas import to_datetime
import traceback
import re

def FindMIfFloat(x):
    matchstring = r'\d{1,2}([:.]\d\d)?'
    find_time = re.search(matchstring, x)
    if find_time is None:
        return x
    else:
        return float(re.sub(':', '.', find_time.group()))

def IsInteger(x):
    try:
        int(x)
        return True
    except:
        return False

def run_process():
    try:
        input_fp = inputs_filepath.get()
        outputs_fp = outputs_filepath.get()
        title = label.get()

        selection_value = selection.get()
        with open(selection_value, mode='r') as ner_file:
            json_value = ner_file.read()
            columns_required = json.loads(json_value)

        raw_data = pd.read_csv(input_fp, on_bad_lines='skip', dtype='str')
        
        firstname_column = 7
        surname_column = 8
        raw_data.columns.values[firstname_column] = 'firstname'
        raw_data.columns.values[surname_column] = 'surname'
        check_deletion = delete_flag.get()

        check_duplicates = duplicates_flag.get()

        if check_deletion == 1:
            delete_before_date = pd.to_datetime(delete_before_entry.get_date())
            formatted_date_column = pd.to_datetime(raw_data['Submission Date'], format='%d/%m/%y %H:%M:%S')
            raw_data = raw_data[formatted_date_column > delete_before_date]

        writer = pd.ExcelWriter(outputs_fp, engine='xlsxwriter')
        workbook = writer.book
        format_cells = workbook.add_format({'font_size':22})
        format_titles = workbook.add_format()
        format_titles.set_text_wrap()
        format_titles.set_bold()
        format_titles.set_border()
        format_titles.set_align('center')
        format_titles.set_align('vcenter')
        format_counts = workbook.add_format()
        format_counts.set_align('center')

        format_duplicates = workbook.add_format({'bg_color':'red', 'font_color':'white'})

        # Loop through sheets
        for item, column_names in columns_required.items():
            # Get the name for the sheet
            sheet_name = str(item)
            ## Start the column count at 0
            number = 0
            ## Loop through the columns
            for col, column_name in column_names.items():
                # Pick out the Info
                info_item = raw_data.get(column_name)

                if info_item is None:
                    messagebox.showwarning(title='Column not found', detail=column_name + ' was not found')
                    continue

                info_item_save_commas = info_item.str.replace(', ', '>>')

                split_column = info_item_save_commas.str.get_dummies(sep=',')
                # Get unique list of options within the column
                unique_options_with_dash = [val.replace('>>', ', ') for val in list(split_column.columns.values)]
                unique_options = [i for i in unique_options_with_dash if i!='-']
                print(unique_options)
                # If > 1 option, sort by trying to find the time and turning it into a number
                UO_sort = sorted(unique_options)#, key=FindMIfFloat)
                # Loop through the options
                for option in UO_sort:
                    # Find which rows match the option
                    filterrows = split_column[option.replace(', ', '>>')]==1
                    # Title the column in Excel based on selection above
                    if col == '' or IsInteger(col):
                        name = option
                    else:
                        name = col

                    list_of_attendees_name = (raw_data['surname'].replace(np.nan, '', regex=True) + ', ' + raw_data['firstname'].replace(np.nan, '', regex=True))[filterrows].rename(name)
                    list_of_attendees_booker_name = raw_data['Person'][filterrows]

                    list_of_attendees_combined = pd.concat([list_of_attendees_name, list_of_attendees_booker_name], axis=1)


                    duplicate_flag = list_of_attendees_name.duplicated(keep=False)
                    list_of_attendees_name_no_duplicates = list_of_attendees_name.drop_duplicates()
                    duplicate_both_flag = list_of_attendees_combined[duplicate_flag].duplicated(keep=False)                
                    duplicates_with_different_booker = ~(list_of_attendees_booker_name[duplicate_flag] == list_of_attendees_name[duplicate_flag])
                    extra_names = list_of_attendees_booker_name[duplicate_flag][ ~duplicate_both_flag & duplicates_with_different_booker].rename(name)

                    if check_duplicates==0:
                        ListOfAttendees_Unordered = list_of_attendees_name
                    else:
                        # Filter the attendees and apply a header to the column
                        ListOfAttendees_Unordered = pd.concat([list_of_attendees_name_no_duplicates, extra_names]).drop_duplicates()
                    # Order the attendees alphabetically
                    ListOfAttendees_Ordered = ListOfAttendees_Unordered.sort_values(key=lambda x:x.str.lower())

                    # Count attendees
                    number_of_attendees = len(ListOfAttendees_Ordered)

                    output_columns = number_of_attendees // 40 + 1

                    for i in range(output_columns):
                        start_row = i * 40
                        end_row = min((i+1) * 40, number_of_attendees)
                        # Write the list to Excel in the 3rd row
                        ListOfAttendees_Ordered[start_row:end_row].to_excel(writer, sheet_name=sheet_name, index=False, startcol=number+i, startrow=3, header=False)
                    # Select the column and set its width
                    worksheet = writer.sheets[sheet_name]

                    if check_duplicates==0:
                        worksheet.conditional_format(3, number, 42, number+output_columns-1, {'type':'duplicate', 'format':format_duplicates})

                    print(type(ListOfAttendees_Ordered))
                    try:
                        max_name_length = ListOfAttendees_Ordered.map(lambda x: len(x)).max()
                    except:
                        max_name_length = 20
                    column_width = max(len(name)/2, max_name_length)
                    worksheet.set_column(number, number + output_columns - 1, column_width)
                    worksheet.write(2, number, name, format_titles)
                    #print(number)
                    # Write the name in the top left
                    worksheet.write('A1', sheet_name, format_cells)
                    # Write the title in next row
                    worksheet.write('A2', title, format_cells)
                    # Write count next to table
                    worksheet.write(2, number + 1, number_of_attendees, format_counts)
                    # Set height of first row
                    worksheet.set_row(0, 22)
                    worksheet.set_row(2, 30)
                    worksheet.fit_to_pages(1,1)
                    worksheet.set_paper(9)
                    # Increment 2 columns across for the next list
                    number += 1 + output_columns
                    
        # Save the workbook
        writer.close()

        # Open the workbook for the user
        os.startfile(outputs_fp)
    except:
        messagebox.showerror(title="Python error", detail= traceback.format_exc())

def file_explore_inputs():
    filename = filedialog.askopenfilename(initialdir=os.path.join(Path.home(), 'Downloads'),
                                          title='Select a file',
                                          filetypes=(('CSV files', '*.csv*'),))
    inputs_filepath.set(filename)

def file_explore_outputs():
    filename = filedialog.asksaveasfilename(initialdir=os.path.join(Path.home(), 'Documents'),
                                          title='Select a file',
                                          filetypes=(('Excel files', '*.xlsx'),))
    if filename[-5:] != '.xlsx':
        filename = filename + '.xlsx'
    outputs_filepath.set(filename)

if __name__=="__main__":
    try:
        should_I_download = sys.argv[1] == "Download"
    except:
        should_I_download = False
    
    if os.environ['USERNAME']=='SteinbergMoshe' and should_I_download:
        import get_from_web

    root = Tk()
    root.title("Ner booking process")
    root.resizable(0,0)

    mainframe = ttk.Frame(root, padding="3 3 12 12")
    mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)	

    inputs_filepath = StringVar()
    outputs_filepath = StringVar()
    label = StringVar()
    delete_before = StringVar()

    input_row = 1
    output_row = 2
    date_row = 3
    Label_Row = 4
    Selection_Row = 5
    duplicates_row = 6

    downloads_folder = os.path.join(os.environ['USERPROFILE'], "Downloads")
    files_in_downloads = [os.path.join(downloads_folder, x) for x in os.listdir(downloads_folder) if x.endswith('.csv')]
    inputs_filepath.set(max(files_in_downloads, key=os.path.getctime))

    ttk.Label(mainframe, text="Location of csv file").grid(column=1, row=input_row, sticky=W)
    inputs_filepath_entry = ttk.Entry(mainframe, width=60, textvariable=inputs_filepath)
    inputs_filepath_entry.grid(column=2, row=input_row, sticky=(W, E), columnspan=2)
    ttk.Button(mainframe, text="Browse", command=file_explore_inputs).grid(column=4, row=input_row, sticky=W)

    ttk.Label(mainframe, text="Name of output file").grid(column=1, row=output_row, sticky=W)
    outputs_filepath_entry = ttk.Entry(mainframe, textvariable=outputs_filepath)
    outputs_filepath_entry.grid(column=2, row=output_row, sticky=(W, E), columnspan=2)
    ttk.Button(mainframe, text="Browse", command=file_explore_outputs).grid(column=4, row=output_row, sticky=W)

    ttk.Label(mainframe, text="Delete entries from before:").grid(column=1, row=date_row, sticky=W)
    default_date = date.today() - timedelta(days=2)
    delete_before_entry = DateEntry(mainframe, locale='en_UK')
    delete_before_entry.set_date(default_date)
    delete_before_entry.grid(column=3, row=date_row, sticky=(W, E))
    delete_before_entry.grid_remove()

    def show_or_hide_date():
        check_flag = delete_flag.get()
        if check_flag==0:
            delete_before_entry.grid_remove()
        elif check_flag==1:
            delete_before_entry.grid(column=3, row=date_row, sticky=(W, E))
    delete_flag = IntVar()
    delete_flag.set(0)
    date_toggle = ttk.Checkbutton(mainframe, variable=delete_flag, command=show_or_hide_date)
    date_toggle.grid(column=2, row=date_row, sticky=(W, E))

    label = StringVar()
    ttk.Label(mainframe, text="Label").grid(column=1, row=Label_Row, sticky=W)
    label_entry = ttk.Entry(mainframe, textvariable=label)
    label_entry.grid(column=2, row=Label_Row, columnspan=2, sticky=(W, E))

    ttk.Label(mainframe, text="Select type").grid(column=1, row=Selection_Row, sticky=E)
    selection = StringVar()

    if getattr(sys, 'frozen', False):
        current_directory = os.path.dirname(sys.executable)
    else:
        current_directory = os.getcwd()
    Options = [json_file for json_file in os.listdir(current_directory) if json_file.endswith('.ner')]
    if 'Shabbos.ner' in Options:
        Default = 'Shabbos.ner'
    else:
        Default = Options[0]

    selection_dropdown = ttk.OptionMenu(mainframe, selection, Default, *Options)
    selection_dropdown.grid(column=2, row=Selection_Row, sticky=(W))

    ttk.Label(mainframe, text="Remove duplicates").grid(column=1, row=duplicates_row, sticky=E)
    duplicates_flag = IntVar()
    duplicates_flag.set(1)
    duplicates_toggle = ttk.Checkbutton(mainframe, variable=duplicates_flag)
    duplicates_toggle.grid(column=2, row=duplicates_row, sticky=(W, E))


    ttk.Button(mainframe, text="Run", command=run_process).grid(column=4, row=6, sticky=E)


    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.bind('<Return>', run_process)
    root.mainloop()
