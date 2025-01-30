from tkinter.tix import COLUMN
from unittest import result
import pandas as pd
import xlrd
import xlwt
from datetime import datetime, timedelta
from persiantools.jdatetime import JalaliDate
from xlutils.copy import copy
import os
import shutil
import tkinter as tk
from tkinter import filedialog
from functools import partial
import customtkinter

# Function to get the weekday name from a date string
def week_day(data_day):
    res = ''
    if(data_day): 
        s = data_day.split()
        res = s[1]
        if(len(s) == 3):
            res = ' '.join([s[1], s[2]])
    return res

# Function to manipulate dates, especially converting between Jalali and Gregorian dates
def date(data_date, num):
    jd = (data_date.split()[0]).split('-')
    
    # Convert to Gregorian date
    gd = JalaliDate(int(jd[0]), int(jd[1]), int(jd[2])).to_gregorian()
    gd = str(gd).replace("-", "/")
    date = datetime.strptime(gd, "%Y/%m/%d")
    
    # Modify the date by adding or subtracting days
    if(num < 0):
        modified_date = date - timedelta(days=(num * (-1)))
    else:    
        modified_date = date + timedelta(days=num)
    
    # Convert back to Jalali and return
    if(num != 0):    
        gd = datetime.strftime(modified_date, "%Y/%m/%d").split('/')
        jd = str(JalaliDate.to_jalali(int(gd[0]), int(gd[1]), int(gd[2])))    
    if(num == 0):
        jd = '-'.join(jd)
    return jd

# Function to get a list of Excel files in a specified directory
def folder_content(path):
    res = []
    for file in os.listdir(path):
        if file.endswith('.xls'):
            res.append(file)
    return res

# Helper function to standardize path formatting
def standard_path(path):
    list = path.split('/')
    return '\\'.join(list)

# Main process function that reads, processes, and saves the Excel file
def main_process(lbl, source_path, dest):
    # Standardize file paths
    source_path = standard_path(source_path.get())  
    dest = standard_path(dest.get())  
    
    # Open the source Excel file
    workbook = xlrd.open_workbook(source_path)

    list_of_files = folder_content(dest)  # Get list of existing files in destination
    tempsheet = workbook.sheet_by_index(4)
    first_unwriten_day = tempsheet.cell(1, 3).value.split()[0]
    last_e_date = date(first_unwriten_day, -1)
    file_name = []

    # Check if the file already exists
    for file in list_of_files:
        file_name.append(file.split('.')[0])
    path = ''    
    if(last_e_date in file_name):
        path = dest + '\\' + last_e_date + '.xls'    

    # Create a new workbook for saving processed data
    wb = xlwt.Workbook()
    full = True
    try:   
        rb = xlrd.open_workbook(path)
        worksheet = rb.sheet_by_index(0)
        len_rows = worksheet.nrows
        wbook = copy(rb)
        sh = wbook.get_sheet(0)    
        index = len_rows
        
        # Check if the current month data is complete
        day_e = int(worksheet.cell(len_rows-1, 1).value.split('-')[2])
        if(day_e == 25 ):
            full = True
        else:
            full = False    
    except Exception:
        pass
    
    writed = False  # Flag to track if data has been written
    for sh in range(4, len(workbook.sheet_names())):  # Loop through sheets
        sheet = workbook.sheet_by_index(sh)
        num_rows = sheet.nrows

        new_id = False
        a_data = [0, 15, 30]  # Different column indexes for data
        counter = 0
        for i in [9, 24, 39]:  # Loop through specific columns for employees
            new_id = True
            if(full):
                index = 1  # Write data starting from the second row if the file is full
            else:
                index = len_rows
            last_date = ''

            for j in range(11, num_rows):  # Iterate through rows for the current sheet
                id_cell = sheet.cell(3, i).value  
                
                if (id_cell != ''):
                    day = week_day(sheet.cell(j, a_data[counter] + 0).value)
                    if(last_date == ''):
                        last_date = sheet.cell(1, 3).value
                        num = 0
                    else:
                        num = 1
                    date_ = date(last_date, num)
                    last_date = date_
                    
                    # Handle Saturdays and Sundays for extra work and attendance
                    if(day == 'شنبه' or day == 'يکشنبه'):
                        start = sheet.cell(j, a_data[counter] + 10).value
                        if(start == 'غیبت'):
                            start = ''
                        end = sheet.cell(j, a_data[counter] + 12).value
                    else:    
                        start = sheet.cell(j, a_data[counter] + 1).value
                        if(start == 'غیبت'):
                            start = ''
                        end = sheet.cell(j, a_data[counter] + 3).value

                    # Write data to the new workbook
                    if(not full):
                        sh = wbook.get_sheet(id_cell)
                        sh.write(index, 0, index)
                        sh.write(index, 1, date_)
                        sh.write(index, 2, day)
                        sh.write(index, 3, int(id_cell))
                        sh.write(index, 4, start)
                        sh.write(index, 5, end)
                        index += 1
                        writed = True

                    # Add a new sheet if new employee ID is found
                    if(new_id and full):
                        new_id = False
                        w_sheet = wb.add_sheet(id_cell)
                        header_font = xlwt.Font()
                        header_font.bold = True
                        header_style = xlwt.XFStyle()
                        header_style.font = header_font

                        # Write column headers
                        w_sheet.write(0, 0, 'Row', header_style)
                        w_sheet.write(0, 1, 'Date', header_style)
                        w_sheet.write(0, 2, 'Weekday', header_style)
                        w_sheet.write(0, 3, 'Employee ID', header_style)
                        w_sheet.write(0, 4, 'Start Time', header_style)
                        w_sheet.write(0, 5, 'End Time', header_style)                    
                        writed = True

                    if(full):
                        w_sheet.write(index, 0, index)
                        w_sheet.write(index, 1, date_)
                        w_sheet.write(index, 2, day)
                        w_sheet.write(index, 3, int(id_cell))
                        w_sheet.write(index, 4, start)
                        w_sheet.write(index, 5, end)
                        index += 1            
            if(counter + 1 < 3):        
                counter += 1 
            
            new_path = dest + '\\' + last_date + '.xls' 

    # Save the workbook if any data was written
    if(writed):
        if(not full):
            if(path != new_path):
                shutil.move(path, new_path)
                path = new_path
            wbook.save(path)
        else:
            wb.save(new_path)
        result = 'file created!'    
        lbl.config(text=result)  # Display success message

# Setup Tkinter appearance and window
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")  

# Initialize Tkinter window
app = customtkinter.CTk()  
app.title('Excel File')
app.geometry("450x350+500+250")

source_path = customtkinter.StringVar()
dest = customtkinter.StringVar()

# Function to select the source file
def get_source_address():
    txtfld1.delete(0, tk.END)
    filename = filedialog.askopenfilename(initialdir = '/', title = 'Select File', filetype = (('executable, *.xls'), ('all files', '*.*')))
    txtfld1.insert(tk.END, filename)

# Function to select the destination folder
def get_dest_address():
    txtfld2.delete(0, tk.END)
    directory = filedialog.askdirectory(initialdir = '/', title = 'Select File')
    txtfld2.insert(tk.END, directory)

# UI Labels and Buttons
lbl1 = customtkinter.CTkLabel(master=app, text="excel report file: ")
lbl1.place(relx=0.13, rely=0.1, anchor=tk.CENTER)

lbl2 = customtkinter.CTkLabel(master=app, text="destination: ")
lbl2.place(relx=0.11, rely=0.3, anchor=tk.CENTER)

lbl3 = customtkinter.CTkLabel(master=app, text="")
lbl3.place(relx=0.4, rely=0.9, anchor=tk.CENTER)

txtfld1 = customtkinter.CTkEntry(master=app, textvariable=source_path)
txtfld1.place(relx=0.4, rely=0.1, anchor=tk.CENTER)

txtfld2 = customtkinter.CTkEntry(master=app, textvariable=dest)
txtfld2.place(relx=0.4, rely=0.3, anchor=tk.CENTER)

btn1 = customtkinter.CTkButton(master=app, text="select", command=get_source_address)
btn1.place(relx=0.8, rely=0.1, anchor=tk.CENTER)

btn2 = customtkinter.CTkButton(master=app, text="select", command=get_dest_address)
btn2.place(relx=0.8, rely=0.3, anchor=tk.CENTER)

# Button to start processing
main_process = partial(main_process, lbl3, source_path, dest) 
btn4 = customtkinter.CTkButton(master=app, text="process", command=main_process)
btn4.place(relx=0.8, rely=0.9, anchor=tk.CENTER)

# Start Tkinter main loop
app.mainloop()
