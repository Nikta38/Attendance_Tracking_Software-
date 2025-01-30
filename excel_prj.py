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


def week_day(data_day):
    res = ''
    if(data_day): 
        s = data_day.split()
        res = s[1]
        if(len(s) == 3):
            res = ' '.join([s[1], s[2]])
    return res

def date(data_date, num):
    jd = (data_date.split()[0]).split('-')
    
    gd = JalaliDate(int(jd[0]), int(jd[1]), int(jd[2])).to_gregorian()
    gd = str(gd).replace("-", "/")
    date = datetime.strptime(gd, "%Y/%m/%d")
    if(num < 0):
        modified_date = date - timedelta(days=(num * (-1)))
    else:    
        modified_date = date + timedelta(days=num)
    if(num != 0):    
        gd = datetime.strftime(modified_date, "%Y/%m/%d").split('/')
        jd = str(JalaliDate.to_jalali(int(gd[0]), int(gd[1]), int(gd[2])))    
    if(num == 0):
        jd = '-'.join(jd)
    return jd

def folder_content(path):
    res = []
    for file in os.listdir(path):
        if file.endswith('.xls'):
            res.append(file)
    return res

   
    date = sorted(list, key = lambda d: datetime.strptime(d, '%Y-%m-%d'), reverse=True)[0]
    return date

def standard_path(path):
    list = path.split('/')
    return '\\'.join(list)

def main_process(lbl, source_path, dest):
    source_path = standard_path(source_path.get())  
    dest = standard_path(dest.get())  
    workbook = xlrd.open_workbook(source_path)

    list_of_files = folder_content(dest)
    tempsheet = workbook.sheet_by_index(4)
    first_unwriten_day = tempsheet.cell(1, 3).value.split()[0]
    last_e_date = date(first_unwriten_day, -1)
    file_name = []
    for file in list_of_files:
        file_name.append(file.split('.')[0])
    path = ''    
    if(last_e_date in file_name):
        path = dest + '\\' + last_e_date + '.xls'    

    wb = xlwt.Workbook()
    full = True
    try:   
        rb = xlrd.open_workbook(path)
        worksheet = rb.sheet_by_index(0)
        len_rows = worksheet.nrows
        wbook = copy(rb)
        sh = wbook.get_sheet(0)    
        index = len_rows
        
        day_e = int(worksheet.cell(len_rows-1, 1).value.split('-')[2])

   
        if(day_e == 25 ):
            full = True
        else:
            full = False    
    except Exception:
        pass
    writed=False
    for sh in range(4, len(workbook.sheet_names())):
        sheet = workbook.sheet_by_index(sh)
        num_rows = sheet.nrows

        new_id = False
        a_data = [0, 15, 30]
        counter = 0
        for i in [9, 24, 39]:
            new_id=True
            if(full):
                index = 1
            else:
                index = len_rows
            last_date = ''

            for j in range(11, num_rows):   
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
                        if(day == 'شنبه' or day == 'يکشنبه'):
                            start = sheet.cell(j, a_data[counter]  + 10).value
                            if(start == 'غیبت'):
                                start=''
                            end = sheet.cell(j, a_data[counter] + 12).value
                        else:    
                            start = sheet.cell(j, a_data[counter]  + 1).value
                            if(start == 'غیبت'):
                                start=''
                            end = sheet.cell(j, a_data[counter] + 3).value

                        
                        if(not full):
                            sh = wbook.get_sheet(id_cell)
                            sh.write(index, 0, index)
                            sh.write(index, 1, date_)
                            sh.write(index, 2, day)
                            sh.write(index, 3, int(id_cell))
                            sh.write(index, 4, start)
                            sh.write(index, 5, end)
                            index +=1
                            writed = True

                        if(new_id and full):
                            new_id = False
                            w_sheet = wb.add_sheet(id_cell)

                            header_font = xlwt.Font()
                            header_font.bold = True
                            header_style = xlwt.XFStyle()
                            header_style.font = header_font

                            w_sheet.write(0, 0, 'ردیف', header_style)
                            w_sheet.write(0, 1, 'تاریخ', header_style)
                            w_sheet.write(0, 2, 'روز هفته', header_style)
                            w_sheet.write(0, 3, 'شناسه', header_style)
                            w_sheet.write(0, 4, 'ساعت ورود', header_style)
                            w_sheet.write(0, 5, 'ساعت خروج', header_style)                    
                    
                            writed=True

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

    if(writed):
        if(not full):
            if(path != new_path):
                shutil.move(path, new_path)
                path = new_path
            wbook.save(path)
        else:
            wb.save(new_path)
        result = 'file created!'    
        lbl.config(text=result)      

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("blue")  

app = customtkinter.CTk()  
app.title('Excel File')
app.geometry("450x350+500+250")

source_path = customtkinter.StringVar()
dest = customtkinter.StringVar()

def get_source_address():
    txtfld1.delete(0, tk.END)
    filename = filedialog.askopenfilename(initialdir = '/', title = 'Select File', filetype = (('executable, *.xls'), ('all files', '*.*')))
    txtfld1.insert(tk.END, filename)


def get_dest_address():
    txtfld2.delete(0, tk.END)
    directory = filedialog.askdirectory(initialdir = '/', title = 'Select File')
    txtfld2.insert(tk.END, directory)

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

main_process = partial(main_process, lbl3, source_path, dest) 
btn4 = customtkinter.CTkButton(master=app, text="process", command=main_process)
btn4.place(relx=0.8, rely=0.9, anchor=tk.CENTER)

app.mainloop()           