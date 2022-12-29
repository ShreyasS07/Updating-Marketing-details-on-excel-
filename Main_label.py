import json
import xlrd
from xlutils.copy import copy
import xlwt
import pandas as pd
from xlrd import open_workbook
from tkinter import *
from tkinter import ttk
import tkinter as tk

from tkinter import ttk, StringVar, filedialog, messagebox
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
from tkinter.filedialog import askopenfilename, askdirectory, askopenfilenames, asksaveasfilename, askopenfile
texttoshow = None


select_raw_file = None
def select_raw_excel1():
    global select_raw_file
    select_raw_file = askopenfilename(filetypes=[('Excel Files', '*.xls;*.xlsx;')])
    return select_raw_file if bool(select_raw_file) else None

select_order_file = None
def select_order_excel1():
    global select_order_file
    select_order_file = askopenfilename(filetypes=[('Excel Files', '*.xls;*.xlsx;')])
    return select_order_file if bool(select_order_file) else None

def process():
    count =0
    global texttoshow
    global select_raw_file,select_order_file, row, order_component
# if select_raw_file and select_order_file == None:
    order_file= select_order_file.split("/")[-1].split( )
    print(order_file,'order file')
    raw_file = select_raw_file.split("/")[-1].split()
    print(raw_file,"raw file")
    if "LLB" in raw_file:
# -----checking the order file is Label-Main or not------#
        if "MAIN" in order_file:
            raw_sheet_name = []
            order_sheet_name = []
            rawf = xlrd.open_workbook(select_raw_file)
            for sht in rawf.sheets():
                print(sht.name,"name--",sht.sheet_selected,"selected sheet--",sht.sheet_visible,"visible sheet====")
                if sht.sheet_visible == 1:
                    raw_sheet_name.append(sht.name)
            orderf = xlrd.open_workbook(select_order_file)
            for sht in orderf.sheets():
                count +=1
                print(count,sht.name, "name--", sht.sheet_selected, "selected sheet--", sht.sheet_visible,
                      "visible sheet====")
                if sht.sheet_visible == 1:
                    order_sheet_name.append(sht.name)
                    order_sheet_name.append(count-1)
            print(raw_sheet_name[0],"raw",order_sheet_name[0],"order",order_sheet_name[1],"index number")
            df = pd.read_excel(select_raw_file, sheet_name=raw_sheet_name[0])
            df1 = pd.read_excel(select_order_file, sheet_name=order_sheet_name[0])
            out = df1[26:]
            df = df[df['MAT CLASSF'] == 'Label - Main']
            label = []
            for index, row in out.iterrows():
                print(out.loc[index, 'LL BEAN'])
                val = out.loc[index, 'LL BEAN']

                if str(val) == 'nan':
                    break
                else:
                    ddd = {
                        "index": index,
                        "value": val
                    }
                    label.append(ddd)
            print(label)

            quantity_value = []
            for iX in label:
                i = iX['index']
                v = iX['value']
                df4_new = df[df['REF NO'].str.contains(str(v), na=False)]
                if ~ df4_new.empty:
                    # print(i,df4_new['QTY'].sum(axis=0))
                    qty = df4_new['QTY'].sum(axis=0)
                    if qty != 0:
                        ddd = {
                            "row": i + 1,
                            "quantity": qty
                        }
                        quantity_value.append(ddd)
            print(quantity_value)
            file = str(select_order_file)
            print(file)
            workbook = xlrd.open_workbook(file, formatting_info=True)
            sheet = order_sheet_name[0]
            # sheet = workbook.sheet_by_index(4)
            print(sheet, "11111")
            wb = copy(workbook)
            new_sheet = wb.get_sheet(order_sheet_name[1])
            # new_sheet = wb.sheet_by_index(4)
            print(new_sheet, "22222")
            for qv in quantity_value:
                col = 6
                row_ix = qv['row']
                Qty = qv['quantity']
                print(row_ix, col, Qty)
                # s.cell(row, col).value
                new_sheet.write(row_ix, col, Qty)
                print("===================")
            wb.save(file)
            texttoshow.set('Main-Label TASK COMPLETED ')

        else:
            print("Not the correct Main Label file you selected")

def gi_start():
    global texttoshow
    window = tk.Tk()
    window.geometry("300x200")
    window.resizable(0, 0)
    window.configure(background='white')
    title = "LLB"
    window.title(title)
    tk.Button(window, text="Select Raw File", font=('Helvetica', 9), bg="white", command=select_raw_excel1, width=30).pack()
    tk.Button(window, text="Select Order Form", font=('Helvetica', 9), bg="white", command=select_order_excel1, width=30).pack()
    tk.Button(window, text="Submit", font=('Helvetica', 9), bg="white", command=process, width=30).pack()
    texttoshow = StringVar()
    tk.Label(window, textvariable=texttoshow, wraplength=350, bg="white", font=('Helvetica', 8)).pack(pady=2)
    texttoshow.set('')
    window.mainloop()
gi_start()