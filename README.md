# General-ledger
from tkinter import *
from tkinter import messagebox
import os
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Font

currentpath = os.getcwd()


Account_head_path = currentpath+"\ledger\Account_Head.xlsx"
account_wb = load_workbook(Account_head_path)
account_ws = account_wb.active
column = account_ws['A']
alias_list = [column[x].value for x in range(len(column))]

def window_atribute():
    screen_width = ledger.winfo_screenwidth()
    screen_height = ledger.winfo_screenheight()
    ledger.title("Ledger")
    ledger.state('zoomed')

def on_keyrelease(event):
    party = Party_entry.get()
    party = party.upper()
    account_wb = load_workbook(Account_head_path)
    account_ws = account_wb.active
    column = account_ws['A']
    alias_list = [column[x].value for x in range(len(column))]
    if party == '':
        data = alias_list
    else:
        data = []
        for item in alias_list:
            if party in item:
                data.append(item)
    listbox_update(data)

def listbox_update(data):

    party_list.delete(0,'end')
    # sorting data
    data = sorted(data)

    # put new data
    for item in data:
        party_list.insert('end',item)

def on_select(event):
     party = event.widget.get('active')

     path = os.getcwd()+ "\ledger\ "+party+".xlsx"
     os.startfile(path)

def create_party():
    party = Party_entry.get()
    party = party.upper()
    account_wb = load_workbook(Account_head_path)
    account_ws = account_wb.active
    Account = 0
    ar = account_ws.max_row
    ac = account_ws.max_column
    print(ar)
    print(ac)
    column = account_ws['A']
    list = [column[x].value for x in range(len(column))]
    for items in list:
        if(items == party or party == ''):
            Account = 1
            break

    if(Account == 0):
        ar = account_ws.max_row
        account_ws.cell(row=ar+1, column=1).value = party
        account_wb.save(Account_head_path)
        path = os.getcwd()+"\ledger\ "+party+".xlsx"
        sample_path = os.getcwd()+"\ledger\SampleSheet.xlsx"
        new = Workbook()
        sample = load_workbook(sample_path)
        new_ws = new.active
        sample_ws = sample.active
        mr = sample_ws.max_row
        mc = sample_ws.max_column

        for i in range(1, mr + 1):
            for j in range(1, mc + 1):
                # reading cell value from source excel file
                new_ws.cell(row=i, column=j).value = sample_ws.cell(row=i, column=j).value

        new_ws.merge_cells('B1:C2')
        new_ws.merge_cells('A1:A2')

        for i in range(1, mr + 1):
            for j in range(1, mc + 1):
                new_ws.cell(row=i, column=j).alignment = Alignment(horizontal='center', vertical='center')

        font_style = Font(bold=True, size='14')
        new_ws.cell(row=1, column=2, value=party).font = font_style

        font_style = Font(bold=True, size='14')
        new_ws.cell(row=1, column=1).font = font_style

        for i in range(1, 3):
            for j in range(4, 8):
                font_style = Font(bold=True, size='12')
                new_ws.cell(row=i, column=j).font = font_style

        new_ws.column_dimensions['B'].width = 10
        new_ws.column_dimensions['C'].width = 40
        new_ws.column_dimensions['D'].width = 15
        new_ws.column_dimensions['E'].width = 15
        new_ws.column_dimensions['F'].width = 20

        new.save(path)
        listbox_update(list)

    else:
        messagebox.showerror("Error", "Party Already Exists")

def statement():
    print("Statement")

def balance():
    print("balance")

def open_ledger():
    party = Party_entry.get()
    party = party.upper()
    account_wb = load_workbook(Account_head_path)
    account_ws = account_wb.active
    column = account_ws['A']
    alias_list = [column[x].value for x in range(len(column))]
    path = os.getcwd()+"\ledger\ "+party +".xlsx"
    os.startfile(path)


def list_select(event):
    print("down")


def ledger_entry(event):
    party = event.widget.get('active')
    path = os.getcwd()+"\ledger\ "+party+".xlsx"
    os.startfile(path)

def single_click(event):
    party = str(event.widget.get('active'))
    Party_entry.delete(0,'end')
    Party_entry.insert(0,party)


ledger = Tk()
window_atribute()  #Title and state of window

canvas1 = Canvas(bg = '#EFE4B0',height = 857 ,width = 1236)
canvas1.pack(side = LEFT , fill=BOTH , expand = True)

canvas2 = Canvas(bg = '#73D5FF',height = 857, width = 300)
canvas2.pack(side = RIGHT,fill = BOTH, expand = False)

lable1 = Label(canvas1,text = "Party Name", bg = '#EFE4B0', font = ('Cambria Math',12))
lable1.grid(row = 1, column = 1, padx  =15, pady =15)

Party_entry = Entry(canvas1 , width  = 60)
Party_entry.grid(row = 1, column = 2, padx  =15)
Party_entry.bind('<KeyRelease>', on_keyrelease)
Party_entry.bind('<Down>',list_select)

party_list = Listbox(canvas1 ,height= 40, width =60)
party_list.grid(row = 2, column = 2 )
party_list.bind('<Double-Button-1>', on_select)
party_list.bind('<Return>',ledger_entry)
party_list.bind('<Button-1>', single_click)
listbox_update(alias_list)

new_party_button = Button(canvas2,text = "Create New Party", command = create_party)
new_party_button.grid(row = 1, column = 1, padx  =100, pady =15)

statement_button = Button(canvas2,text = "Statement", command = statement)
statement_button.grid(row = 2, column = 1, padx  =100, pady =15)

Balance_button = Button(canvas2,text = "Balance", command = balance)
Balance_button.grid(row = 3, column = 1, padx =100, pady =15)

open_ledger_button = Button(canvas2,text = "Open Party ledger", command = open_ledger)
open_ledger_button.grid(row = 4, column = 1, padx  =100, pady =15)

ledger.mainloop()
