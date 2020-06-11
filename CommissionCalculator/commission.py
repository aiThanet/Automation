import pyodbc 
import openpyxl
import csv
import os
import sys
import tkinter
from tkinter import messagebox
from openpyxl.styles import colors
from openpyxl.styles import Font, Color

root = tkinter.Tk()

def scan_bill(filename='scan.xlsx'):
    scan_wb = openpyxl.load_workbook(filename=filename)
    scan_sheet = scan_wb['Sheet1']
    row_scan = 1

    bills = set()
    while True:
        if scan_sheet.cell(row=row_scan, column=1).value:
            bills.add(scan_sheet.cell(row=row_scan, column=1).value)
        else:
            break
        row_scan += 1
    return bills


def get_data(bills,server='127.0.0.1',database='database',username='uid',password = 'password',rate_engine=4,rate_part=6):

    output = {}
    config = 'DRIVER={SQL Server};' + f'SERVER={server};DATABASE={database};UID={username};PWD={password}'
    cnxn = pyodbc.connect(config)
    cursor = cnxn.cursor()

    for bill in bills:
        query_1 = '''
            SELECT *
            FROM SOInvHD INNER JOIN EMCust ON (SOInvHD.CustID=EMCust.CustID)
            WHERE DocuNo=?
            '''
        cursor.execute(query_1,bill)
        for row in cursor.fetchall():
            amt_engine = 0
            amt_part = 0
            amt_total = 0

            query_2 = '''
            SELECT *
            FROM SOInvDT INNER JOIN EMGood ON (SOInvDT.GoodID=EMGood.GoodID)
            WHERE SOInvID=?
            '''
            cursor.execute(query_2, row.SOInvID)
            for row2 in cursor.fetchall():
        
                amt_total += row2.GoodAmnt 
                if row2.GoodGroupID == 1000:
                    amt_engine += row2.GoodAmnt
                elif row2.GoodGroupID == 1001:
                    amt_part += row2.GoodAmnt

            if row.CustName not in output:
                output[row.CustName] = {}
            if amt_engine > 0:
                output[row.CustName][bill + '_' + str(rate_engine)] = float(amt_engine)
            if amt_part > 0:
                output[row.CustName][bill + '_' + str(rate_part)] = float(amt_part)
    return output


def write_output(output, dest_filename='output.xlsx'):
    output_wb = openpyxl.Workbook()
    output_worksheet = output_wb.active

    col_name = ['ชื่อร้าน', 'เลขที่บิล(BL)','ยอดบิล','ธนาคาร','เลขที่เช็ค','เช็คผ่าน','ยอดเช็ค','เปอร์เซ็นต์','ยอดคอม']

    for i,col in enumerate(col_name):
        output_worksheet.cell(row=1, column=i+1).value = col

    curr_row = 2
    for customer in sorted(output):
        output_worksheet.cell(row=curr_row, column=1).value = customer
        total_amount = 0
        total_commission = 0
        for bill in sorted(output[customer]):
            bill_code, percent = bill.split("_")
            commission = round(output[customer][bill] * (float(percent)/100),2)

            output_worksheet.cell(row=curr_row, column=2).value = bill_code
            output_worksheet.cell(row=curr_row, column=3).value = output[customer][bill]
            output_worksheet.cell(row=curr_row, column=8).value = percent + "%"
            output_worksheet.cell(row=curr_row, column=9).value = commission

            total_amount += output[customer][bill]
            total_commission += commission

            curr_row += 1

        
        output_worksheet.cell(row=curr_row, column=2).value = 'รวม'
        output_worksheet.cell(row=curr_row, column=2).font = Font(bold=True)
        output_worksheet.cell(row=curr_row, column=3).value = total_amount
        output_worksheet.cell(row=curr_row, column=3).font = Font(bold=True)
        output_worksheet.cell(row=curr_row, column=9).value = total_commission
        output_worksheet.cell(row=curr_row, column=9).font = Font(bold=True)
        curr_row += 1


    output_wb.save(filename = dest_filename)

def run(e):
    try:
        global root
        bills = set([bill.upper() for bill in e['bills'].get("1.0","end-1c").split('\n')])

        dest_filename='output.xlsx'
        rate_engine = float(e['เครื่องยนต์'].get())
        rate_part = float(e['อะไหล่'].get())

        # bills = scan_bill()
        output = get_data(bills,server="localhost",rate_engine=rate_engine,rate_part=rate_part)
        write_output(output,dest_filename=dest_filename)
        
        os.system(f'start excel {dest_filename}')

        root.destroy()

    except ValueError:
        tkinter.messagebox.showerror("Error", "เกิดข้อผิดพลาด")



def makeform(root, fields=['เครื่องยนต์','อะไหล่']):
    entries = {}
    row = tkinter.Frame(root)
    row.pack(side=tkinter.TOP, fill=tkinter.X, padx=5)
    lab = tkinter.Label(row, text='ใส่เลขที่บิล', anchor='w',font=("Courier", 20))
    lab.pack(side=tkinter.TOP)
    
    row = tkinter.Frame(root)
    row.pack(side=tkinter.TOP, fill=tkinter.X, padx=5, pady=5)
    text_fleid = tkinter.Text(row,height=20,width=40)
    scrollbar = tkinter.Scrollbar(row, command=text_fleid.yview)
    text_fleid.configure(yscrollcommand=scrollbar.set)
    text_fleid.pack(side=tkinter.LEFT,padx=5, pady=5)
    scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    entries['bills'] = text_fleid
  
    for field in fields:
        row = tkinter.Frame(root)
        lab = tkinter.Label(row, width=10, text=field, anchor='w',font=("Courier", 20))
        ent = tkinter.Entry(row)
        row.pack(side=tkinter.TOP, fill=tkinter.X, padx=5, pady=5)
        lab.pack(side=tkinter.LEFT)
        ent.pack(side=tkinter.RIGHT, expand=tkinter.YES, fill=tkinter.X)
        entries[field] = ent
    return entries


def main():
    global root
    
    root.title("Commission")
    root.minsize(100,100)

    ents = makeform(root)

    ok_btn = tkinter.Button(root,text='OK',font=("Courier", 20),command=(lambda e=ents: run(e)))
    ok_btn.pack(side=tkinter.TOP, padx=5, pady=5)

    root.mainloop()
    

if __name__ == '__main__':
    main()

# '''
# gid
# 1000 เครื่องยน
# 1001 อะไหล่
# 1002 ทั่วไป
# 2000 ตลับเมตร
# 4000 สมาคุณ
# 4001 มัเทำครัว
# '''
