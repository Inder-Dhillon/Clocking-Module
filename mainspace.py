from tkinter import *
from pandas import DataFrame
from pandas import ExcelWriter
import datetime
import openpyxl
from openpyxl import load_workbook
root = Tk()
root.title("Travel Transport Employee Clock")
root.geometry("600x600")
now = datetime.datetime.now()
writer = ExcelWriter(r'C:\Users\Ferox\Projects\Clock-Module\venv\test.xlsx', engine='openpyxl')
book = load_workbook(r'C:\Users\Ferox\Projects\Clock-Module\venv\test.xlsx')
writer.book = book
writer.sheets = {ws.title: ws for ws in book.worksheets}
check1 = False
check2 = False
check3 = False
#Functions
def enterval():
    global check1
    global check2
    global check3
    text = empcode.get()

    if text == "1111":
        if check1 == False:
            df = DataFrame({'Data': [now.strftime("%Y-%m-%d")], 'Data1': [now.strftime("%H:%M")], 'Data2': "IN"})
            df.to_excel(writer, sheet_name='Dispatch', index=False, startrow=writer.sheets['Dispatch'].max_row, header=False)
            check1 = True
            writer.save()
            empcode.delete(0,END)
            return
        if check1 == True:
            df = DataFrame({'Data': [now.strftime("%Y-%m-%d")], 'Data1': [now.strftime("%H:%M")], 'Data2': "OUT"})
            df.to_excel(writer, sheet_name='Dispatch', index=False, startrow=writer.sheets['Dispatch'].max_row,
                        header=False)
            check1 = False
            writer.save()
            empcode.delete(0, END)
            return

    if text == "6744":
        df1 = DataFrame({'Data': [now.strftime("%Y-%m-%d")], 'Data1': [now.strftime("%H:%M")]})
        df1.to_excel(writer, sheet_name='Inder', index=False, startrow=writer.sheets['Inder'].max_row,header=False)
        writer.save()
    if text == "2222":
        df2 = DataFrame({'Data': [now.strftime("%Y-%m-%d")], 'Data1': [now.strftime("%H:%M")]})
        df2.to_excel(writer, sheet_name='Safety', index=False, startrow=writer.sheets['Safety'].max_row,header=False)
        writer.save()
    else:
        return



def set_text(text):
    empcode.insert(END,text)
    return

#GUI
root["bg"] = "white"
Label(root, text="Code:", font="Roboto", bg="white").grid(row=1, column=2, sticky=NSEW)
empcode = Entry(root, justify=CENTER, bg="#dcdcdc", width=80, borderwidth=0)
empcode.grid(row=2,column=1, sticky=NSEW, columnspan=3,pady=10)
#Good way might work on it later
#numbers = [['1', '2', '3'],\
#        ['4', '5', '6'],\
#         ['7', '8', '9']]
#for i, (x, y, z) in enumerate(numbers):
#    Button(root, text=x, command=lambda:set_text(x)).grid(row=i+3, column=1, ipadx=60, ipady=30, sticky=NSEW)
#    Button(root, text=y, command=lambda:set_text(y)).grid(row=i+3, column=2, ipadx=60, ipady=30,sticky=NSEW)
#    Button(root, text=z, command=lambda:set_text(z)).grid(row=i+3, column=3, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="1", command=lambda:set_text("1")).grid(row=3, column=1, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="2", command=lambda:set_text("2")).grid(row=3, column=2, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="3", command=lambda:set_text("3")).grid(row=3, column=3, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="4", command=lambda:set_text("4")).grid(row=4, column=1, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="5", command=lambda:set_text("5")).grid(row=4, column=2, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="6", command=lambda:set_text("6")).grid(row=4, column=3, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="7", command=lambda:set_text("7")).grid(row=5, column=1, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="8", command=lambda:set_text("8")).grid(row=5, column=2, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="9", command=lambda:set_text("9")).grid(row=5, column=3, ipadx=60, ipady=30, sticky=NSEW)
Button(root, text="Enter ->", command=enterval).grid(row=6, column=2, ipadx=60, ipady=30,sticky=NSEW, pady=20)
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(7, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(4, weight=1)
root.mainloop()