from tkinter import *
import tkinter.ttk
import openpyxl

excel= openpyxl.load_workbook('a.xlsx')
sheetlist=excel.get_sheet_names()

root= Tk()
#사이즈 title
root.geometry('1200x800')
root.title("Food! Where?!?")

#통로

def callback(selection):
    selected_sheet=excel.get_sheet_by_name(selection)
    drawingsheet(selected_sheet)

def drawingsheet(sheet):
    drawA2(sheet)
    # drawA1(sheet)
    # drawB1(sheet)
    # drawB2(sheet)
    # drawC2(sheet)
    # drawC1(sheet)
    # drawD1(sheet)
    # drawD2(sheet)


def drawA2(sheet):
    print(sheet)
    A2=PanedWindow(root, relief="solid")
    A2.grid(row= 3, column = 0)
    # a2Vrow=Label(A2,text='hi')
    # a2row.grid(row=0,column=0)
    # a2row2=Label(A2, text='me')
    # a2row2.grid(row=1,column=1)
    i=6
    row = f'A{i}'
    while sheet[row].value == "A":
        place= sheet[f'B{i}'].value
        innertext= place
        text=Label(A2,text=innertext)
        text.grid(row=i,column=0)
        i=i+1

mainlabel= Label(root,font=40, height= 5, width=80,  text="Food! Where!?!",relief='solid')
mainlabel.grid(row=0, column=0, columnspan=8)
# stylesheet select
variable=StringVar(root)
variable.set(sheetlist[0])
opt=OptionMenu(root,variable, *sheetlist, command=callback)
opt.config(width= 10)
opt.grid(row=0, column=8)

#a,b,c,d
lbl_A= Label(root, text="A",font=15, width=20, height= 3)
lbl_B= Label(root, text="B",font=15, width=20, height= 3)
lbl_C= Label(root, text="C",font=15, width=20, height= 3)
lbl_D= Label(root, text="D",font=15, width=20, height= 3)
lbl_A.grid(row=1, column=0, columnspan=2)
lbl_space= Label(root, text='', width=10)
lbl_space.grid(row=1,column=2 )
lbl_B.grid(row=1, column=3, columnspan=2)
lbl_C.grid(row=1, column=5, columnspan=2)
lbl_space= Label(root, text='', width=10)
lbl_space.grid(row=1,column=7)
lbl_D.grid(row=1, column=8, columnspan=2)


#a

lbl_a2 = Label(root, text="A2", width=10)
lbl_a2.grid(row= 2, column= 0)

lbl_a1 = Label(root, text="A1",width=10)
lbl_a1.grid(row=2, column= 1)

#space
lbl_space= Label(root, text='', width=10)
lbl_space.grid(row=2, column=2)

#b
lbl_b1 = Label(root, text="B1",width=10)
lbl_b1.grid(row= 2, column= 3)

lbl_b2 = Label(root, text="B2",width=10)
lbl_b2.grid(row= 2, column= 4)

lbl_c2 = Label(root, text="C2",width=10)
lbl_c2.grid(row= 2, column= 5)

lbl_c1 = Label(root, text="C1",width=10)
lbl_c1.grid(row= 2, column= 6)
#space
lbl_space= Label(root, text='', width=10)
lbl_space.grid(row=2, column=7)


#d
lbl_d1 = Label(root, text="D1",width=10)
lbl_d1.grid(row= 2, column= 8)

lbl_d2 = Label(root, text="D2",width=10)
lbl_d2.grid(row= 2, column= 9)

# txt = Entry(root)
# txt.grid(row=2,column=0)

btn = Button(root, text= "ok", width=10)
btn.grid(row=0, column=9)


print()




















root.mainloop()