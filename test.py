from tkinter import *
import tkinter.ttk
import openpyxl
from datetime import datetime as dt
excel= openpyxl.load_workbook('a.xlsx')
sheetlist=excel.sheetnames

root= Tk()
root.title("푸드뱅크 창고 현황")

    
main_label=Label(root, width= 80, text="푸드뱅크 창고 현황")
main_label.grid(row=0, column=0, columnspan= 8)


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

#draw a2,a1,b1,b2~d2
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
lbl_D1 = Label(root, text="D1",width=10)
lbl_D1.grid(row= 2, column= 8)

lbl_D2 = Label(root, text="D2",width=10)
lbl_D2.grid(row= 2, column= 9)

A2=PanedWindow(root, relief="solid")
A1=PanedWindow(root, relief="solid")
B1=PanedWindow(root, relief="solid")
B2=PanedWindow(root, relief="solid")
C2=PanedWindow(root, relief="solid")
C1=PanedWindow(root, relief="solid")
D1=PanedWindow(root, relief="solid")
D2=PanedWindow(root, relief="solid")
A2.grid(row=3, column=0)
A1.grid(row=3, column=1)
B1.grid(row=3, column=3)
B2.grid(row=3, column=4)
C2.grid(row=3, column=5)
C1.grid(row=3, column=6)
D1.grid(row=3, column=8)
D2.grid(row=3, column=9)


lbl_test=Label(A2, text= "test" )
lbl_test.grid(row=0,column=0)


def draw_A2(sheet):
    i=6
    while sheet[f'A{i}'].value =='A':
        place= sheet[f'b{i}'].value
        product= sheet[f'E{i}'].value
        if place[3:5]=='02':
            lbl_place=Label(A2,text= place)
            lbl_place.grid(row=i-5, column=0)
        
        i+=1


def callback(selection):
    selected_sheet=excel[selection]
    control_draw()


def control_draw():
    if A2.children != None:
        print('hi')



















variable=StringVar(root)
variable.set(sheetlist[0])
opt=OptionMenu(root,variable, *sheetlist, command=callback)
opt.config(width= 10)
opt.grid(row=0, column=8)






root.mainloop()