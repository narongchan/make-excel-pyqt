from tkinter import *
import tkinter.ttk
import openpyxl
from datetime import datetime as dt 
excel= openpyxl.load_workbook('a.xlsx')
sheetlist=excel.sheetnames

root= Tk()
#사이즈 title
root.geometry('1200x800')
root.title("Food! Where?!?")

main_Label= Label(root,font=40, height= 5, width=80,  text="Food! Where!?!",relief='solid')
main_Label.grid(row=0, column=0, columnspan=8)


# 엑셀 시트 선택 
def callback(selection):

     selected_sheet=excel[selection]
     drawingsheet(selected_sheet)
     

# stylesheet select
variable=StringVar(root)
variable.set(sheetlist[0])
opt=OptionMenu(root,variable, *sheetlist, command=callback)
opt.config(width= 10)
opt.grid(row=0, column=8)


#a,b,c,d label
lbl_A= Label(root, text="A",font=15, width=20, height= 3)
lbl_B= Label(root, text="B",font=15, width=20, height= 3)
lbl_C= Label(root, text="C",font=15, width=20, height= 3)
lbl_D= Label(root, text="D",font=15, width=20, height= 3)
lbl_A.grid(row=1, column=0, columnspan=6)
lbl_space= Label(root, text='', width=10)
lbl_space.grid(row=1,column=6, )
lbl_B.grid(row=1, column=7, columnspan=6)
lbl_C.grid(row=1, column=13, columnspan=6)
lbl_space= Label(root, text='', width=10)
lbl_space.grid(row=1,column=19)
lbl_D.grid(row=1, column=20, columnspan=6)
#a
lbl_a2 = Label(root, text="A2", width=10)
lbl_a2.grid(row= 2, column= 0, columnspan= 3)
lbl_a1 = Label(root, text="A1",width=10)
lbl_a1.grid(row=2, column= 3,columnspan=3)
#space
lbl_space= Label(root, text='', width=10)
lbl_space.grid(row=2, column=6)
#b
lbl_b1 = Label(root, text="B1",width=10)
lbl_b1.grid(row= 2, column= 7, columnspan=3)
lbl_b2 = Label(root, text="B2",width=10)
lbl_b2.grid(row= 2, column= 10, columnspan=3)
lbl_c2 = Label(root, text="C2",width=10)
lbl_c2.grid(row= 2, column= 13, columnspan=3)
lbl_c1 = Label(root, text="C1",width=10)
lbl_c1.grid(row= 2, column= 16, columnspan =3)
#space
lbl_space= Label(root, text='', width=10)
lbl_space.grid(row=2, column=19)
#d
lbl_d1 = Label(root, text="D1",width=10)
lbl_d1.grid(row= 2, column=20, columnspan=3)
lbl_d2 = Label(root, text="D2",width=10)
lbl_d2.grid(row= 2, column= 23, columnspan=3)


def drawingsheet(sheet):
    a=3
    b=3
    c=3
    d=3
    i=6
    while sheet[f'a{i}'].value != None:
        lack=sheet[f'a{i}'].value
        num= sheet[f'b{i}'].value
        product=sheet[f'e{i}'].value
        if lack == 'A' and num[3:5]== '02':
            where=Label(root, text= f'{lack}{num}')
            where.grid(row=a-1,column=0)
            product_lbl= Label(root, text= product)
            product_lbl.grid(row=a-1, column=1)
            a=a+1
        elif lack =='A' and num[3:5]== '01':
            where=Label(root, text= f'{lack}{num}')
            where.grid(row=a,column=3)
            product_lbl= Label(root, text= product)
            product_lbl.grid(row=a-1, column=4)
            a=a+1
            # B렉
        if lack == 'B' and num[3:5]== '01':
            where=Label(root, text= f'{lack}{num}')
            where.grid(row=b,column=7)
            b=b+1
        elif lack =='B' and num[3:5]== '02':
            where=Label(root, text= f'{lack}{num}')
            where.grid(row=b-1,column=10)
            b=b+1
            #C렉
        if lack == 'C' and num[3:5]== '02':
            where=Label(root, text= f'{lack}{num}')
            where.grid(row=c-1,column=13)
            c=c+1
        elif lack =='C' and num[3:5]== '01':
            where=Label(root, text= f'{lack}{num}')
            where.grid(row=c,column=16)
            c=c+1
            #D렉
        if lack == 'D' and num[3:5]== '01':
            where=Label(root, text= f'{lack}{num}')
            where.grid(row=d,column=20)
            d=d+1
        elif lack =='D' and num[3:5]== '02':
            where=Label(root, text= f'{lack}{num}')
            where.grid(row=d-1,column=23)
            d=d+1
            

        i=i+1


















root.mainloop()