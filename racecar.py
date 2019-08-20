
# Following Project is a BOQ Builder ie Bill of Quantities
# Project Being Developed is for Citrix Technologies by Aryan Shetty

import tkinter as tk
from tkinter import *
from tk import *

#import pyinstaller
import xlrd
import xlwt
from xlwt import Workbook

# Give the location of the file

import os
import sys


config_name = 'inputfile.xlsx'

# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

config_path = os.path.join(application_path, config_name)

loc = (config_path)

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# To collect input
# reqcol=input("Enter column name wanted")

# To get number of rows and columns
rowsn=sheet.nrows
colsn=sheet.ncols

# print(rowsn)
# print(colsn)

# Workbook is created
global wb2
wb2 = Workbook()

# add_sheet is used to create sheet.
global sheet1
sheet1 = wb2.add_sheet('Sheet 1')



# To create a list that will check if any columns are blank
yn_list=[0]*colsn
counter=0
second_counter=0

sheet1.write(0,0,"SKU")
sheet1.write(0,1,"Description")
sheet1.write(0,2,"SRP")
sheet1.write(0,3,"Quantity")
sheet1.write(0,4,"Total")



# To get unique values from the columns
line_unique = []
for x in range(2,rowsn):
    if sheet.cell_value(x,3) not in line_unique:
        line_unique.append(sheet.cell_value(x,3))

model_unique = []
for x in range(2,rowsn):
    if sheet.cell_value(x,4) not in model_unique:
        model_unique.append(sheet.cell_value(x,4))

series_unique=[]
for x in range(2,rowsn):
    if sheet.cell_value(x,5) not in series_unique:
        series_unique.append(sheet.cell_value(x,5))

edition_unique=[]
for x in range(2,rowsn):
    if sheet.cell_value(x,6) not in edition_unique:
        edition_unique.append(sheet.cell_value(x,6))

type_unique=[]
for x in range(2,rowsn):
    if sheet.cell_value(x,13) not in type_unique:
        type_unique.append(sheet.cell_value(x,13))
print ("|||||||||||||||||||||||||||")


#########################################################

# GUI STUFF BEGINS

# To create demo tkinter button and dropdown button
r =Tk()
r.title('BOQ Builder')
OPTIONS = [
"Jan",
"Feb",
"Mar"
] #etc

frame = Frame(r)
frame.grid(row=0,column=0,sticky=(N, W, E, S))

# Label(r, text="First").grid(row=6)
# Label(r, text="Second").grid(row=7)


# To write stuff
# w = Message(r, text="this is a message")
# w.pack(padx=5, pady=10, side=tk.LEFT)
# topframe=Frame(r)
# topframe.pack(side=TOP)
# bottomframe = Frame(r)
# bottomframe.pack(side=BOTTOM)

# To display text above dropdown


# label.pack(side=LEFT)
# To create first dropdown
var = StringVar()
label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=0,row=0)

var.set("Product Line")
line_var = StringVar(r)
line_var.set(line_unique[0]) # default value

w = OptionMenu(frame, line_var, *line_unique).grid(column=0,row=1,pady=5,padx=5)
# w.pack(side = LEFT)

# To create frame to organize widgets


# To create second dropdown
var = StringVar()
label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=1,row=0,pady=5,padx=5)

var.set("Product Model")
model_var = StringVar(r)
model_var.set(model_unique[0]) # default value

w = OptionMenu(frame, model_var, *model_unique).grid(column=1,row=1,pady=5,padx=5)
# w.pack(side = LEFT)

# To create third dropdown
var = StringVar()
label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=2,row=0,pady=5,padx=5)

var.set("Series")
series_var = StringVar(r)
series_var.set(series_unique[0]) # default value

w = OptionMenu(frame, series_var, *series_unique).grid(column=2,row=1,pady=5,padx=5)
# w.pack(side = LEFT)

# To create fourth dropdown
var = StringVar()
label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=3,row=0,pady=5,padx=5)

var.set("Edition")
edition_var = StringVar(r)
edition_var.set(edition_unique[0]) # default value

w = OptionMenu(frame, edition_var, *edition_unique).grid(column=3,row=1,pady=5,padx=5)
# w.pack(side = LEFT)

# To create fifth dropdown
var = StringVar()
label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20) ).grid(column=4,row=0,pady=5,padx=5)

var.set("Type")
type_var = StringVar(r)
type_var.set(type_unique[0]) # default value

w = OptionMenu(frame, type_var, *type_unique).grid(column=4,row=1,pady=5,padx=5)
# w.pack(side = LEFT)

# god=Label(r,text="GODMODE").grid(column=5,row=1)



def ddg():
    print("line is:", line_var.get())
    print("model is:", model_var.get())
    print("series is:", series_var.get())
    print("edition is:", edition_var.get())
    print("type is:", type_var.get())




    var = StringVar()
    label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=0, row=4,pady=5,padx=5)
    var.set("SKU")

    var = StringVar()
    label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=1, row=4,pady=5,padx=5)
    var.set("Description")

    var = StringVar()
    label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=2, row=4,pady=5,padx=5)
    var.set("SRP")

    var = StringVar()
    label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=3, row=4,pady=5,padx=5)
    var.set("Quantity")


    yoyo = 0
    global list_q
    list_q=[]
    global list_s
    list_s=[]
    entries=[]
    for i in range(rowsn):
        if (sheet.cell_value(i, 3) ==line_var.get()) and (sheet.cell_value(i, 4) == model_var.get()) and (
                (sheet.cell_value(i, 5)) == float(series_var.get())) \
                and (sheet.cell_value(i, 6) == edition_var.get()) and (sheet.cell_value(i, 13) == type_var.get()):
            yoyo += 1
            print ("yoyo is",yoyo)
            # To write into excel file


            # shiit.write(yoyo, 4, sheet.cell_value(i, 17)*int(quantity))
            # To write SKU
            var = StringVar()
            label = Label(frame, textvariable=var,borderwidth=2, relief="groove",font=("",15)).grid(column=0, row=yoyo+4,pady=5,padx=5)
            var.set(sheet.cell_value(i,2))

            # To write Description
            var = StringVar()
            label = Label(frame, textvariable=var,justify=RIGHT,borderwidth=2, relief="groove",font=("",15)).grid(column=1, row=yoyo + 4, pady=5,padx=5)
            var.set(sheet.cell_value(i, 1))
            # To write SRP
            var = StringVar()
            label = Label(frame, textvariable=var,borderwidth=2, relief="groove",font=("",15)).grid(column=2, row=yoyo + 4,pady=5,padx=5)
            var.set(sheet.cell_value(i, 17))
            list_s.append(sheet.cell_value(i, 17))
            # To write
            qty_entered = StringVar()
            E1 = Entry(frame, textvariable=qty_entered, bd=2)
            E1.grid(column=3,columnspan=1,row=yoyo+4, sticky=E+W+S+N,pady=10,padx=5)
            entries.append(E1)
            list_q.append(qty_entered.get())

            # E1.place(x=500,y=500)

        else:
            continue



    def booyah():
        jase=0
        wgod = Workbook()
        global shiit
        shiit = wgod.add_sheet("shiit")
        yoyo=0
        shiit.write(0, 0, "SKU")
        shiit.write(0, 1, "Description")
        shiit.write(0, 2, "SRP")
        shiit.write(0, 3, "Quantity")
        shiit.write(0, 4, "Total")
        for i in range(rowsn):
            if (sheet.cell_value(i, 3) == line_var.get()) and (sheet.cell_value(i, 4) == model_var.get()) and (
                    (sheet.cell_value(i, 5)) == float(series_var.get())) \
                    and (sheet.cell_value(i, 6) == edition_var.get()) and (sheet.cell_value(i, 13) == type_var.get()):
                yoyo += 1
                print("yoyo is", yoyo)
                # To write into excel file
                shiit.write(yoyo, 0, sheet.cell_value(i, 2))
                shiit.write(yoyo, 1, sheet.cell_value(i, 1))
                shiit.write(yoyo, 2, sheet.cell_value(i, 17))
        shiit.write(yoyo + 2, 3, "Grand Total")
        var = StringVar()
        label = Label(frame, textvariable=var, borderwidth=2, relief="solid",font=("",20)).grid(column=4, row=4)
        var.set("Total")
        counter=0
        sum=0
        new_counter=yoyo

        for entry in (entries):
            # jase+=2
            # print (jase)
            shiit.write(new_counter-1,3,entry.get())
            shiit.write(new_counter - 1, 4,list_s[counter] * int((entry.get())))

            new_counter+=1
            var = StringVar()
            label = Label(frame, textvariable=var).grid(column=4, row=5+ counter)
            var.set(list_s[counter] * int((entry.get())))

            sum += list_s[counter] * int((entry.get()))
            counter+=1

        var = StringVar()
        label = Label(frame, textvariable=var, relief=SOLID,borderwidth=2,fg="green",font=("",20)).grid(column=3, row=yoyo + 4 + counter)
        var.set("Grand Total")
        var = StringVar()
        label = Label(frame, textvariable=var).grid(column=4, row=yoyo + 4 + counter)
        var.set(sum)
        shiit.write(new_counter,4,sum)
        wgod.save("skrrt22.xls")
    test=Button(frame,text="Calculate Total",command=booyah,fg="blue",relief = RAISED,font=("",18)).grid(column=1,row=yoyo+5)



# def tgen_func(qty_entered):
#     print ("please work")
#     var = StringVar()
#     label = Label(frame, textvariable=var, relief=RAISED).grid(column=4, row=4)
#     var.set("Total")
#     # qtext=qty_entered.get()
#     print (qty_entered)
#     return "hi"
butgo = Button(frame, text="Generate BOQ", command=ddg, fg="blue",relief=RAISED,font=("",18)).grid(row=2, column=1,pady=10,padx=5)

wb2.save('big_test_29.xls')

# tgen = Button(frame, text="Calculate Total", command=tgen_func(5))
r.mainloop()





# To create demo dropdown list
#
# OPTIONS = [
# "Jan",
# "Feb",
# "Mar"
# ] #etc
#
# master = Tk()
#
# variable = StringVar(master)
# variable.set(OPTIONS[0]) # default value
#
# w = OptionMenu(master, variable, *OPTIONS)
# w.pack()
#
# def ok():
#     print ("value is:", variable.get())
#
# button = Button(master, text="OK", command=ok)
# button.pack()
#
# mainloop()


#pyinstaller ("racecar.py")

# To add textbox entry widget

# name=tk.StringVar()
# name_entered=Entry(r,width=50,textvariable=name)
# name_entered.grid(column=0,row=1)
#
# # To add button
# action=Button(r,text="Click Me!",command=ddg())
# action.grid(column=2,row=1)
#
# Label(r,text="Choose a number:").grid(column=1,row=0)
# number=tk.StringVar
# number_chosen=OptionMenu(r,width=50,textvariable=number)
# number_chosen['values']=(1,2,4,42,100)
# number_chosen.grid(column=1,row=1)
# number_chosen.current(0)
# r.mainloop()


#      sheet1.write(0, 0, "SKU")
    # sheet1.write(0, 1, "Description")
    # sheet1.write(0, 2, "Quantity")
    # sheet1.write(0, 3, "SRP")
    # sheet1.write(0, 4, "Total")

    # To make a table of output



    # for x  in range (rowsn):
    #     if (sheet.cell_value(x, 3) == line_var.get()) and (sheet.cell_value(x, 4) == model_var.get()) and (
    #             (sheet.cell_value(x, 5)) == (series_var.get())) \
    #             and (sheet.cell_value(x, 6) == edition_var.get()) and (sheet.cell_value(x, 13) == type_var.get()):
    #         print ("Worked")

    # height = rowsn
    # width = colsn
    # for i in range(height):
    #     if (sheet.cell_value(i, 3) == line_var.get()) and (sheet.cell_value(i, 4) == model_var.get()) and (
    #             (sheet.cell_value(i, 5)) == (series_var.get())) \
    #             and (sheet.cell_value(i, 6) == edition_var.get()) and (sheet.cell_value(i, 13) == type_var.get()):
    #         yoyo+=1
    #         print("reached here")
    #         var = StringVar()
    #         label = Label(frame, textvariable=var).grid(column=yoyo+2, row=0)
    #         print ("this is the value of",sheet.cell_value(i,2))
    #         var.set(sheet.cell_value(i,2))
    #         # hope=Text(r,sheet.cell_value(i,2) ).grid(column =5, row =5 )
    #
    #     # for j in range(width):  # Columns
    #     #
    #
    # # mainloop()

# for i in range(rowsn-1000):
#     if (sheet.cell_value(i,3)==prod_line) and (sheet.cell_value(i,4)==prod_model) and ((sheet.cell_value(i,5))==int(series))\
#             and (sheet.cell_value(i,6)==edition) and (sheet.cell_value(i, 13) == prog_type):
#
#         yoyo+=1
#         #and (sheet.cell_value(i,4)==prod_model) and (sheet.cell_value(i,5)==series) and (sheet.cell_value(i,6)==edition) \
#         #   and (sheet.cell_value(i, 13) == prog_type)
#
#         sheet1.write(yoyo,0,sheet.cell_value(i,2))
#         sheet1.write(yoyo,1,sheet.cell_value(i,1))
#         sheet1.write(yoyo, 2, quantity)
#         sheet1.write(yoyo, 3, sheet.cell_value(i, 17))
#         sheet1.write(yoyo, 4, sheet.cell_value(i, 17)*int(quantity))
#     else:
#         continue


# # To write (or not write) the column into output file
# if (material_input.lower()=="y"):
#     second_counter+=1
#     yn_list[counter]="y"
#     for l in range(0,colsn):
#         if (sheet.cell_value(1,l).lower()=="material"):
#             identifier=l
#             if (yn_list[counter-1]=="n"):
#                 for a in range(0,rowsn):
#                     sheet1.write(a,counter-1,sheet.cell_value(a,identifier))
#             else:
#                 for a in range(0,rowsn):
#                     sheet1.write(a,counter,sheet.cell_value(a,identifier))
# else:
#     yn_list[counter]="n"
#
# counter+=1
# print ("counter after first iteration:",  counter)
# print ("second counter after first iteration", second_counter)
#
# if (sku_input.lower()=="y"):
#     yn_list[counter] = "y"
#     second_counter+=1
#     for l in range(0,colsn):
#         if (sheet.cell_value(1,l).lower()=="sku long description"):
#             identifier=l
#             if (yn_list[counter-1]=="n"):
#                 for a in range(0,rowsn):
#                     sheet1.write(a,counter-1,sheet.cell_value(a,identifier))
#             else:
#                 for a in range(0,rowsn):
#                     sheet1.write(a,counter,sheet.cell_value(a,identifier))
# else:
#     yn_list[counter]="n"
#     second_counter-=1
#
# counter+=1
# print ("counter after second iteration:",  counter)
# print ("second counter after second iteration", second_counter)
# if (ext_sku_input.lower()=="y"):
#     yn_list[counter] = "y"
#     second_counter+=1
#     for l in range(0,colsn):
#         if (sheet.cell_value(1,l).lower()=="ext sku"):
#             identifier=l
#             if (yn_list[counter-1]=="n"):
#                 for a in range(0,rowsn):
#                     sheet1.write(a,counter-1,sheet.cell_value(a,identifier))
#             else:
#                 for a in range(0,rowsn):
#                     sheet1.write(a,counter,sheet.cell_value(a,identifier))
# else:
#     yn_list[counter]="n"
#     second_counter-=1
# counter+=1
# print ("counter after third iteration:",  counter)
# print ("second counter after third iteration", second_counter)
# if (product_line_input.lower()=="y"):
#     yn_list[counter] = "y"
#     second_counter+=1
#     for l in range(0,colsn):
#         if (sheet.cell_value(1,l).lower()=="product line"):
#             identifier=l
#             if (yn_list[counter-1]=="n"):
#                 for a in range(0,rowsn):
#                     sheet1.write(a,counter-1,sheet.cell_value(a,identifier))
#             else:
#                 for a in range(0,rowsn):
#                     sheet1.write(a,counter,sheet.cell_value(a,identifier))
# else:
#     yn_list[counter]="n"
#     second_counter-=1
# print ("counter after fourth iteration:",  counter)
# print ("second counter after fourth iteration", second_counter)



# if (sheet.cell_value(0,0)==''):
#     print("value was none")

# cell_val= sheet.cell(0, 0).value
# print (sheet.cell(0, 0).value)


# To test if script is stopped

# To print wanted column
# for i in range(0,colsn):
#
#     if (sheet.cell_value(1,i).lower()==reqcol.lower()):
#         print ("here wtaf")
#         print (i)
#         print(sheet.col(i))
#         print (type(sheet.col(i)))
#         q=i

# To test list
# for j in (sheet.col(q)):
#     print ("this is j", j, end=' '),

# To copy one sheet to another
# for h in range(0,rowsn):
#     for q in range(0,colsn):
#         test_value=sheet.cell_value(h,q)
#         sheet1.write(h,q,sheet.cell_value(h,q))