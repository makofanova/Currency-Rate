import os
import tkinter as tki
import tkinter.ttk as ttk
import pandas as pd
db_1 = pd.read_excel('C:\Work\Data\БД1.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_2 = pd.read_excel('C:\Work\Data\БД2.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_3 = pd.read_excel('C:\Work\Data\БД3.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_4 = pd.read_excel('C:\Work\Data\БД4.xlsx',sheet_name='Лист 1', index_col=('id_country'))
                
def delete_dataframe_stolb(df, str_l, i):
    for str_ in str_l:
        df = df.drop(str_, i)
    df.to_excel('C:\Work\Output\БД5.xlsx',sheet_name='Лист 1')
           
def delete_dataframe_str(df, ind):
    df = df.drop(db_1[db_1['name_']!=ind].index[0])
    df.to_excel('C:\Work\Output\БД5.xlsx',sheet_name='Лист 1')

def clck_1():
    l_1 = 0
    l_2 = 0
    DATA = []
    DATA_1 =[]
    for k in list(db_3.columns[1:]):
        DATA.append(str(k).split(' ')[0])
        DATA_1.append(k)
    for i in range(len(DATA)):
        if DATA[i]==cmb_.get():
            l_1=i
        if DATA[i]==cmb_2.get():
            l_2=i
    if l_2<l_1:
        c = l_1
        l_1= l_2
        l_2 = c
    DATA_ = DATA_1[:l_1]+DATA_1[l_2:]
    delete_dataframe_stolb(db_3, DATA_, 1)
    
def clck_2():
    i = db_1[db_1['name_']==cmb_3.get()].index[0]
    delete_dataframe_str(db_3, i)
    
def clck_3():
    db_3 = pd.read_excel('C:\Work\Data\БД3.xlsx',sheet_name='Лист 1', index_col=('id_'))
    l_1 = 0
    l_2 = 0
    DATA = []
    DATA_1 =[]
    for k in list(db_3.columns[1:]):
        DATA.append(str(k).split(' ')[0])
        DATA_1.append(k)
    for i in range(len(DATA)):
        if DATA[i]==cmb_.get():
            l_1=i
        if DATA[i]==cmb_2.get():
            l_2=i
    if l_2<l_1:
        c = l_1
        l_1= l_2
        l_2 = c
    DATA_ = DATA_1[:l_1]+DATA_1[l_2:]
    delete_dataframe_stolb(db_3, DATA_, 1)
    db_3 = pd.read_excel('C:\Work\Output\БД5.xlsx',sheet_name='Лист 1', index_col=('id_'))
    i = db_1[db_1['name_']==cmb_3.get()].index[0]
    delete_dataframe_str(db_3, i)
    
def make_combobox_1():
    lbl_2.place(x = 0, y = 20)
    lbl_3.place(x = 0, y = 40)
    cmb_.place(x = 100, y = 20)
    cmb_2.place(x = 100, y = 42)
    if not var_1.get():
        lbl_2.place_forget()
        lbl_3.place_forget()
        cmb_.place_forget()
        cmb_2.place_forget()
        
def make_combobox_2():
    lbl_4.place(x = 0, y = 80)
    cmb_3.place(x = 100, y = 80)
    if not var_2.get():
        lbl_4.place_forget()
        cmb_3.place_forget()
    

def make_combobox_4():
    lbl_2.place(x = 0, y = 120)
    lbl_3.place(x = 0, y = 160)
    lbl_4.place(x = 0, y = 200)
    cmb_.place(x = 100, y = 120)
    cmb_2.place(x = 100, y = 160)
    cmb_3.place(x = 100, y = 200)
    if not var_3.get():
        lbl_2.place_forget()
        lbl_3.place_forget()
        cmb_.place_forget()
        cmb_2.place_forget()
    
win0= tki.Tk()
win0.title('Отбор')
win0.geometry('800x500+300+200')


lbl_2 = tki.Label(win0, text = 'Позднее ', font = ('Times', 12))
lbl_3 = tki.Label(win0, text = 'Ранее', font = ('Times', 12))
lbl_4 = tki.Label(win0, text = 'Валюты кроме ', font = ('Times', 12))
lbl_5 = tki.Label(win0, text = 'Территория обращения ', font = ('Times', 12))
txt = tki.Entry(win0)
txt.focus()
txt_1 = tki.Entry(win0)
txt_1.focus()

var_1 = tki.BooleanVar()
var_1.set(0)
c_1 = tki.Checkbutton(win0, text= 'Отобрать по дате',  font = ('Times', 12), variable=var_1,
                 onvalue=1, offvalue=0, command = make_combobox_1)
c_1.place(x =350, y = 10)

var_2 = tki.BooleanVar()
var_2.set(0)
c_2 = tki.Checkbutton(win0, text= 'Отобрать по валюте',  font = ('Times', 12), variable=var_2,
                 onvalue=1, offvalue=0, command = make_combobox_2)
c_2.place(x =350, y = 80)

cmb_ = ttk.Combobox(win0)
DATA = []
for k in list(db_3.columns[1:]):
    DATA.append(str(k).split(' ')[0])
cmb_['values'] = DATA
cmb_.current(0)
cmb_.focus()
cmb_2 = ttk.Combobox(win0)
cmb_2['values'] = DATA
cmb_2.current(0)
cmb_2.focus()
cmb_3 = ttk.Combobox(win0)
DATA_ = [db_1.values[i][0] for i in range(len(list(db_1.values)))]
cmb_3['values'] = DATA_
cmb_3.current(0)
cmb_3.focus()
cmb_4 = ttk.Combobox(win0)
DATA_1 = [db_4.values[i][0] for i in range(len(list(db_4.values)))]
cmb_4['values'] = DATA_1
cmb_4.current(0)
cmb_4.focus()
'''
i = 0
while list(db_3.columns[1:])[i]!=str_2:
    i+=1
DATA = []
for k in list(db_3.columns[i:]):
    DATA.append(str(k).split(' ')[0])
cmb_['values'] = DATA

'''
var_3 = tki.BooleanVar()
var_3.set(0)
c_3 = tki.Checkbutton(win0, text= 'Отобрать по валюте и дате',  font = ('Times', 12), variable=var_3,
                 onvalue=1, offvalue=0, command = make_combobox_4)
c_3.place(x =350, y = 150)

btn_1 = tki.Button(win0, text = 'Отобрать и сохранить', font = ('Times', 12), command = clck_1)
btn_1.place(x = 350, y = 40)

btn_2 = tki.Button(win0, text = 'Отобрать и сохранить', font = ('Times', 12), command = clck_2)
btn_2.place(x = 350, y = 120)

btn_3 = tki.Button(win0, text = 'Отобрать и сохранить', font = ('Times', 12), command = clck_3)
btn_3.place(x = 350, y = 180)

win0.mainloop()