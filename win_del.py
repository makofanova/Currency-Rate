import pandas as pd
import tkinter as tki
import tkinter.ttk as ttk
db_1 = pd.read_excel('C:\Work\Data\БД1.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_2 = pd.read_excel('C:\Work\Data\БД2.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_3 = pd.read_excel('C:\Work\Data\БД3.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_4 = pd.read_excel('C:\Work\Data\БД4.xlsx',sheet_name='Лист 1', index_col=('id_country'))

win1= tki.Tk()
win1.title('Удаление валюты из базы данных')
win1.geometry('600x400+300+200')
var_3 = []
list_id =[]
c = []
len_ = len(db_1.values)
l=0
for i in range(len_):
    var_3.append(tki.BooleanVar())
    k = var_3[i]
    k.set(0)
    d = tki.Checkbutton(win1, text= db_1.values[i][0],  font = ('Times', 12), variable=k,
                 onvalue=1, offvalue=0)#, command = del_)
    c.append(d)
    if (i< len_//2+1):
        c[i].place(x =0, y = 0+25*l)
        l = i+1
    else:
        if l==i:
            l=0
        else:
            l+=1
        c[i].place(x = 300, y = 0+25*l)
btn_6 = tki.Button(win1, text = "Удалить", font = ('Times', 12), command = win1.destroy)
btn_6.place(x = 340, y = 25*(l+1), width = 100, height = 35 )

win1.mainloop()

for i in range(len(var_3)):
    if var_3[i].get():
        print(db_1.values[i])
        list_id.append(db_1.values[i][0])
df_1 = db_1
df_2 = db_2
df_3 = db_3
df_4 = db_4

for i in list_id:
    df_1 = df_1.drop(db_1[db_1['name_']==i].index[0])
    df_3 = df_3.drop(db_1[db_1['name_']==i].index[0])
    
#writer  = pd.ExcelWriter('C:\Work\Data\БД.xlsx')
df_4.to_excel('C:\Work\Data\БД4.xlsx', sheet_name='Лист 1')

df_1.to_excel('C:\Work\Data\БД1.xlsx',sheet_name='Лист 1')

df_2.to_excel('C:\Work\Data\БД2.xlsx', sheet_name='Лист 1')

df_3.to_excel('C:\Work\Data\БД3.xlsx', sheet_name='Лист 1')





