import pandas as pd
import tkinter as tki
import tkinter.ttk as ttk
db_1 = pd.read_excel('C:\Work\Data\БД1.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_2 = pd.read_excel('C:\Work\Data\БД2.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_3 = pd.read_excel('C:\Work\Data\БД3.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_4 = pd.read_excel('C:\Work\Data\БД4.xlsx',sheet_name='Лист 1', index_col=('id_country'))


def clck_1():
    colmn = [i for i in list(db_3.columns) if str(i).split(' ')[0]==cmb_4.get()]
    print(colmn)
    df1=pd.DataFrame({colmn[0]: entry_2.get()}, index=[int(db_1.index[-1])+1])#cmb_4.get() entry_2.get(), entry_1.get()]},index=[db_1.index[-1]+1])
    df1.index.name = 'id_'
    db_6 = db_3.append(df1)
    df2=pd.DataFrame({colmn[0]: entry_1.get()}, index=[int(db_1.index[-1])+1])#cmb_4.get() entry_2.get(), entry_1.get()]},index=[db_1.index[-1]+1])
    df2.index.name = 'id_'
    db_6 = db_6.append(df2)
    db_6=db_6.iloc[:,::-1]
    db_6.to_excel('C:\Work\Data\БД3.xlsx', sheet_name='Лист 1')
    
def clck():
    str_2 = cmb_3.get()
    dict_1 = {'name_': entry.get()}
    df_data=pd.DataFrame(dict_1, index = [db_1.index[-1]+1])
    db = db_1.append(df_data)
    db.index.name = 'id_'
    db.to_excel('C:\Work\Data\БД1.xlsx',sheet_name='Лист 1')    
    i = 0
    while str(list(db_3.columns[1:])[i]).split(' ')[0]!=str_2:
        i+=1
    DATA = []
    for k in list(db_3.columns[1:i+2]):
        DATA.append(str(k).split(' ')[0])
    cmb_4['values'] = DATA
    cmb_4.current(0)

win1 = tki.Tk()
win1.title('Добавление данных о валюте')
win1.geometry('600x400+300+200')
lbl_3 = tki.Label(win1, text = 'Введите название валюты', font = ('Times', 12))
lbl_3.place(x = 10, y = 0)
entry = tki.Entry(win1)
entry.focus()
cmb_4 = ttk.Combobox(win1)
entry.place(x = 300, y = 2)
lbl_4 = tki.Label(win1, text = 'Введите дату начала обмена валют', font = ('Times', 12))
lbl_4.place(x = 10, y = 30)
cmb_3 = ttk.Combobox(win1)
cmb_3['values'] = ['No']+[str(i).split(' ')[0] for i in list(db_3.columns[1:])]
cmb_3.current(0)
cmb_3.focus()
cmb_3.place(x = 300, y = 30)
entry_1 = tki.Entry(win1)
entry_2 = tki.Entry(win1)

cmb_4.focus()
cmb_4.place(x = 60, y = 150)
lbl_5 = tki.Label(win1, text = 'ДАТА', font = ('Times', 12))
lbl_5.place(x = 10, y = 120)

lbl_6 =  tki.Label(win1, text = 'Единиц', font = ('Times', 12))
lbl_6.place(x = 250, y = 120)
lbl_7 =  tki.Label(win1, text = 'Курс', font = ('Times', 12))
lbl_7.place(x = 250, y = 180)

entry_1.focus()
entry_1.place(x = 320, y = 150)

entry_2.focus()
entry_2.place(x = 320, y = 200)
btn_7 = tki.Button(win1, text = "Сохранить!", font = ('Times', 12), command = clck_1)
btn_7.place(x = 340, y = 240, width =200, height = 35)
lbl_5 = tki.Label(win1, text = 'ДАТА', font = ('Times', 12))
btn_6 = tki.Button(win1, text = "Заполнить данные о валюте", font = ('Times', 12), command = clck)
btn_6.place(x = 340, y = 60, width =200, height = 35)
win1.mainloop()

    