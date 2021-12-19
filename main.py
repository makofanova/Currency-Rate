import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.gridspec as gridspec
import os


db_1 = pd.read_excel('C:\Work\Data\БД1.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_2 = pd.read_excel('C:\Work\Data\БД2.xlsx',sheet_name='Лист 1', index_col=('id_country'))
db_3 = pd.read_excel('C:\Work\Data\БД3.xlsx',sheet_name='Лист 1', index_col=('id_'))
db_4 = pd.read_excel('C:\Work\Data\БД4.xlsx',sheet_name='Лист 1', index_col=('id_country'))
    
def clck_1(): # удаление строки с данными
    os.system('C:/Work/Library/win_del.py')
    
import tkinter as tki
import tkinter.ttk as ttk
        
def make_combobox_1():
    cmb_1['values'] = list_
    cmb_1.current(0)
    cmb_1.focus()
    cmb_1.place(x = 375, y = 40)
    if not var_1.get():
        cmb_1.place_forget()



def Kurs_priv(Kurs, Unit):
    sum = 0
    for i in Unit:
        if i=='nan':
            sum+=0
        else:
            sum+=i
        
    sr = sum/len(Unit)
    Flag = True
    if sr>1 and sr<10:
        k = 1
    elif sr>10 and sr<100:
        k = 10
    elif sr>100 and sr<1000:
        k = 100
    elif sr>1000:
        k = 1000
    elif sr==1 or sr==10 or sr==100 or sr==1000:
        k = sr
        Flag = False
    if Flag:
        print(len(Kurs))
        print(len(Unit))
        for i in range(len(Kurs)):
            if Unit[i]!=k:
                Kurs[i]=Kurs[i]*k/Unit[i]

def graph_1(j, str_val):
    print(j)
    Data=[]
    for i in range(len(db_3.columns[1:])+1):
        Data.append(str(db_3.columns[i]).split(' ')[0])
    Data = Data[1:]
    Kurs=list(db_3.values)[j]
    Kurs = Kurs[1:]
    Kurs_priv(Kurs, list(db_3.values)[j-1][1:])
    gs = gridspec.GridSpec(10, 40)
    fig = plt.figure(figsize=(340,170))
    ax = fig.add_subplot(gs[:7, :])
    ax.grid(False)
    plt.plot(Data,Kurs)
    plt.xlabel('Дата')
    plt.ylabel('Курс валюты')
    plt.title(str_val)
    plt.xticks([Data[i] for i in range(len(Data)) if i % 18 == 0], rotation='vertical')
    plt.rc('xtick', labelsize=7)
    canvas = FigureCanvasTkAgg(fig, win0)
    canvas.get_tk_widget().place(x = 10, y = 10,  width = 340, height = 180)
    path = "C:/Work/Graphics/"+'Курс '+ str_val
    plt.savefig(path)
    
def graph_2():
    count = []
    currency = []
    for i in range(1,len(db_1['name_'])+1):
        currency.append(list(db_1['name_'])[i-1])
    
    for i in range(1,len(db_1['name_'])+1):
        count.append(list(db_2['id_']).count(i))
    
    #plt.bar(currency, count)
    val = [int(100*i/len(list(db_2['id_']))) for i in count]
    labels =['Другое']
    vals =[0]
    
    for i in range(len(val)):
        if val[i]>1:
            vals.append(val[i])
            labels.append(currency[i])
        else:
            vals[0]+=1
            
    fig, ax = plt.subplots(figsize=(340,170))
    ax.set_title('Круговая диаграмма распространенности валют')
    path = "C:/Work/Graphics/"+'Круговая диаграмма распространенности валют'
    ax.pie(vals, labels = labels)
    ax.axis("equal")
    canvas = FigureCanvasTkAgg(fig, win0)
    canvas.get_tk_widget().place(x = 10, y = 10,  width = 340, height = 180)    
    plt.savefig(path)
    
def clck():
    print(cmb_1.get(), cmb_2.get())
    if cmb_1.get()!= 'No' and cmb_1.get()!='Распространенность валют в мире':
        graph_1(int(db_1[(db_1.values== cmb_1.get())].index[0])*2+1, cmb_1.get())
    elif cmb_1.get()=='Распространенность валют в мире':
        graph_2()
    else:
        lbl = tki.Label(f_top, text = 'Упс... Вы ничего не выбрали!', font = ('Times', 12))
        lbl.pack(side = 'top')
        
def clck_2():
    os.system('C:/Work/Library/add_win.py')

def clck_4():
    db_1 = pd.read_excel('C:\Work\Data\БД1.xlsx',sheet_name='Лист 1', index_col=('id_'))
    db_2 = pd.read_excel('C:\Work\Data\БД2.xlsx',sheet_name='Лист 1', index_col=('id_'))
    db_3 = pd.read_excel('C:\Work\Data\БД3.xlsx',sheet_name='Лист 1', index_col=('id_'))
    db_4 = pd.read_excel('C:\Work\Data\БД4.xlsx',sheet_name='Лист 1', index_col=('id_country'))
    tree = ttk.Treeview(win0)
    tree['show'] = 'headings'
    col = ('Название валюты', 'Данные')
    for i in db_3.columns[1:]:
        col = col + ((str(i).split(' ')[0]),)
    tree['columns']=col
    tree.column('Название валюты', width= 150)
    tree.heading('Название валюты', text = 'Название валюты')
    tree.column('Данные', width= 100)
    tree.heading('Данные', text = 'Данные')
    for i in col[2:]:
        tree.column(i, width= 70)
        tree.heading(i, text = i)
    
    list_=['No', 'Распространенность валют в мире']
    date_list = [list(k) for k in list(db_3.values)]
    for i in range(len(list(db_1.values))):
        date_list[2*i].insert(0,db_1.values[i][0])
        date_list[2*i+1].insert(0,' ')
        list_.append(db_1.values[i][0])
        
    date_list.reverse()
    tree.place(x = 15 , y = 225, width=750, height=240)
    for i in date_list:
        tree.insert('', 0, values = tuple(i))
        
def clck_3():
    os.system('C:\Work\Library\отбор.py')

#создание окна
win0= tki.Tk()
win0.title('Изменение курса валют в период Пандемии')
win0.geometry('800x500+300+200')

f_top = tki.LabelFrame(win0, text = 'визуализация')
f_top.place(x = 10, y = 10,  width = 340, height = 180)

# создание лейблов


lbl_2 = tki.Label(win0, text = 'База данных', font = ('Times', 12))
lbl_2.place(x = 15, y = 200 )
i =0
var_3 =[]
# создание кнопки
btn = tki.Button(win0, text = 'Выполнить', font = ('Times', 12), command = clck)
btn.place(x = 650, y = 120 )
# создание таблицы
tree = ttk.Treeview(win0)
tree['show'] = 'headings'
col = ('Название валюты', 'Данные')
for i in db_3.columns[1:]:
    col = col + ((str(i).split(' ')[0]),)
tree['columns']=col
tree.column('Название валюты', width= 150)
tree.heading('Название валюты', text = 'Название валюты')
tree.column('Данные', width= 100)
tree.heading('Данные', text = 'Данные')
for i in col[2:]:
    tree.column(i, width= 70)
    tree.heading(i, text = i)

list_=['No', 'Распространенность валют в мире']
date_list = [list(k) for k in list(db_3.values)]
print(len(date_list))
print(len(list(db_1.values)))
for i in range(len(list(db_1.values))):
    date_list[2*i].insert(0,db_1.values[i][0])
    date_list[2*i+1].insert(0,' ')
    list_.append(db_1.values[i][0])
    
date_list.reverse()
tree.place(x = 15 , y = 225, width=750, height=240)

for i in date_list:
    tree.insert('', 0, values = tuple(i))
    
cmb_1 = ttk.Combobox(win0)
cmb_2 = ttk.Combobox(win0)
var_1 = tki.BooleanVar()
var_1.set(0)
c_1 = tki.Checkbutton(win0, text= 'Провести общий анализ данных',  font = ('Times', 12), variable=var_1,
                 onvalue=1, offvalue=0, command = make_combobox_1)
c_1.place(x =350, y = 10)

# создание двух скроллбаров    
scrl_x = ttk.Scrollbar(win0, command = tree.xview, orient = tki.HORIZONTAL)
scrl_x.place(x = 15, y = 465, width = 750, height = 20)
tree.configure(xscrollcommand = scrl_x.set)

scrl_y = ttk.Scrollbar(win0, command = tree.yview, orient = tki.VERTICAL)
scrl_y.place(x = 765, y = 225, width = 20, height = 240)
tree.configure(yscrollcommand = scrl_y.set)

btn_1 = tki.Button(win0, text = 'Удалить валюту', font = ('Times', 12), command = clck_1)
btn_1.place(x = 125, y = 190, width = 130, height = 35)

btn_2 = tki.Button(win0, text = 'Добавить валюту', font = ('Times', 12), command = clck_2)
btn_2.place(x = 250, y = 190, width = 130, height = 35 )

btn_5 = tki.Button(win0, text = "Отобрать", font = ('Times', 12), command = clck_3)
btn_5.place(x = 640, y = 190, width = 140, height = 35 )

win0.mainloop()




























