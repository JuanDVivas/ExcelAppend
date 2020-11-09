import numpy as np
import pandas as pd
import os
import glob
from tkinter import *
from tkinter import messagebox
try:
    import Tkinter as tkinter
    import ttk
except ImportError:
    import tkinter
    from tkinter import ttk

data = pd.DataFrame()

def data_pr():
    path = path_info.get() + '/'

    filelist = []

    os.chdir(path)
    for files in glob.glob('*.xlsx'):
        filelist.append(files)
        
    progbar = ttk.Progressbar(formulario, variable=count, maximum=len(filelist))
    progbar.place(x=245 , y=280)
    

    def bring_data():
        for counter, fileitem in enumerate(filelist):
            global data
            temp = pd.read_excel(path + fileitem)
            temp.loc[:,'Origen Data'] = fileitem
            data = data.append(temp)
            count.set(counter+1)
            formulario.update_idletasks()
            refresh()
            
        data.to_csv(path + 'Archivo Full.csv', index=False)

    bring_data()

    messagebox.showinfo('Union de Datos', 'Union de Datos Completada'+ '\n' + 'Archivos Procesados:'+ '\n' + '\n'.join(''.join(file) for file in filelist))
    

formulario = Tk()
formulario.geometry('600x300')
formulario.title('Union de Datos')
heading=Label(text = 'App para Union de Datos con Formato Excel', bg='navy', fg='white', width='500', height= '5', font=('Helvetica 18 bold',12))
heading.pack()

path_name=Label(text = 'Path:', bg='navy', fg='white')

path_name.place(x=280 , y=140)

path_info=StringVar()

path_entry= Entry(textvariable = path_info, width='57')

path_entry.place(x=120 , y=170)

script= Button(formulario, text='Ejecutar Script', width= '30', height='2' , command= data_pr, fg='white', bg='navy')

script.place(x=185 , y=230)

count = DoubleVar()

def refresh():
    formulario.update()
    formulario.after(1000,refresh)

formulario.mainloop()
