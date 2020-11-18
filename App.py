import numpy as np
import pandas as pd
import os
import glob
from tkinter import *
from tkinter import messagebox
from xlsx2csv import Xlsx2csv
try:
    import Tkinter as tkinter
    import ttk
except ImportError:
    import tkinter
    from tkinter import ttk


def data_pr():
    path = path_info.get() + '/'

    filelist = []  

    os.chdir(path)
        
    def ex_csv():
        
        for files in glob.glob('*.xlsx'):
            filelist.append(files)
            
        progbar = ttk.Progressbar(formulario, variable=count, maximum=len(filelist))
        progbar.place(x=130 , y=280)
        
        lista = pd.DataFrame(filelist)
        lista.to_csv(path + 'lista archivos.txt', header=False)
    
        for counter, fileitem in enumerate(filelist):
            csv=fileitem.replace('.xlsx','.csv')
            
            try:    
                Xlsx2csv(path + fileitem, outputencoding='utf-8').convert(path + csv)
            except:
                pass
            
            count.set(counter+1)
            formulario.update_idletasks()
            refresh()    
     
    ex_csv()

    messagebox.showinfo('Conversión de Datos', 'Conversión de Datos Completada'+ '\n' + 'Archivos Procesados:'+ '\n' + '\n'.join(''.join(file) for file in filelist))    

def data_ps():
    
    path = path_info.get() + '/'
    
    filelist2 = []
    
    os.chdir(path)
        
    def bring_data():
        
        for files in glob.glob('*.csv'):
            filelist2.append(files)
        
        try:
            filelist2.remove('Archivo Full.csv')
            filelist2.remove('lista archivos.txt')
        except:
            pass
        
        progbar = ttk.Progressbar(formulario, variable=count2, maximum=len(filelist2))
        progbar.place(x=370 , y=280)
        
        for counter, fileitem in enumerate(filelist2):
            temp = pd.read_csv(path + fileitem)
            temp.loc[:,'Origen Data'] = fileitem
                
            if counter==0:
                temp.to_csv(path + 'Archivo Full.csv', index=False)
            else:
                temp.to_csv(path + 'Archivo Full.csv', index=False, mode='a', header=False)
                
            count2.set(counter+1)
            formulario.update_idletasks()
            refresh()
            
    bring_data()

    messagebox.showinfo('Union de Datos', 'Union de Datos Completada'+ '\n' + 'Archivos Procesados:'+ '\n' + '\n'.join(''.join(file) for file in filelist2))
    

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

script= Button(formulario, text='Xlsx to Csv', width= '30', height='2' , command= data_pr, fg='white', bg='navy')
script2= Button(formulario, text='Csv Append', width= '30', height='2' , command= data_ps, fg='white', bg='navy')

script.place(x=70 , y=230)
script2.place(x=310 , y=230)

count = DoubleVar()
count2 = DoubleVar()

def refresh():
    formulario.update()
    formulario.after(1000,refresh)

formulario.mainloop()
