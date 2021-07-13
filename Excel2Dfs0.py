"""
Excel to Dfs0 Tool:

Created on Wed June 1, 2021

@author: Shubhneet Singh 
ssin@dhigroup.com
DHI,US

"""
# Dependencies
import os
import clr
import sys
import time
import openpyxl
import numpy as np #
import pandas as pd #
import datetime
import math

from winreg import ConnectRegistry, OpenKey, HKEY_LOCAL_MACHINE, QueryValueEx

def get_mike_bin_directory_from_registry():
    x86 = False
    dhiRegistry = "SOFTWARE\Wow6432Node\DHI\\"
    aReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
    try:
        _ = OpenKey(aReg, dhiRegistry)
    except FileNotFoundError:
        x86 = True
        dhiRegistry = "SOFTWARE\Wow6432Node\DHI\\"
        aReg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
        try:
            _ = OpenKey(aReg, dhiRegistry)
        except FileNotFoundError:
            raise FileNotFoundError
    year = 2030
    while year > 2010:
        try:
            mikeHomeDirKey = OpenKey(aReg, dhiRegistry + str(year))
        except FileNotFoundError:
            year -= 1
            continue
        if year > 2020:
            mikeHomeDirKey = OpenKey(aReg, dhiRegistry + "MIKE Zero\\" + str(year))

        mikeBin = QueryValueEx(mikeHomeDirKey, "HomeDir")[0]
        mikeBin += "bin\\"

        if not x86:
            mikeBin += "x64\\"

        if not os.path.exists(mikeBin):
            print(f"Cannot find MIKE ZERO in {mikeBin}")
            raise NotADirectoryError
        return mikeBin

    print("Cannot find MIKE ZERO")
    return ""

sys.path.append(get_mike_bin_directory_from_registry())
clr.AddReference("DHI.Generic.MikeZero.DFS")
clr.AddReference("DHI.Generic.MikeZero.EUM")
clr.AddReference("DHI.Projections")

# import xlrd
from  mikeio import *
from mikeio.eum import ItemInfo, EUMType, EUMUnit
from DHI.Generic.MikeZero.DFS import DataValueType

from tkinter import Frame, Label, Button, Entry, Tk, W, END
from tkinter import messagebox as tkMessageBox
# from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

 
#------------------------------------------------------------------------------
# UI for this tool:
     
class interface(Frame):
    def __init__(self, master = None):
        """ Initialize Frame. """
        Frame.__init__(self,master)
        self.grid()
        self.createWidgets()
            
    def message(self):
        tkMessageBox.showinfo("Task Complete", "Dfs0 Created!")
    
    def run(self):
        
        # input1 - Data in excel:
        filename1 = self.file_name1.get()
        # filename1 = r"C:\Users\ssin\OneDrive - DHI\Documents\Shuby's Toolbox\2 Excel2Dfs0(Large) Tool\Excel2Dfs0_Example.xlsx"
        # Output:
        outputFile = self.file_name5.get()
        # outputFile = r"C:\Users\ssin\OneDrive - DHI\Documents\Shuby's Toolbox\2 Excel2Dfs0(Large) Tool\Test1.dfs0"
        # Tool
        begin_time = datetime.datetime.now()
        
        if str(os.path.splitext(filename1)[1])== '.csv':
            df = pd.read_csv(filename1, index_col=0)
        elif str(os.path.splitext(filename1)[1])== '.xlsx':
            df = pd.read_excel(filename1, index_col=0,engine='openpyxl')
                  
        print('Dataframe created. Time taken > ' + str(datetime.datetime.now() - begin_time))
        
        # df.index = [df.index[i].round('1s') for i in range(len(df.index))]
        item_names = df.columns
        items_type = df.iloc[0]        
        item_units =  df.iloc[1]
        item_datatype = df.iloc[2]
        
        
        if type(item_names[0]) != str:
            item_names = [str(item_names[i]) for i in range(len(item_names))]
        df=df[3:]
        
        data=[df.iloc[:,i].values for i in range(len(df.columns))]
        time=df.index
        
        
        if pd.isnull(item_units.iloc[0]) != True:
            items=[ItemInfo(item_names[i], EUMType[items_type.iloc[i]], EUMUnit[item_units.iloc[i]]) for i in range(len(item_names))]
        else: 
            items=[ItemInfo(item_names[i], EUMType[items_type.iloc[i]]) for i in range(len(item_names))]

        data_type = [item_datatype.iloc[i] for i in range(len(item_datatype))]
        dtypes = ['Instantaneous', 'Accumulated', 'StepAccumulated' , 'MeanStepBackward', 'MeanStepForward']
        
                
        
        data_value_type=[None]*len(data_type)
        
        for i in range(len(data_type)):
            for a in range(len(dtypes)):
                if data_type[i]== dtypes[a]:
                    data_value_type[i] = a

        ds = Dataset(data, time,items)
        
        qq = exec('DataValueType' + "." + item_datatype.iloc[0])
        
        outputFile_dropExt = str(os.path.splitext(outputFile)[0])
        dfs_title = outputFile_dropExt.rsplit("/", 1)[1]
        print('Dfs0 writing commences > ' + str(datetime.datetime.now() - begin_time))
        
        dfs = Dfs0()

        dfs.write(filename= outputFile, data=ds, title= dfs_title, data_value_type=data_value_type)
        
        print('Congrats, dfs0 created! Time taken > ' + str(datetime.datetime.now() - begin_time))

        self.message()
        

    def createWidgets(self):
        
        # set all labels of inputs:

        Label(self, text = "Data (*.csv or *.xlsx) :")\
            .grid(row=0, column=0, sticky=W)
            
        Label(self, text = "Output File (*.dfs0) :")\
            .grid(row=1, column=0, sticky=W)
            
        # set buttons
        Button(self, text = "Browse", command=self.load_file1, width=10)\
            .grid(row=0, column=6, sticky=W)
        
        Button(self, text = "Save As", command=self.load_file5, width=10)\
            .grid(row=1, column=6, sticky=W)            
        Button(self, text = "Run", command=self.run, width=20)\
            .grid(row=4, column=3, sticky=W)
       
        # set entry field
        self.file_name1 = Entry(self, width=65)
        self.file_name1.grid(row=0, column=1, columnspan=4, sticky=W)

        self.file_name5 = Entry(self, width=65)
        self.file_name5.grid(row=1, column=1, columnspan=4, sticky=W)

    def load_file1(self):
        self.filename = askopenfilename(initialdir=os.path.curdir, defaultextension=".xlsx", filetypes=(("csv", "*.csv"),("xlsx File", "*.xlsx"),("All Files", "*.*") ))
        if self.filename: 
            try: 
                #self.settings.set(self.filename)
                self.file_name1.delete(0, END)
                self.file_name1.insert(0, self.filename)
                self.file_name1.xview_moveto(1.0)
            except IOError:
                tkMessageBox.showerror("Error","Failed to read file \n'%s'"%self.filename) 
   
     
    def load_file5(self):
        self.filename = asksaveasfilename(initialdir=os.path.curdir,defaultextension=".dfs0", filetypes=(("Dfs0 File", "*.dfs0"),("All Files", "*.*") ))
        if self.filename: 
            try: 
                #self.settings.set(self.filename)
                self.file_name5.delete(0, END)
                self.file_name5.insert(0, self.filename)
                self.file_name5.xview_moveto(1.0)
            except IOError:
                tkMessageBox.showerror("Error","Failed to read file \n'%s'"%self.filename) 
                
##### main program


root = Tk()
UI = interface(master=root)
UI.master.title("Excel to Dfs0 Tool")
UI.master.geometry('625x120')
for child in UI.winfo_children():
    child.grid_configure(padx=4, pady =6)
    
UI.mainloop()
