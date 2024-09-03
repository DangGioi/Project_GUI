import numpy as np 
import pandas as pd 
import tkinter as tk
import tkinter.ttk as ttk
import ttkbootstrap as ttkb
import tkinter.font as tkfont 
import win32com.client as win32
import xml.etree.ElementTree as ET
import re, os, sys, string, random, shutil, types, webbrowser, datetime 
import xlrd, xlwt, time, inspect, zipfile, warnings, fileinput 
from decimal import * 
from tkinter import * 
from itertools import * 
from pathlib import Path 
from datetime import date 
from ctypes import wintypes 
from xlutils.copy import copy 
from selenium import webdriver 
from PIL import ImageTk, Image 
from ttkbootstrap import Style 
from asammdf import MDF, Signal 
from inspect import currentframe 
from alive_progress import alive_bar
from tkinter import messagebox, filedialog 
from tabulate import tabulate
from openpyxl import load_workbook
import xlsxwriter
import subprocess

# Lib using for Jenkins
from api4jenkins import Jenkins
# Support for cakk macro in VBA
import openpyxl
import xlwings as xw


class Window(tk.Tk):
    
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        dir_path = os. path. dirname(sys.argv[0])

        # Initialize style
        style = Style(theme='flatly')
        # Create style used by default for all Frames
        style.configure( 'BW.TLabel', background='white')
        style.theme_use('flatly')


        # If you want to add Italic for the current font, use: slant='italic'
        self.title_font = tkfont.Font(family='Monoid', size=10, weight='bold')
        self.frame_font = tkfont.Font(size=13 ,weight='bold')
        self.title( 'CSWPR Tools')
        self.geometry ('800x400')
        self.resizable(width=False, height=False)
        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others
        container = ttk.Frame(self)
        container.grid(row=0, column=0, sticky='nsew')
        container.grid_rowconfigure(0, weight=1) 
        container.grid_columnconfigure(0, weight=1)
    
        self.frames = {}
        # for F in (StartPage, PageOne, Page_Unvailable):
        for F in (StartPage, PageOne, Page_Unavailable):
            page_name = F.__name__
            frame = F(parent = container, controller = self)
            self.frames[page_name] = frame

            # put all of the pages in the same location
            # the one on the top of the stacking order
            # will be the one that is visible
            frame.grid(row = 0, column = 0, sticky = 'snew')
        
        # Create menu bar Which belongs to window
        Menu_Bar = Menu(self)
        self.config(menu = Menu_Bar)

        # main tool
        Analysis_Support_Menu = Menu(Menu_Bar, tearoff = 0)
        Menu_Bar.add_cascade(label='Select Tools', menu= Analysis_Support_Menu)
        Analysis_Support_Menu.add_command(label="Compiler Warrning", command=lambda:self.Show_Frame('PageOne'))
       
       # Support tool
        Support_Review_Menu = Menu(Menu_Bar, tearoff = 0)
        Menu_Bar.add_cascade(label='Support Review', menu= Support_Review_Menu)
        Support_Review_Menu.add_command(label="Check task", command=lambda:self.Show_Frame('PageTwo'))

       # Help
        Help_Menu = Menu(Menu_Bar, tearoff = 0)
        Menu_Bar.add_cascade(label='Help', menu= Help_Menu)
        Help_Menu.add_command(label="Link Docupedia", command=lambda:self.Show_Frame('PageThree'))
        Help_Menu.add_command(label="About", command=lambda:self.Show_Frame('PageThree'))
       

        self.Show_Authors_And_Contributes()
        self.Show_Frame('StartPage')
        
        def Close_Window():
            if(messagebox.askokcancel("Quit", "Do you really wish to quit?")):
                self.destroy()

        self.protocol("WM_DELETE_WINDOW", Close_Window)


    def Show_Frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()
        return



    def Show_Authors_And_Contributes(self):
        print("HELLO_WORLD\n")
        return


class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller

        ratio = 1
        dir_path = os.path.dirname(sys.argv[0])
        self.Entertainment_Image = ImageTk.PhotoImage((Image.open(f'{dir_path}\Resources\Picture\Automotive.png')).resize((790, int(390/ratio))))
        self.Entertainment_Image_Label = ttk.Label(self, image = self.Entertainment_Image)
        self.Entertainment_Image_Label.grid(row = 0 , column = 0 , columnspan = 2, padx = 5,pady = 5, sticky = tk.NSEW )

class PageOne(tk.Frame):

    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller
        self.columnconfigure(0, weight = 1)
        self.columnconfigure(1, weight = 1)

        # set default comment for execution Status
        self.Execution_Status_String_1 = StringVar()
        self.Execution_Status_String_1.set(" Have a nice day!")
        self.Execution_Status_String_2 = StringVar()
        self.Execution_Status_String_2.set(" Please enter input into all above blank boxes.")

        # Function name
        self.CSV_Emendation_Label = ttk.Label(self, text= "<< TRIAL TOOLS >>", font = controller.frame_font, anchor= CENTER)

        # Information set up
        self.Input_Information()

        # Execution status set up
        self.Execution_Status()

        # Execution Command set up
        self.Execution_Command()

        # Entertainment set  up
        self.Entertainment()

    def function_main(self):       
        filename = self.CSW_File_Path_Entry.get()
        if (os.path.exists(filename)):
            self.Execution_Status_String_1.set(" Successfully!")
            self.Execution_Status_String_2.set("")
        else:
            self.Execution_Status_String_1.set(" File not found. Please check again!")
            self.Execution_Status_String_2.set("")

        self.Execution_Status_1(self.Execution_Status_String_1)
        self.Execution_Status_2(self.Execution_Status_String_2)

        return
    
    def main(self):
        self.function_main()
        return

    def Input_Information(self):
        # Input information
        self.Information_1_Label_Frame = ttk.LabelFrame(self, text = "Input Information")
        self.Information_1_Label_Frame.grid(row = 1, column = 0, padx=5, pady=5, sticky = tk.EW, columnspan= 6)
        self.Information_1_Label_Frame.columnconfigure(1, weight = 20)
        self.Information_1_Label_Frame.columnconfigure(2, weight = 21)
        self.Information_1_Label_Frame.columnconfigure(3, weight = 13)
        self.Information_1_Label_Frame.columnconfigure(4, weight = 10)
                
        # for input file
        self.CSW_File_Path_Label = ttk.Label(self.Information_1_Label_Frame, text= "        PRT File:    " )
        self.CSW_File_Path_Label.grid(row = 0 , column = 0 , sticky =tk.E)
        self.CSW_File_Path_Entry = ttk.Entry(self.Information_1_Label_Frame, width = 95)
        self.CSW_File_Path_Entry.grid(row = 0 , column = 1 , sticky =tk.EW, pady = 2, ipadx= 10, columnspan=3)
        self.Browse_CSW_Button = ttk.Button(self.Information_1_Label_Frame, text= "Browse", command = self.Browse_CSV_File_Path, width= 12)
        self.Browse_CSW_Button.grid(row = 0 , column = 4 ,padx= 7, pady = 2 ,sticky =tk.EW)
        return
    
    def Execution_Status(self):
        self.Execution_Status_Label_Frame = ttk.LabelFrame(self, text = "Execution Status", width = 390, height = 85)
        self.Execution_Status_Label_Frame.grid(row= 2, column = 0 , padx = 5, pady = 5, sticky = tk.NSEW)
        
        self.Execution_Status_1(self.Execution_Status_String_1)
        self.Execution_Status_2(self.Execution_Status_String_2)
        return
    
    def Execution_Command(self):
        self.Execution_Commands_Label_Frame = ttk.LabelFrame(self,  text = "Execution Commands")
        self.Execution_Commands_Label_Frame.grid(row = 2, column = 1, padx = 5, pady = 5, sticky=tk.NSEW)

        self.Clear_All_Button = ttk.Button(self.Execution_Commands_Label_Frame, text= "Destination", command = self.Destination, width = 10)
        self.Clear_All_Button.grid(row = 5, column = 3, padx = 20, pady = 15, sticky=tk.EW)

        self.Execution_Button = ttk.Button(self.Execution_Commands_Label_Frame, text= "Execute", command = self.Execute, width = 10)
        self.Execution_Button.grid(row = 5, column = 2, padx = 20, pady = 15, sticky=tk.EW)
        
        self.Home_Button = ttk.Button(self.Execution_Commands_Label_Frame, text= "Home", command=lambda: self.controller.Show_Frame("StartPage"), width = 10)
        self.Home_Button.grid(row = 5, column = 4, padx = 20, pady = 15, sticky=tk.EW)
        return
    
    def Show_Frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()
        return

    def Entertainment(self):

        ratio = 1.1
        dir_path = os.path.dirname(sys.argv[0])
        self.Entertainment_Image = ImageTk.PhotoImage((Image.open(f'{dir_path}\Resources\Picture\Semichip.png')).resize((790, int(260/ratio))))
        self.Entertainment_Image_Label = ttk.Label(self, image = self.Entertainment_Image)
        self.Entertainment_Image_Label.grid(row = 3 , column = 0 , columnspan = 2, padx = 5,pady = 5, sticky = tk.NSEW )
        return
    
    def Execute(self):
        self.start_time =  time.time()
        print("\n @ Compiler Warning @")
        print("\n ==> Starting...")
        print("\n     Please kindly wait.")
        self.Execute_1()
        return
    
    def Destination(self):
        global destinationFolder
        global filename
        destinationFolder = os.path.dirname(filename)
        os.startfile(destinationFolder)
        return
    
    def Browse_CSV_File_Path(self):
        global name
        self.CSW_File_Path_Entry.config(state = "normal")
        self.CSW_File_Path_Entry.delete(0, END)
        filename = filedialog.askopenfilename(filetypes = [('*.prt file', '*.prt')])
        filename = self.Modify_path(filename)
        if (str(filename) == "."):
            pass
        else:
            self.CSW_File_Path_Entry.insert(0, str(filename))
        return filename
    
    def Modify_path(self, string):
        string = os.path.normpath(string)
        return string
    
    def Execution_Status_1(self, string):
        if(string.get()  == " Have a nice day!"):
            self.Execution_Status_Label_1 = ttk.Label(self.Execution_Status_Label_Frame, textvariable=string, foreground="blue")
        elif (string.get()  ==  " Successfully!"):
            self.Execution_Status_Label_1 = ttk.Label(self.Execution_Status_Label_Frame, textvariable=string, foreground="green")
        else:
            self.Execution_Status_Label_1 = ttk.Label(self.Execution_Status_Label_Frame, textvariable=string, foreground="#8B0000")
      
        self.Execution_Status_Label_1.place(x = 10 , y = 20, anchor = "w")
        return
    
    def Execution_Status_2(self, string):
        if(string.get()  == " Please enter input into all above blank boxes."):
            self.Execution_Status_Label_2 = ttk.Label(self.Execution_Status_Label_Frame, textvariable=string, foreground="blue")
        # elif (string.get()  ==  " Successfully!"):
        #     self.Execution_Status_Label_2 = ttk.Label(self.Execution_Status_Label_Frame, textvariable=string, foreground="green")
        else:
            self.Execution_Status_Label_2 = ttk.Label(self.Execution_Status_Label_Frame, textvariable=string, foreground="#8B0000")
      
        self.Execution_Status_Label_2.place(x = 10 , y = 45, anchor = "w")
        return
    
    def Execute_1(self):
        
        if(self.CSW_File_Path_Entry.get() and (self.CSW_File_Path_Entry.get() != ".")):
            self.main()
        else:
            print("Missing CSV file path")
            self.Execution_Status_String_1.set(" Missing *.prt file Path. Check again")
            self.Execution_Status_String_2.set("")

        self.Execution_Status_1(self.Execution_Status_String_1)
        self.Execution_Status_1(self.Execution_Status_String_2)

        print("\n    Time         : % secs" % round((time.time() -self.start_time), 1 ))
        print("\n--- DONE ---")
        return
    


class Page_Unavailable(tk.Frame):

    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent)
        self.controller = controller
        label = ttk.Label(self, text = "This page is still unavailable", font= controller.title_font, anchor = CENTER)
        label.pack(side= "top", fill = "x", pady = 10)

if __name__ == "__main__":

    app = Window()
    app.mainloop()

