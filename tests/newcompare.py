from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
import tkinter.filedialog
import tkinter.messagebox
import sys

sys.path.append("C:\Program Files (x86)\PTI\PSSE33\PSSBIN")

import psse35
import psspy

import win32com.client
import time
import string
import os
import winreg
import pywintypes
import re

#create a new window
#_____________________________________________________________________________________________________
root = Tk()
root.title('Compare .sav Case Tool')
root.geometry('900x300+367+276')
root.resizable(width = False, height = False)
#_____________________________________________________________________________________________________
global SourceDir
global FirstFile
global SecondFile
global DestinationDir

SourceDir = StringVar()
FirstFile = StringVar()
SecondFile = StringVar()
DestinationRootDir = StringVar()
#_____________________________________________________________________________________________________
def first_file_select():
    FirstFile = tkinter.filedialog.askopenfilename(title = "Select file", filetypes = [("Save Case File","*.sav")])
    FirstFileEntry.insert(0, FirstFile)
    FirstFileEntry.xview(END)

    if len(FirstFileEntry.get()) != 0:
        FirstFileEntry.delete(0, END)
        FirstFileEntry.insert(0, FirstFile)
        FirstFileEntry.xview(END)

def second_file_select():
    SecondFile = tkinter.filedialog.askopenfilename(title = "Select file", filetypes = [("Save Case File","*.sav")])
    SecondFileEntry.insert(0, SecondFile)
    SecondFileEntry.xview(END)

    if len(SecondFileEntry.get()) != 0:
        SecondFileEntry.delete(0, END)
        SecondFileEntry.insert(0, SecondFile)
        SecondFileEntry.xview(END)

def destination_dir_select():
    DestinationDir = tkinter.filedialog.askdirectory(title = "Select Folder")
    DestinationRootDirEntry.insert(0, DestinationDir)
    DestinationRootDirEntry.xview(END)

    if len(DestinationRootDirEntry.get()) != 0:
        DestinationRootDirEntry.delete(0, END)
        DestinationRootDirEntry.insert(0, DestinationDir)
        DestinationRootDirEntry.xview(END)

#_____________________________________________________________________________________________________

def createExcelSpreadsheet():
    # Create the workbook.
    xlApp = win32com.client.Dispatch("Excel.Application")
    xlApp.DisplayAlerts = False

    # Set below to False to hide the spreadsheet while working, or True to show it.
    xlApp.Visible = False
    xlApp.Workbooks.Add()
    xlBook = xlApp.ActiveWorkbook
    return xlApp, xlBook

def createTitleBlock(xlBook):
    xlSheet = xlBook.Worksheets.Add()    # Add a sheet for the title.
    xlSheet.Name = 'Title'
    xlSheet.Range("A:Z").Font.Name = "Courier New"

    print("Creating title block.")
    title1, title2 = psspy.titldt()

    data = [[title1],[title2]]
    data.append([''])    #creates a blank row

    #data.append(["REGION CHECKED : %s" % region[0]])
    data.append(["AREAS CHECKED : SCFG"])

    data.append([''])    #creates a blank row

    data.append([time.strftime("PROCESSED: %a, %d %b %Y %I:%M:%S %p",time.localtime()).upper()])

    data.append([''])    #creates a blank row

    data.append(['WORKING CASE = FIRST CASE'])
    data.append(['IN = SECOND CASE'])

    xlSheet.Cells(8,1).Font.Bold = True
    xlSheet.Cells(9,1).Font.Bold = True

def BranchRating(xlBook):
    xlSheet = xlBook.Worksheets.Add()
    xlSheet.Name = 'BranchRating'
    xlSheet.Range("A:Z").Font.Name = "Courier New"
    xlSheet.Tab.ColorIndex = 6    # Yellow tab

    data = []
    directory = os.getcwd()
    filename = os.path.join(directory, "BranchRating_temp.txt")

    psspy.bsys(0,0,[0.0,999.],1,[343],0,[],0,[],0,[])
    #prepare case for compare script. It is set up to compare based on bus numbers

    psspy.report_output(2, filename, [0,0])
    #Redirect progress device to a file.

    #psspy.diff(0,0,1,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    psspy.diff(0,1,2,[0,20,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    #compare BRANCHES WITH DIFFERENT LINE RATINGS

    #psspy.diff(0,1,3,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    #psspy.close_report()

    print("Reading file BranchRating_temp.txt")

    datafile = open(filename, 'r')
    for line in datafile:
        data.append([line.strip()])

    datafile.close()

    sendToExcel(xlSheet, data, startrow = 2)

    xlSheet.Cells(1,1).Value = "CASE COMPARE"
    xlSheet.Cells(1,1).HorizontalAlignment = -4108
    xlSheet.Cells(1,1).Font.Bold = True

    xlSheet.rows(2).EntireRow.Delete()
    xlSheet.rows(2).EntireRow.Delete()

    xlSheet.Columns("A:Z").AutoFit()

    #os.remove(filename)

    print("Done with Branch Ratings")

def XFMRs(xlBook):
    xlSheet = xlBook.Worksheets.Add()
    xlSheet.Name = 'XFMRs'
    xlSheet.Range("A:Z").Font.Name = "Courier New"
    xlSheet.Tab.ColorIndex = 6    # Yellow tab

    data = []
    directory = os.getcwd()
    filename = os.path.join(directory, "XFMRs_temp.txt")

    psspy.bsys(0,0,[0.0,999.],1,[343],0,[],0,[],0,[])
    #prepare case for compare script. It is set up to compare based on bus numbers

    psspy.report_output(2, filename, [0,0])

    #psspy.diff(0,0,1,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    psspy.diff(0,1,2,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())

    #compare TRANSFORMERS WITH DIFFERENT CONFIGURATION OR RATIO
    #DIFFERING BY MORE THAN 0.0000 PU

    #psspy.diff(0,1,3,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    #psspy.close_report()

    print("Reading file XFMRs_temp.txt")

    datafile = open(filename, 'r')
    for line in datafile:
        data.append([line.strip()])

    datafile.close()

    sendToExcel(xlSheet, data, startrow = 2)

    xlSheet.Cells(1,1).Value = "CASE COMPARE"
    xlSheet.Cells(1,1).HorizontalAlignment = -4108
    xlSheet.Cells(1,1).Font.Bold = True

    xlSheet.rows(2).EntireRow.Delete()
    xlSheet.rows(2).EntireRow.Delete()

    xlSheet.Columns("A:Z").AutoFit()

    #os.remove(filename)

    print("Done with XFMRs")

def Loads(xlBook):
    xlSheet = xlBook.Worksheets.Add()
    xlSheet.Name = 'Loads'
    xlSheet.Range("A:Z").Font.Name = "Courier New"
    xlSheet.Tab.ColorIndex = 6    # Yellow tab

    data = []
    directory = os.getcwd()
    filename = os.path.join(directory, "Loads_temp.txt")

    psspy.bsys(0,0,[0.0,999.],1,[343],0,[],0,[],0,[])

    psspy.report_output(2, filename, [0,0])

    #psspy.diff(0,0,1,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    psspy.diff(0,1,1,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    #compare BUSES WITH DIFFERENT LOADS OR LOAD STATUS

    psspy.diff(0,1,2,[0,32,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())

    #psspy.diff(0,1,3,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    #psspy.close_report()

    print("Reading file Loads_temp.txt")

    datafile = open(filename, 'r')
    for line in datafile:
        data.append([line.strip()])

    datafile.close()

    sendToExcel(xlSheet, data, startrow = 2)

    xlSheet.Cells(1,1).Value = "CASE COMPARE"
    xlSheet.Cells(1,1).HorizontalAlignment = -4108
    xlSheet.Cells(1,1).Font.Bold = True

    xlSheet.rows(2).EntireRow.Delete()
    xlSheet.rows(2).EntireRow.Delete()

    xlSheet.Columns("A:Z").AutoFit()

    #os.remove(filename)

    print("Done with Loads")

def LineLength(xlBook):
    xlSheet = xlBook.Worksheets.Add()
    xlSheet.Name = 'LineLength'
    xlSheet.Range("A:Z").Font.Name = "Courier New"
    xlSheet.Tab.ColorIndex = 6    # Yellow tab

    data = []
    directory = os.getcwd()
    filename = os.path.join(directory, "LineLength_temp.txt")

    psspy.bsys(0,0,[0.0,999.],1,[343],0,[],0,[],0,[])
    #prepare case for compare script. It is set up to compare based on bus numbers
    psspy.report_output(2, filename, [0,0])
    #psspy.diff(0,0,1,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    psspy.diff(0,1,1,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    #compare BRANCHES WITH DIFFERENT LINE LENGTHS
    psspy.diff(0,1,2,[0,33,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    #psspy.diff(0,1,3,[0,0,0,0],[0.0,0.0,0.0,0.0],SecondFileEntry.get())
    #psspy.close_report()

    print("Reading file LineLength_temp.txt")
    datafile = open(filename, 'r')
    for line in datafile:
        data.append([line.strip()])
    datafile.close()
    #print("Write results.")

    sendToExcel(xlSheet, data, startrow = 2)

    xlSheet.Cells(1,1).Value = "CASE COMPARE"
    xlSheet.Cells(1,1).HorizontalAlignment = -4108
    xlSheet.Cells(1,1).Font.Bold = True

    xlSheet.rows(2).EntireRow.Delete()
    xlSheet.rows(2).EntireRow.Delete()

    xlSheet.Columns("A:Z").AutoFit()

    #os.remove(filename)

    print("Done with Line Length")

def Compare_Cases_Excel():

    psspy.psseinit(150000)
    psspy.case(FirstFileEntry.get())

    print("**Working**")

    xlApp, xlBook = createExcelSpreadsheet()

    Buses(xlBook)          #NOTE: You must initialize the Diff API on first call
    LineLength(xlBook)
    Loads(xlBook)
    XFMRs(xlBook)
    BranchRating(xlBook)
    BranchImp(xlBook)
    BranchStatus(xlBook)
    Machines(xlBook) #NOTE: You must close the Diff API and close the report before you can delete the temporary file.

    createTitleBlock(xlBook)

    psspy.close_powerflow()
    #Dummy()

    directory = os.getcwd()

    xlBook.SaveAs(os.path.join(directory, "CompareCase.xlsx"))
    xlBook.Close()

    os.remove("LineLength_temp.txt")
    os.remove("Loads_temp.txt")
    os.remove("XFMRs_temp.txt")
    os.remove("BranchRating_temp.txt")
    os.remove("BranchImp_temp.txt")
    os.remove("BranchStatus_temp.txt")
    #os.remove("Machines_temp.txt")
    os.remove("Buses_temp.txt")

    print("**Finished!**")
    tkinter.messagebox.showinfo('Info','Finished! The file "CompareCase.xlsx" has been created!' )

    sys.exit(0)
#_____________________________________________________________________________________________________

DirectorySetupData = Frame(root, width = 900, height = 350)
DirectorySetupData.pack()

StaticLabel1 = Label(DirectorySetupData, justify = LEFT ,text = 'Select Cases Below You Want To Compare. \nIf Destination Folder is left out, the result will be saved in Source Folder.' )
StaticLabel1.place(x = 48, y = 30)

#First .save case info
StaticLabel2 = Label(DirectorySetupData, justify = LEFT ,text = 'First .sav Case:' )
StaticLabel2.place(x = 48, y = 70)

FirstFileEntry = Entry(DirectorySetupData, width = 90)
FirstFileEntry.place(x = 50, y = 92)

SourceDirButton = Button(DirectorySetupData, text = 'Select', width = 10, height = 1, command = first_file_select)
SourceDirButton.place(x = 650, y = 90)

#second .save case info
StaticLabel2 = Label(DirectorySetupData, justify = LEFT ,text = 'Second .sav Case:' )
StaticLabel2.place(x = 48, y = 115)

SecondFileEntry = Entry(DirectorySetupData, width = 90)
SecondFileEntry.place(x = 50, y = 137)

SourceDirButton = Button(DirectorySetupData, text = 'Select', width = 10, height = 1, command = second_file_select)
SourceDirButton.place(x = 650, y = 135)

#Destination Folder
StaticLabel3 = Label(DirectorySetupData, justify = LEFT ,text = 'Results Destination Folder:' )
StaticLabel3.place(x = 48, y = 160)

DestinationRootDirEntry = Entry(DirectorySetupData, width = 90)
DestinationRootDirEntry.place(x = 50, y = 182)

DestinationRootDirButton = Button(DirectorySetupData, text = 'Select', width = 10, height = 1, command = destination_dir_select)
DestinationRootDirButton.place(x = 650, y = 182)

GenerateCompare = Button(DirectorySetupData, text = 'Compare Cases', width = 20, height = 1, command = Compare_Cases_Excel)
GenerateCompare.place(x = 370, y = 235)

root.mainloop()
