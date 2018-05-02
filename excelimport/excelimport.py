from openpyxl import *
from tkinter import *
from tkinter.filedialog import askopenfilename

# =======================================
#           File Import Pop Up
# =======================================

# This creates the whole open/select a file thing and asks what sheet the data is on

window = Tk()

print("Starting Tkinter Open Window")

filetypes = [("Spreadsheet", "*.xlsx")]
title = "Import Spreadsheet"
initialdir = "C:\\"

window.fileName = askopenfilename(filetypes=filetypes, initialdir=initialdir, title=title)
filename = window.fileName
# Above is case sensitive

sheetname = ".dog"
sheetnameentryvar = StringVar()

def sheetnamebutton():
    sheetname = sheetnameentryvar.get()
    window.destroy()

Label(window, text="What sheet is the data on?", bg="#1e1e1e", fg="#f9f9f9").grid(row=0, column=0, sticky=W)
sheetnameentry = Entry(window, width=30, bg="#f9f9f9", fg="#1e1e1e", textvariable=sheetnameentryvar)
sheetnameentry.grid(row=1, column=0, sticky=W)
Button(window, text="Enter", width=20, command=sheetnamebutton).grid(row=2, column=0, sticky=W)

window.mainloop()


class FileImport:
    __init
