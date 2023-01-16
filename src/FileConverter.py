from tkinter import messagebox, ttk
from win32com import client
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from PIL import ImageTk
from docx2pdf import convert



#CONTROL_FORMAT
def control_format():
    if format_combobox.get() == 'excel':
        excel()
    elif format_combobox.get() == 'word':
        word()
    else:
        messagebox.showinfo("Info", "Please select file type!")

#EXCEL
def excel():
    file = filedialog.askopenfilename(initialdir="/",title="Choose a file")
    
    try:
        
        app= client.DispatchEx("Excel.Application")
        app.interactive = False
        app.Visible = False

        workbook = app.Workbooks.open(file)
        workbook.ActiveSheet.ExportAsFixedFormat(0,file)
        workbook.Close()

        messagebox.showinfo("Success", "Conversion done successfully!")
    except:
        messagebox.showerror("Error", "Fatal Error!")

#WORD
def word():
    file = filedialog.askopenfilename(initialdir="/",title="Choose a file")
    directory = filedialog.askdirectory(initialdir="/",title="Choose a directory")
    try:

        convert(file,directory)
        messagebox.showinfo("Success", "Conversion done successfully!")

    except:
        messagebox.showerror("Error", "Fatal Error!")
        
        
#WINDOW
root = tk.Tk()
root.title("ExcelToPdf")
root.tk.call('wm', 'iconphoto', root._w, tk.PhotoImage(file='./icon.png'))
root.geometry("500x500")
root.resizable(False,False)

frame = tk.Frame(root).grid(row=0, column=0, sticky="nw")
background = ImageTk.PhotoImage(file="./bg.png")

#BACKGROUND
global widget_background
widget_background = tk.Label(frame, image=background)
widget_background.image = background
widget_background.grid(row=0, column=0, sticky="nw")

#COMBOBOX
format_selected = StringVar()
format_combobox = ttk.Combobox(root, textvariable=format_selected)
format_combobox.grid(row=0, column=0,sticky="nw")
format_combobox.place(x=180,y=340)
format_combobox['values'] = ["excel","word"]
format_combobox['state'] = 'readonly'

#DOWNLOAD_BUTTON
btn = ImageTk.PhotoImage(file="./download_btn.png")
widget_btn = tk.Button(frame, image=btn,command=control_format, bg="Red")
widget_btn.image = btn
widget_btn.place(x=150, y=380)




root.mainloop()