from tkinter import *
import tkinter as tk
from geopy.geocoders import Nominatim
from tkinter import ttk,messagebox
from ctypes import windll
import customtkinter as ctk
from datetime import datetime
import requests
import pandas as pd
from pandastable import Table


root = Tk()

# root.attributes('-fullscreen', False)
root.title("No Order Why")
# root.geometry("900x500+300+200")
width= root.winfo_screenwidth() 
height= root.winfo_screenheight()
#setting tkinter window size
root.geometry("%dx%d" % (width, height))
root.resizable(False,False)
root.configure(bg='#E5E4E2')

    # Create the main window
# root = tk.Tk()
# root.title("Three Frames Example")

# Create Frame 1
frame1 = tk.Frame(root, width=1500, height=100, bg="red")
frame1.grid(row=0, column=0, sticky="ew")

# Create Frame 2 (Middle)
frame2 = tk.Frame(root, width=400, height=200, bg="green")
frame2.grid(row=1, column=0, sticky="ew")

# Create Frame 3 (Bottom)
frame3 = tk.Frame(root, width=400, height=100, bg="blue")
frame3.grid(row=2, column=0, sticky="ew")

# Configure row and column weights to make frames expand vertically
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)

# Start the Tkinter event loop
# root.mainloop()

# Call the function to create frames


def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()



LABEL2 = Label(frame1,text="Prepared By",font=('Helvetica 30 bold',12),bg='#E5E4E2',fg="Black")
LABEL2.place(x=80,y=60)

preparedVar = tk.StringVar()
preparedBox = ttk.Entry(frame1, textvariable=preparedVar, background = 'white',width = 30)
preparedBox.place(x=40,y=90)


LABEL3 = Label(frame1,text="Quote Number and Coustomer Name",font=('Helvetica 30 bold',12),bg='#E5E4E2',fg="Black")
LABEL3.place(x=540,y=60)

quote_customer = tk.StringVar()
quote_customerBox = ttk.Entry(frame1, textvariable=preparedVar, background = 'white',width = 60)
quote_customerBox.place(x=490,y=90)

LABEL4 = Label(frame1,text="Sales Person Name",font=('Helvetica 30 bold',12),bg='#E5E4E2',fg="Black")
LABEL4.place(x=980,y=60)

salesperson = tk.StringVar()
salespersonBox = ttk.Entry(frame1, textvariable=preparedVar, background = 'white',width = 30)
salespersonBox.place(x=960,y=90)


LABEL5 = Label(frame1,text="Conversion Status",font=('Helvetica 30 bold',12),bg='#E5E4E2',fg="Black")
LABEL5.place(x=1200,y=60)

status = tk.StringVar()
statusBox = ttk.Entry(frame1, textvariable=preparedVar, background = 'white',width = 30)
statusBox.place(x=1180,y=90)




root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()