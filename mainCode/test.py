import tkinter as tk
from tkinter import messagebox

def quit(window):
    if not popup:  # destroy the window only if popup is not displayed
        window.destroy()

def show():
    global popup
    popup = True
    root.attributes('-topmost', True)
    messagebox.showinfo("Test Popup", "Hello world", parent=root)
    root.attributes('-topmost', False)
    popup = False

root = tk.Tk()
popup = False
root.protocol("WM_DELETE_WINDOW", lambda: quit(root))
root.title("Main Window")
root.geometry("500x500")

toplevel = tk.Toplevel(root)
toplevel.protocol("WM_DELETE_WINDOW", lambda: quit(toplevel))
toplevel.title("Toplevel Window")

show_button = tk.Button(root, text="Show popup", command=show)
show_button.pack()

root.mainloop()

def show_pop(root, title, message): 
    root.attributes('-topmost', True)
    messagebox.showinfo(title, message, parent=root)
    root.attributes('-topmost', False)