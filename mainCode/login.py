import tkinter as tk
from tkinter import ttk
from tkinter import messagebox



def login(root,top,user):

        
        
        def login_command():
                
                user_dict = {"Imam":"Biourja@2022"}
                
                if username.get() in user_dict.keys() and password.get() == user_dict[username.get()]:
                        # global user
                        user = username.get()
                        root.deiconify() #Unhides the root window
                        root.state('zoomed')
                        top.destroy()
                        # top.wait_window()
                        
                        
                        # user = username.get().copy()
                else:
                        password.delete(0, tk.END)
                        messagebox.showinfo("Invalid Credentials", "Username or password is incorrect")

        top_frame = ttk.Frame(top)
                
        top_frame.grid(row=0, column=1,pady=(24,0),columnspan=3, padx=(10,0))
        user_label = ttk.Label(top_frame, text="Username:", font=("Book Antiqua bold", 12), foreground="#ff8c00", background="white")
        user_label.grid(row=0, column=0)
        user_text = tk.StringVar()
        username = ttk.Entry(top_frame, textvariable=user_text) #Username entry
        # user_text.set("IMAM")
        username.grid(row=0, column=1)
        password_label = ttk.Label(top_frame, text="Password:", font=("Book Antiqua bold", 12), foreground="#ff8c00", background="white")
        password_label.grid(row=1, column=0, pady=10)
        password_text = tk.StringVar()
        password = ttk.Entry(top_frame, show="*", textvariable=password_text) #Password entry
        password.grid(row=1, column=1, pady=10)
        # password_text.set("Biourja@2021#7515")

        loginButton_text = tk.StringVar()
        loginButton = tk.Button(top_frame, textvariable=loginButton_text, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, command=login_command, activebackground="#20bebb")
        loginButton.grid(row=2, column=1, pady=30)
        loginButton_text.set("Login")

        loginButton.bind("<Return>", (lambda event: login_command()))
        password.bind("<Return>", (lambda event: login_command()))
        return user



        # if len(password.get())>0:
        # top.bind('<Return>', login_command)
        # loginButton.bind("<Return>", (lambda event: login_command()))
        # password.bind("<Return>", (lambda event: login_command()))






        # top.protocol("WM_DELETE_WINDOW", on_closing)