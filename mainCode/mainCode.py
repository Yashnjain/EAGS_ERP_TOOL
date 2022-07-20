import tkinter as tk
from tkinter import ttk
from datetime import datetime, date
from tkinter import messagebox, Tk
from PIL import Image, ImageTk
from sfTool import get_connection, loginChecker
from tkinter.messagebox import showerror
from generalQuote import quoteGenerator
import sys, traceback
from Tools import resource_path

today = datetime.strftime(date.today(), format = "%d%m%Y")

S_TABLE = "EAGS_SALESPERSON"








class App():

    def __init__(self,root):
        
        # super().__init__()
        #Get Snowflake Connection
        conn = get_connection()
        root.withdraw()
        global user
        
        # self = Tk()
        root.title('EAGS Quote Generator')
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        width = 830
        height = 800
        # calculate position x and y coordinates
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        root.geometry('%dx%d+%d+%d' % (width, height, x, y))
        # root.geometry('648x696')
        biourjaLogo = resource_path('biourjaLogo.png')
        photo = tk.PhotoImage(file = biourjaLogo)
        root.iconphoto(False, photo)
        root["bg"]= "white"
        mFrame = tk.Frame(width=width, height=height, background=root["bg"])
        mFrame.place(in_=root, anchor="c", relx=.5, rely=.5)
        ################Main Window########################
        entry1_path = resource_path("Entry1.png")
        image1 = Image.open(entry1_path)
        image1 = image1.resize((420,199), Image.ANTIALIAS)
        top_img = ImageTk.PhotoImage(image1)

        entry2_path = resource_path("Entry2.png")
        image2 = Image.open(entry2_path)
        image2 = image2.resize((204,423), Image.ANTIALIAS)
        left_img = ImageTk.PhotoImage(image2)

        entry3_path = resource_path("Entry3.png")
        image3 = Image.open(entry3_path)
        image3 = image3.resize((204,423), Image.ANTIALIAS)
        right_img = ImageTk.PhotoImage(image3)

        entry4_path = resource_path("Entry4.png")
        image4 = Image.open(entry4_path)
        image4 = image4.resize((420,199), Image.ANTIALIAS)
        bottom_img = ImageTk.PhotoImage(image4)


        center_img_path = resource_path("center.png")
        center_img = Image.open(center_img_path)
        center_img = center_img.resize((285,285), Image.ANTIALIAS)
        cent_photo = ImageTk.PhotoImage(center_img)



        cent_Lable = tk.Label(mFrame,image=cent_photo,borderwidth=0,bg=root["bg"])
        cent_Lable.place(x=270, y=240)

        left_button = tk.Button(mFrame, image=left_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"])#,command=my_command)
        # left_button.grid(row=1,column=0,padx=0)
        left_button.place(x=50,y=170)

        right_button = tk.Button(mFrame, image=right_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"])
        right_button.place(x=570, y=170)
        top_button = tk.Button(mFrame, image=top_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"],command=lambda:quoteGenerator(root,user,conn))
        top_button.place(x=200, y=20)
        bottom_button = tk.Button(mFrame, image=bottom_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"])
        bottom_button.place(x=200, y=540)
        # bottom_button.place(x=200, y=540,relx=0.2,rely=0.2)

        #############Login Window Section#############################
        top = tk.Toplevel(root)
        top.title('EAGS Quote Generator Login')
        # top.geometry('320x180')
        width2 = 320
        height2 = 180
        x2 = (screen_width/2) - (width2/2)
        y2 = (screen_height/2) - (height2/2)
        top.geometry('%dx%d+%d+%d' % (width2, height2, x2, y2))
        top.iconphoto(True, photo)
        top["bg"]= "white"
 
        # frame_options = ttk.Frame(root)
       
        s = ttk.Style()
        s.configure("TMenubutton", background="#f5fcfc",width=19, font=("Book Antiqua", 12))
        s.configure("TMenu", width=19)
        s.configure("TFrame", background="white")

        def report_callback_exception(self, exc, val, tb):
            msg = traceback.format_exc()
            showerror("Error", message=msg)
            root.update()
        def on_closing():
            try:
                if messagebox.askokcancel("Quit", "Do you want to quit?"):
                    top.destroy()
                    root.destroy()
                    sys.exit()
            except Exception as e:
                raise e
        
        def login(root,top):
            try:
            
        
                def login_command():
                    try:
                        user_dict = {"Imam":"Biourja@2022"}
                        global user
                        user = username.get()
                        pwd = password.get()
                        # if username.get() in user_dict.keys() and password.get() == user_dict[username.get()]:
                        #         global user
                        #         user = username.get()
                        user = loginChecker(conn,S_TABLE, user, pwd)
                        if user:
                                root.deiconify() #Unhides the root window
                                root.state('zoomed')
                                top.destroy()
                                # top.wait_window()
                                
                                
                                # user = username.get().copy()
                        else:
                                password.delete(0, tk.END)
                                messagebox.showinfo("Invalid Credentials", "Username or password is incorrect")
                                password.focus()
                    except Exception as e:
                        raise e

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
                top_frame.focus()
                username.focus()
            except Exception as e:
                raise e
        
        # root = root
        user = login(root,top)
        Tk.report_callback_exception = report_callback_exception

        top.protocol("WM_DELETE_WINDOW", on_closing)
        root.protocol("WM_DELETE_WINDOW", on_closing)

        # Tk.report_callback_exception = report_callback_exception
        
        root.mainloop()




def main():
    root = tk.Tk()
    app = App(root)
    


if __name__ == '__main__':
    main()









