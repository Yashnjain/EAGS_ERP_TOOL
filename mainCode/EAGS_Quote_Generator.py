import tkinter as tk
from tkinter import ttk
from datetime import datetime, date
from tkinter import messagebox, Tk
from PIL import Image, ImageTk
from sfTool import get_connection, loginChecker, get_inv_df
from tkinter.messagebox import showerror
from generalQuote import quoteGenerator
from baker import bakerQuoteGenerator
import sys, traceback
import numpy as np
import tkcap, os
from mail import send_mail
from Tools import resource_path
from quote_revision_final import quoteRevision
from appUpdaterV2 import appUpdater
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)
from eagsReport import reportGenerator

today = datetime.strftime(date.today(), format = "%d%m%Y")

S_TABLE = "EAGS_SALESPERSON"
INV_TABLE = "EAGS_INVENTORY"

VERSION = "0.0.0.3" 
#Version format:
#Major revision (new UI, lots of new features, conceptual change, etc.)
#Minor revision (maybe a change to a search box, 1 feature added, collection of bug fixes)
#Bug fix release
#Build number





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
        def on_enter(e):
            try:
                e.widget['image'] = button_dict[e.widget][1]
            except Exception as e:
                raise e
            # addRowbut['background'] = 'green'

        def on_leave(e):
            try:
                e.widget['image'] = button_dict[e.widget][0]
            except Exception as e:
                raise e
        def resize_image(event):
            new_width = event.width
            new_height = event.height
        print(os.path.abspath(os.path.dirname(sys.argv[0])))
        entry1_path = resource_path("Entry1.png")
        w_scale_factor = screen_width/1920
        h_scale_factor = screen_height/1080
        image1 = Image.open(entry1_path)
        # image1 = image1.resize((int(420*w_scale_factor),int(199*h_scale_factor)), Image.Resampling.LANCZOS)
        image1 = image1.resize((420,199), Image.Resampling.LANCZOS)
        top_img = ImageTk.PhotoImage(image1)

        entry1New_path = resource_path("Entry1New.png")
        image1New = Image.open(entry1New_path)
        # image1New = image1New.resize((int(420*w_scale_factor),int(199*h_scale_factor)), Image.Resampling.LANCZOS)
        image1New = image1New.resize((420,199), Image.Resampling.LANCZOS)
        top_imgNew = ImageTk.PhotoImage(image1New)

        entry2_path = resource_path("Entry2.png")
        image2 = Image.open(entry2_path)
        # image2 = image2.resize((int(204*w_scale_factor),int(423*h_scale_factor)), Image.Resampling.LANCZOS)
        image2 = image2.resize((204,423), Image.Resampling.LANCZOS)
        left_img = ImageTk.PhotoImage(image2)

        entry2New_path = resource_path("Entry2New.png")
        image2New = Image.open(entry2New_path)
        # image2New = image2New.resize((int(204*w_scale_factor),int(423*h_scale_factor)), Image.Resampling.LANCZOS)
        image2New = image2New.resize((204,423), Image.Resampling.LANCZOS)
        left_imgNew = ImageTk.PhotoImage(image2New)

        entry3_path = resource_path("Entry3.png")
        image3 = Image.open(entry3_path)
        # image3 = image3.resize((int(204*w_scale_factor),int(423*h_scale_factor)), Image.Resampling.LANCZOS)
        image3 = image3.resize((204,423), Image.Resampling.LANCZOS)
        right_img = ImageTk.PhotoImage(image3)

        entry3New_path = resource_path("Entry3New.png")
        image3New = Image.open(entry3New_path)
        # image3New = image3New.resize((int(204*w_scale_factor),int(423*h_scale_factor)), Image.Resampling.LANCZOS)
        image3New = image3New.resize((204,423), Image.Resampling.LANCZOS)
        right_imgNew = ImageTk.PhotoImage(image3New)

        entry4_path = resource_path("Entry4.png")
        image4 = Image.open(entry4_path)
        # image4 = image4.resize((int(420*w_scale_factor),int(199*h_scale_factor)), Image.Resampling.LANCZOS)
        image4 = image4.resize((420,199), Image.Resampling.LANCZOS)
        bottom_img = ImageTk.PhotoImage(image4)

        entry4New_path = resource_path("Entry4New.png")
        image4New = Image.open(entry4New_path)
        # image4New = image4New.resize((int(420*w_scale_factor),int(199*h_scale_factor)), Image.Resampling.LANCZOS)
        image4New = image4New.resize((420,199), Image.Resampling.LANCZOS)
        bottom_imgNew = ImageTk.PhotoImage(image4New)


        button_dict = {}

        center_img_path = resource_path("center.png")
        center_img = Image.open(center_img_path)
        # center_img = center_img.resize((int(285*w_scale_factor),int(285*h_scale_factor)), Image.Resampling.LANCZOS)
        center_img = center_img.resize((285,285), Image.Resampling.LANCZOS)
        cent_photo = ImageTk.PhotoImage(center_img)

        #########Adding report generator button#################
        report_img_path1 = resource_path("reportGenerator1.png")
        report_img1 = tk.PhotoImage(master=mFrame, file=report_img_path1)

        report_img_path2 = resource_path("reportGenerator2.png")
        report_img2 = tk.PhotoImage(master=mFrame, file=report_img_path2)

        reportbut = tk.Button(mFrame, image=report_img1, command=lambda:reportGenerator(root, conn),borderwidth=0, background=root["bg"],activebackground=root["bg"])
        reportbut.image = report_img1
        reportbut.place(x=620, y=20)

        button_dict[reportbut] = [report_img1, report_img2]
        reportbut.bind("<Enter>", on_enter)
        reportbut.bind("<Leave>", on_leave)
        ###############################################################

        cent_Lable = tk.Label(mFrame,image=cent_photo,borderwidth=0,bg=root["bg"])
        cent_Lable.place(x=270, y=240)#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        # cent_Lable.place(x=int(270*w_scale_factor), y=int(240*h_scale_factor))#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)

        left_button = tk.Button(mFrame, image=left_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"],command=lambda:bakerQuoteGenerator(root,user[0],conn, inv_df))#,command=my_command)
        # left_button.grid(row=1,column=0,padx=0)
        left_button.bind('<Configure>', resize_image)
        left_button.place(x=50,y=170)#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        # left_button.place(x=int(50*w_scale_factor),y=int(170*h_scale_factor))#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        
        button_dict[left_button] = [left_img, left_imgNew]
        left_button.bind("<Enter>", on_enter)
        left_button.bind("<Leave>", on_leave)

        right_button = tk.Button(mFrame, image=right_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"])
        right_button.place(x=570, y=170)#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        # right_button.place(x=int(570*w_scale_factor), y=int(170*h_scale_factor))#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        button_dict[right_button] = [right_img, right_imgNew]
        right_button.bind("<Enter>", on_enter)
        right_button.bind("<Leave>", on_leave)


        top_button = tk.Button(mFrame, image=top_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"],command=lambda:quoteGenerator(root,user[0],conn, inv_df))
        top_button.place(x=200, y=20)#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        # top_button.place(x=int(200*w_scale_factor), y=int(20*h_scale_factor))#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        button_dict[top_button] = [top_img, top_imgNew]
        top_button.bind("<Enter>", on_enter)
        top_button.bind("<Leave>", on_leave)
        
        bottom_button = tk.Button(mFrame, image=bottom_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"], command = lambda:quoteRevision(root,user,conn, inv_df))
        bottom_button.place(x=200, y=540)#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        # bottom_button.place(x=int(200*w_scale_factor), y=int(540*h_scale_factor))#, relx=(1-w_scale_factor)/10, rely=(1-h_scale_factor)/10)
        button_dict[bottom_button] = [bottom_img, bottom_imgNew]
        bottom_button.bind("<Enter>", on_enter)
        bottom_button.bind("<Leave>", on_leave)
        # bottom_button.place(x=200, y=540,relx=0.2,rely=0.2)

        #############Login Window Section#############################
        top = tk.Toplevel(root)
        top.title('EAGS Quote Generator Login')
        # top.geometry('320x180')
        width2 = 420
        height2 = 190
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
            msg = msg.replace('\n', '<br>')
            nl = '<br>'
            error_id = np.random.randint(1000000,9999999)
            tempDir = os.path.join(os.environ["HOMEPATH"], "Temp")
            tempDir = os.path.join('C:', tempDir)

            if not os.path.exists(tempDir):
                os.mkdir(tempDir)
            # directories_created = [tempDir]
            # for directory in directories_created:
            #     path3 = os.path.join(os.getcwd(),directory)  
            #     try:
            #         os.makedirs(path3, exist_ok = True)
            #         print("Directory '%s' created successfully" % directory)
            #     except OSError as error:
            #         print("Directory '%s' can not be created" % directory) 

            cap = tkcap.CAP(root)     # master is an instance of tkinter.Tk
            imageV1path = os.path.join(tempDir,f'{error_id}V1.png')
            # cap.capture(f'{error_id}V1.png')
            cap.capture(imageV1path)
            # imageV1path = os.getcwd()+'\\'+f'{error_id}V1.png'
            

            
            # dsp_msg = f"Error: {error_id}\nPlease send a screenshot of this error message along with the app window to devsupport@biourja.com"
            dsp_msg = f"A Report for this issue has been sent to IT India Team with Error ID: {error_id}, it will be resolved soon\nPlease capture this error id for future reference"
            showerror(f"Error", message=dsp_msg)
            imageV2path = os.path.join(tempDir,f'{error_id}V2.png')
            cap = tkcap.CAP(root)     # master is an instance of tkinter.Tk
            cap.capture(imageV2path)       # Capture and Save the screenshot of the tkiner window
            # imageV2path = os.getcwd()+'\\'+f'{error_id}V2.png'
            
            
            send_mail(receiver_email='imam.khan@biourja.com, yashn.jain@biourja.com, devsupport@biourja.com', mail_subject=f"EAGS APP ERROR FOUND {user[0]} {error_id}", 
            mail_body=f"<strong>User: {user[0]}{nl}Error ID: {error_id}</strong>{nl}{msg}", attachment_locations=[imageV1path, imageV2path])
            
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
                def on_enter(e):
                    try:
                        e.widget['bg'] = "#00008B"
                    except Exception as e:
                        raise e
                    # addRowbut['background'] = 'green'

                def on_leave(e):
                    try:
                        e.widget['bg'] = "#20bebe"
                    except Exception as e:
                        raise e

        
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
                                global inv_df
                                loginButton_text.set("Logging In...")
                                root.update()  
                                inv_df = get_inv_df(conn,table = INV_TABLE)
                                root.deiconify() #Unhides the root window
                                root.state('zoomed')
                                top.destroy()
                                filePath = os.path.abspath(__file__)
                                fileDirectory = os.path.dirname(__file__) # relative directory path
                                fileName = os.path.basename(__file__) # the file name only
                                messagebox.showinfo("Current App Location", f"Current file path is {filePath} Version changed to {VERSION}")
                                #sys.exit()
                                # top.wait_window()
                                
                                
                                # user = username.get().copy()
                        else:
                                password.delete(0, tk.END)
                                messagebox.showinfo("Invalid Credentials", "Username or password is incorrect")
                                password.focus()
                    except Exception as e:
                        raise e

                top_frame = ttk.Frame(top)
                       
                top_frame.grid(row=0, column=1,pady=(24,0),columnspan=3, padx=(10,0))#Book Antiqua
                user_label = ttk.Label(top_frame, text="Username:", font=("Segoe UI bold", 12), foreground='black', background="white")#"#ff8c00"
                user_label.grid(row=0, column=0)
                user_text = tk.StringVar()
                username = ttk.Entry(top_frame, textvariable=user_text) #Username entry
                # user_text.set("IMAM")
                username.grid(row=0, column=1)
                password_label = ttk.Label(top_frame, text="Password:", font=("Segoe UI bold", 12), foreground='black', background="white")#"#ff8c00"
                password_label.grid(row=1, column=0, pady=10)
                password_text = tk.StringVar()
                password = ttk.Entry(top_frame, show="*", textvariable=password_text) #Password entry
                password.grid(row=1, column=1, pady=10)
                # password_text.set("Biourja@2021#7515")

                loginButton_text = tk.StringVar()
                loginButton = tk.Button(top_frame, textvariable=loginButton_text, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, command=login_command, activebackground="#20bebb", highlightbackground="#20bebd")
                loginButton.grid(row=2, column=1, pady=30)
                loginButton_text.set("Login")
                loginButton.bind("<Enter>", on_enter)
                loginButton.bind("<Leave>", on_leave)

                loginButton.bind("<Return>", (lambda event: login_command()))
                password.bind("<Return>", (lambda event: login_command()))
                top_frame.focus_set()
                username.focus()
            except Exception as e:
                raise e
        
        def oldExeDeleter(fileDirectory):
            try:
                
                if os.path.exists(fileDirectory+"\\EAGS_Quote_Generator_Old.exe"):
                    os.remove(fileDirectory+"\\EAGS_Quote_Generator_Old.exe")
            except Exception as e:
                raise e
        # root = root
        
        filePath = os.path.abspath(__file__)
        fileDirectory = os.path.dirname(__file__)
        fileName = os.path.basename(__file__) # the file name only
        print(filePath)
        oldExeDeleter(fileDirectory)
        if appUpdater(root,photo, curr_version=VERSION, curr_location=filePath, curr_directory=fileDirectory, currFilename=fileName):
            sys.exit()
        else:
            pass
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









