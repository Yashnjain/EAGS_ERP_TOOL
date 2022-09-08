import tkinter
import customtkinter
from tkinter import *
from tkinter import messagebox, Tk
import sys,traceback
from tkinter.messagebox import showerror
from RsfTool import get_connection, getallquotes, getfullquote
import time
from datetime import date,datetime
from tkinter import ttk
import speech_recognition as SRG 
from general_quote_revision import general_quote_revision
from Rbaker import bakerQuoteGenerator

customtkinter.set_appearance_mode("system")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green

def quoteRevision(root,user,conn,inv_df):
    try:
        # root.withdraw()
        conn = get_connection()
        # quoteList = Toplevel(root)
        def getquote(conn,quoteList,quote_number):
            quotedf=getfullquote(conn,quote_number)
            quoteList.destroy()
            if 'Baker' in quotedf['QUOTENO'][0]:
                bakerQuoteGenerator(root,user,conn,quotedf,quote_number, inv_df)
            else:    
                general_quote_revision(root,user,conn,quotedf,quote_number, inv_df)
            print("sdfgvsd")

        def speech_recognizer(): 
            store = SRG.Recognizer()
            with SRG.Microphone() as s:
                
                print("Speak...")
                
                audio_input = store.record(s, duration=3)
                print("Recording time:",time.strftime("%I:%M:%S"))
                
                try:
                    text_output = store.recognize_google(audio_input)
                    print("Text converted from audio:\n")
                    print(text_output)
                    print("Finished!!")
                    print("Execution time:",time.strftime("%I:%M:%S"))
                    if 'stroke' in text_output:
                        text_output.replace('stroke','/')
                    entry.delete(0,END)    
                    entry.insert(0,text_output)  
                except:
                    print("Couldn't process the audio input.")



        # def report_callback_exception(exc, val, tb):
        #     msg = traceback.format_exc()
        #     showerror("Error", message=msg)
        #     root.update()

        def ChangeTheme(mode):
                customtkinter.set_appearance_mode(mode)
                if mode == 'Light':
                    customtkinter.set_appearance_mode("light")
                elif mode == 'Dark':
                    customtkinter.set_appearance_mode("dark")
                else:
                    customtkinter.set_appearance_mode("system") 

        def on_closing():
            try:
                if messagebox.askokcancel("Quit", "Do you want to quit?"):
                    mFrame.destroy()
                    sys.exit()
            except Exception as e:
                raise e

        def deleter(e):
            if var.get()=="Enter quote number:":
                label_text=var.get()
                entry.delete(0,END)
                # mylabel2=customtkinter.CTkLabel(mFrame,text=label_text,text_font=("Book Antiqua bold", 12))
                # mylabel2.grid(row=0,column=0,pady=5)
                label_1=customtkinter.CTkLabel(mFrame,text=label_text.upper(),text_font=("Roboto Medium", -16))  # font name and size in px
                label_1.grid(row=0, column=0, pady=10)
                # entry.config(text_font='Arial')

        def hh(e=None):
            if var.get()!="Enter quote number:":
                if var.get() == "":
                    messagebox.showerror(title="ENTRY ERROR",message="Please enter a QUOTE NUMBER first")
                    mylabel.config(text = "")
                    mFrame.focus()
                    entry.focus()
                    return
                text="PROCESSING ALL QUOTES FOR" + " "+ entry.get()
                mylabel=customtkinter.CTkLabel(mFrame,text=text,text_font='Calibri')
                mylabel.grid(row=3,column=0, columnspan=2)#,sticky='ew')yash
                quote_number= entry.get()
                quote_numbers=getallquotes(conn,quote_number)
                print(user)
                if len(quote_numbers):
                    if (user[0].upper() == (quote_numbers['PREPAREDBY'].astype(str)).unique()[0].upper()) or user[1].upper()=="ADMIN":    
                        if len(quote_numbers)==0:
                            messagebox.showerror(title="Incorrect QUOTE NUMBER",message="Please enter correct QUOTE NUMBER")
                            mylabel.config(text = "")
                            mFrame.focus()
                            entry.focus()
                            return
                        # list_of_quotes=list(quote_numbers[['QUOTENO','DATE']].unique())
                        list_of_quotes=list(((quote_numbers['QUOTENO']+"|"+quote_numbers['DATE'].astype(str)).astype(str)).unique())
                        # list_of_dates=list(quote_numbers['DATE'].unique())
                        list_of_quotes=list_of_quotes[::-1]
                        print(quote_numbers)
                        quoteSearcher.withdraw()
                        quoteList = customtkinter.CTkToplevel(quoteSearcher)
                        quoteList.title("QUOTE REVISION")
                        quoteList["bg"]= "#e2e1ef"
                        height=int((len(list_of_quotes)*100)/1.5)
                        # quoteList.geometry(f"400x{height}")

                        screen_width = root.winfo_screenwidth()
                        screen_height = root.winfo_screenheight()

                        width = 400
                        # calculate position x and y coordinates
                        x = (screen_width/2) - (width/2)
                        y = (screen_height/2) - (height/2)
                        quoteList.geometry('%dx%d+%d+%d' % (width, height, x, y))
                        # quoteList.eval('tk::PlaceWindow . center')
                        # quoteSearcher.geometry(f'{400}x{height}+%d+%d' % (x2, y2))
                        quoteList.wm_maxsize(width=400,height=height)
                        dict1={}
                        for index,quotes in enumerate(list_of_quotes):
                            quotes_no=quotes.split("|")[0]
                            # style = ttk.Style()
                            # style.theme_use('alt')
                            # style.configure('TButton', text_font=('American typewriter', 14 ,'bold','underline'), background='#232323', foreground='white')
                            # style.map('TButton', background=[('active', '#ff0000'), ('disabled', '#f0f0f0')])
                            # quote_number=quotes
                            # quotes=customtkinter.CTkButton(quoteList,text=quotes,state=tkinter.NORMAL,command=lambda:getquote(conn,quotes,quoteList))#'#state=ACTIVE,fg='#b31bea',bg='white',text_font=('American typewriter',14)
                            dict1[quotes]=customtkinter.CTkButton(quoteList,text=quotes,state=tkinter.NORMAL,command=lambda quote_number = quotes_no:getquote(conn,quoteList,quote_number))
                            dict1[quotes].grid(row=index+1,column=0,pady=5)
                            quoteList.grid_rowconfigure(index=index+1, weight=1)
                        customtkinter.CTkLabel(quoteList, text="SELECT QUOTE FOR REVISION : ",text_font=("Roboto Medium", -16),text_color="black").grid(row=0,column=0)   #,text_font=('Consolas 15',14) 
                        quoteList.grid_rowconfigure(0, weight=1)
                        quoteList.grid_columnconfigure(0, weight=1)   
                        quoteList.mainloop()
                    else:
                        messagebox.showinfo("ERROR", "You are not authorised to access this QUOTE.\nPlease contact your Admin")
                        mylabel.config(text = "")
                        mFrame.focus()
                        entry.focus()

                        return
                else:
                    messagebox.showinfo("ERROR", "Quote number doesn't exist")
                    mylabel.config(text = "")
                    mFrame.focus_set()
                    entry.focus_set()
                    

            else:
                messagebox.showerror(title="ENTRY ERROR",message="Please enter a QUOTE NUMBER first")
                mylabel.config(text = "")
                mFrame.focus()
                entry.focus()

                return      
        quoteSearcher = customtkinter.CTkToplevel(root)
        quoteSearcher.iconbitmap()
        width2 = 420
        height2 = 190
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x2 = (screen_width/2) - (width2/2)
        y2 = (screen_height/2) - (height2/2)
        quoteSearcher.geometry('%dx%d+%d+%d' % (width2, height2, x2, y2))
        # quoteSearcher.geometry("350x200")
        quoteSearcher.wm_maxsize(width=350,height=200)
        #settings frame for app settings
        settings_frame = customtkinter.CTkFrame(quoteSearcher, width=50)
        # settings_frame.grid(row=3,column=0)
        settings_frame.pack(fill=X, side=tkinter.TOP, padx=2, pady=2)
        settings_frame.grid_columnconfigure(0, weight=1)
        settings_frame.grid_rowconfigure(3, weight=1)
        quoteSearcher.title("QUOTE REVISION")
        #for icon
        theme_menu = customtkinter.CTkOptionMenu(settings_frame, values=['Dark','Light','System'], text_font=('Consolas 15',14), command=ChangeTheme)
        theme_menu.grid(row=0, column=3, padx=5, pady=5, sticky='nswe')
        mFrame = customtkinter.CTkFrame(quoteSearcher)
        mFrame.pack(fill=X, side=tkinter.TOP, padx=5, pady=5)
        # mFrame["bg_color"]= "#e2e1ef"
        # biourjaLogo = resource_path('biourjaLogo.png')
        # photo = tk.PhotoImage(file = biourjaLogo)
        # root.iconphoto(False, photo)
        # quoteSearcher["bg_color"]= "#e2e1ef"
        global click_btn
        click_btn= PhotoImage(file='sound1.png')
        click_btn = click_btn.subsample(15, 15)
        speech_btn = customtkinter.CTkButton(mFrame,text='',image=click_btn, command=lambda:speech_recognizer(),hover_color='#2833cd')
        speech_btn.grid(row=1,column=1)
        var =StringVar()
        entry=customtkinter.CTkEntry(mFrame,text_font=("Roboto Medium", 12 ), textvariable=var,width=220)
        entry.grid(row=1,column=0)
        entry.insert(0,"Enter quote number:")
        entry.bind('<1>', deleter)
        entry.bind('<Return>', hh)
        
        style = ttk.Style()
        # style.theme_use('alt')
        style.configure('TButton', text_font=('American typewriter', 14), background='#232323', foreground='white')
        style.map('TButton', background=[('active', '#ff0000'), ('disabled', '#f0f0f0')])          
        mybutn=customtkinter.CTkButton(mFrame,text='Submit',state=NORMAL,command=lambda:hh(),text_font=('American typewriter', 14),hover_color='#2833cd')#,bg_color=mFrame['bg'])  #,fg='#b31bea',bg='white',#bg_color='#232323',fg_color='white',
        mybutn.grid(row=2,column=0,pady=5)#,sticky='ew',padx=5,pady=5
        quoteSearcher.grid_rowconfigure(0, weight=1)
        quoteSearcher.grid_columnconfigure(0, weight=1)
        mFrame.grid_rowconfigure(index=0, weight=1)
        mFrame.grid_rowconfigure(index=1, weight=1)
        mFrame.grid_rowconfigure(index=2, weight=1)
        mFrame.grid_rowconfigure(index=3, weight=1)
        mFrame.grid_columnconfigure(index=0, weight=1)
        mFrame.grid_columnconfigure(index=1, weight=1)
        mFrame.focus()
        mybutn.focus()
        # quoteSearcher.protocol("WM_DELETE_WINDOW", on_closing)
        mFrame.mainloop()
        quoteSearcher.mainloop()
        # Tk.report_callback_exception = report_callback_exception
        # root.destroy()
        # quoteSearcher()
        root.mainloop()
    except Exception as e:
        raise e




        # app.geometry("400x240")

        # def button_function():
        #     print("button pressed")

        # # Use CTkButton instead of tkinter Button
        # button = customtkinter.CTkButton(master=app, text="CTkButton", command=button_function)
        # button.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
# root = customtkinter.CTk() 
# conn = get_connection()
# user = "imam" # create CTk window like you do with the Tk window        
# quoteRevision(root,user,conn)
# app.mainloop()


#C:\Users\Yashn.jain\AppData\Roaming\Python\Python38\Scripts
