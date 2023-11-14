import tkinter as tk
from tkinter import messagebox
from tkinter import ttk,Label,Entry,StringVar,Tk
from tkcalendar import DateEntry
from datetime import date
from sfTool import get_connection, upload_no_order_why, get_master_df
from customComboboxForReport import myCombobox
from datetime import datetime
import os
import pandas as pd
# import ChecklistCombobox

MASTER_TABLE = "EAGS_MASTER"


#Calendar
class MyDateEntry(DateEntry):
    try:
        def __init__(self, master=None, **kw):
            DateEntry.__init__(self, master=master, date_pattern='dd.mm.yyyy',**kw)
            # add black border around drop-down calendar
            self._top_cal.configure(bg='black', bd=1)
            # add label displaying today's date below
            tk.Label(self._top_cal, bg='gray90', anchor='w',
                    text='Today: %s' % date.today().strftime('%x')).pack(fill='both', expand=1)
    except Exception as e:
        raise e
# def eagsReport(root):
#     #fetch searcher fields Username, Location, Grade, Yield, Date Range
#     #based on that pull quotation and baker data
#     try:
#         toproot = tk.Toplevel(root, bg = "#9BC2E6")
#         toproot.title('EAGS Quote Generator Star Search')
#         screen_width = toproot.winfo_screenwidth()
#         screen_height = toproot.winfo_screenheight()

#         width = 315
#         height = 280
#         # calculate position x and y coordinates
#         x = (screen_width/2) - (width/2)
#         y = (screen_height/2) - (height/2)
#         toproot.geometry('%dx%d+%d+%d' % (width, height, x, y))

#         #Declaring Main Labels
#         labelFrame = tk.Frame(toproot, bg= "#9BC2E6")
#         labelFrame.grid(row=0, column=1)
#         buttonFrame= tk.Frame(toproot, bg= "#9BC2E6")
#         buttonFrame.grid(row=1, column=1)

#         #Declaring Sublabels
#         userLabel = tk.Label(labelFrame, text="Username", bg = "#9BC2E6")
#         userLabel.grid(row=0, column=0)

#         locationLabel = tk.Label(labelFrame, text="Location", bg = "#9BC2E6")
#         locationLabel.grid(row=0, column=0)

#         gradeLabel = tk.Label(labelFrame, text="Grade", bg = "#9BC2E6")
#         gradeLabel.grid(row=0, column=0)

#         #Configuring frame sizes based on screen size
#         toproot.grid_rowconfigure(0, weight=1)
#         toproot.grid_columnconfigure(0, weight=1)
#         toproot.grid_rowconfigure(1, weight=1)
#         toproot.grid_columnconfigure(1, weight=1)

#         labelFrame.grid_rowconfigure(0, weight=1)
#         labelFrame.grid_columnconfigure(0, weight=1)
#         labelFrame.grid_rowconfigure(1, weight=1)
#         labelFrame.grid_columnconfigure(1, weight=1)
#         labelFrame.grid_columnconfigure(2, weight=1)

#         boxFrame.grid_rowconfigure(0, weight=1)
#         boxFrame.grid_columnconfigure(0, weight=1)
#         boxFrame.grid_rowconfigure(1, weight=1)
#         boxFrame.grid_columnconfigure(1, weight=1)

#         entryFrame1.grid_rowconfigure(0, weight=1)
#         entryFrame1.grid_columnconfigure(0, weight=1)
#         entryFrame1.grid_rowconfigure(1, weight=1)
#         entryFrame1.grid_columnconfigure(1, weight=1)

#         entryFrame2.grid_rowconfigure(0, weight=1)
#         entryFrame2.grid_columnconfigure(0, weight=1)
#         entryFrame2.grid_rowconfigure(1, weight=1)
#         entryFrame2.grid_columnconfigure(1, weight=1)

#         submitFrame.grid_rowconfigure(0, weight=1)
#         submitFrame.grid_columnconfigure(0, weight=1)
#         submitFrame.grid_rowconfigure(1, weight=1)
#         submitFrame.grid_columnconfigure(1, weight=1)
        
#     except Exception as e:
#         raise e


#########################################################################################

def no_order_why(root,conn):
    def intFloat(inStr,acttyp):
        try:
            # if acttyp == '1': #insert
            if inStr == '' or inStr == "NA":
                return True
            try:
                float(inStr)
                # print('value:', inStr)
            except ValueError:
                toproot.attributes('-topmost', True)
                messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in Integer or Decimal format only",parent=toproot)
                toproot.attributes('-topmost', False)
                return False
            return True
        except Exception as e:
            raise e
    try:
        def tabFunc(e):
            try:
                priceBox.focus_set()
                return "break"           
            except Exception as e:
                raise e
        def tabFuncPrice(e):
            try:
                TimeBox.focus_set()
                return "break"           
            except Exception as e:
                raise e
        def tabFuncTime(e):
            try:
                DeliveryBox.focus_set()
                return "break"           
            except Exception as e:
                raise e
        global df
        df = get_master_df(conn,table = MASTER_TABLE)
        df.columns = map(str.lower  , df.columns)
        toproot = tk.Toplevel(root, bg = "#9BC2E6")
        toproot.title('No Order Why')
        screen_width = toproot.winfo_screenwidth()
        screen_height = toproot.winfo_screenheight()

        width = 600
        height = 500
        # calculate position x and y coordinates
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        toproot.geometry('%dx%d+%d+%d' % (width, height, x, y))
        toproot.resizable(False,False)

        toproot.grab_set()
      

     
        global cxDatadict
        
        cxDatadict = {}

        quote_no_var = []
        cxDatadict["quoteno"] =[]
        cxDatadict["quoteno"].append(quote_no_var)

        # prep_by_var = []
        cxDatadict["preparedby"] =[]
        # cxDatadict["prep_by"].append(prep_by_var)

        # sales_p_var = []
        cxDatadict["sales_person"] =[]
        # cxDatadict["sales_person"].append(sales_p_var)







    

        labelFrame = tk.Frame(toproot, bg= "#9BC2E6")
        labelFrame.grid(row=0, column=1)
        labelFrame2= tk.Frame(toproot, bg= "#9BC2E6")
        labelFrame2.grid(row=1, column=1)
        entryFrame1 = tk.Frame(labelFrame, bg= "#9BC2E6")
        entryFrame1.grid(row=1, column=0)
        entryFrame2 = tk.Frame(labelFrame, bg= "#9BC2E6")
        entryFrame2.grid(row=1, column=1)
        entryFrame3 = tk.Frame(labelFrame, bg= "#9BC2E6")
        entryFrame3.grid(row=1,column=2)
        # quoteFrame = tk.Frame(toproot, bg= "#9BC2E6")
        # quoteFrame.grid(row=1, column=1)
        submitFrame = tk.Frame(toproot, bg= "#9BC2E6")
        submitFrame.grid(row=2, column=1)
        #Declaring Labels
        # tobePrinted = tk.Label(labelFrame, text="To be Printed", bg = "#9BC2E6")
        # tobePrinted.grid(row=0, column=0)

        # tobeCalculated = tk.Label(labelFrame, text="To be Calculated", bg = "#9BC2E6")
        # tobeCalculated.grid(row=0,column=1)

        Quote_number = tk.Label(entryFrame1, text="Quote Number", bg = "#9BC2E6", font=("Segoe UI", 10))
        Quote_number.grid(row=0, column=0)

        PreparedLabel = tk.Label(entryFrame2, text="Prepared By", bg = "#9BC2E6", font=("Segoe UI", 10))
        PreparedLabel.grid(row=0, column=0)

        salesLabel = tk.Label(entryFrame3, text="Sales Person", bg = "#9BC2E6", font=("Segoe UI", 10))
        salesLabel.grid(row=0, column=0)

        PriceLabel = tk.Label(entryFrame1, text="Price", bg = "#9BC2E6", font=("Segoe UI", 10))
        PriceLabel.grid(row=2, column=0)

        timeLabel = tk.Label(entryFrame2, text="Time", bg = "#9BC2E6", font=("Segoe UI", 10))
        timeLabel.grid(row=2,column=0)

        deliveryLabel = tk.Label(entryFrame3, text="Delivery", bg = "#9BC2E6", font=("Segoe UI", 10))
        deliveryLabel.grid(row=2, column=0)


        nofeedbackLabel = tk.Label(labelFrame2, text="No Feedback", bg = "#9BC2E6", font=("Segoe UI", 10))
        nofeedbackLabel.grid(row=4, column=0)

        addcommtLabel = tk.Label(labelFrame2, text="Add Comment", bg = "#9BC2E6",font=("Segoe UI", 10))
        addcommtLabel.grid(row=4, column=1)
        # Configuring frame sizes based on screen size
        toproot.grid_rowconfigure(0, weight=1)
        toproot.grid_columnconfigure(0, weight=1)
        toproot.grid_rowconfigure(1, weight=1)
        toproot.grid_columnconfigure(1, weight=1)
        toproot.grid_rowconfigure(2, weight=1)

        labelFrame.grid_rowconfigure(0, weight=1)
        labelFrame.grid_columnconfigure(0, weight=1)
        labelFrame.grid_rowconfigure(1, weight=1)
        labelFrame.grid_columnconfigure(1, weight=1)
        labelFrame.grid_rowconfigure(2, weight=1)
        labelFrame.grid_columnconfigure(2, weight=1)
        labelFrame.grid_rowconfigure(3, weight=1)
        labelFrame.grid_columnconfigure(3, weight=1)

        entryFrame1.grid_rowconfigure(0, weight=1)
        entryFrame1.grid_columnconfigure(0, weight=1)
        entryFrame1.grid_rowconfigure(1, weight=1)
        entryFrame1.grid_columnconfigure(1, weight=1)
        entryFrame1.grid_rowconfigure(2, weight=1)
        entryFrame1.grid_columnconfigure(2, weight=1)
        entryFrame1.grid_rowconfigure(3, weight=1)
        entryFrame1.grid_columnconfigure(3, weight=1)

        entryFrame2.grid_rowconfigure(0, weight=1)
        entryFrame2.grid_columnconfigure(0, weight=1)
        entryFrame2.grid_rowconfigure(1, weight=1)
        entryFrame2.grid_columnconfigure(1, weight=1)

        entryFrame3.grid_rowconfigure(0, weight=1)
        entryFrame3.grid_columnconfigure(0, weight=1)
        entryFrame3.grid_rowconfigure(1, weight=1)
        entryFrame3.grid_columnconfigure(1, weight=1)

        submitFrame.grid_rowconfigure(0, weight=1)
        submitFrame.grid_columnconfigure(0, weight=1)
        submitFrame.grid_rowconfigure(1, weight=1)
        submitFrame.grid_columnconfigure(1, weight=1)
        
        #Declaring Entry Boxes
        #Manual Entry boxes
        
        
        
        # qn = tk.StringVar()
        priceVar = tk.StringVar()
        priceBox = ttk.Entry(entryFrame1, textvariable=priceVar, background = 'white',width = 20)
        priceBox.grid(row=3,column=0,sticky=tk.EW,padx=10,pady=10)
        priceVar.set("*")
        priceBox.bind("<Tab>",tabFuncPrice)
        
        quote_no_var.append(myCombobox(df,toproot,item_list=list(df['quoteno']),frame=entryFrame1,row=1,column=0,width=10,list_bd = 0,foreground='blue',
         background='white',sticky = tk.EW,cxDict= cxDatadict,val=priceBox))
        # Quotenum = ttk.Entry(entryFrame1, textvariable=qn, background = 'white',width = 20)
        # Quotenum.grid(row=1,column=0,sticky=tk.EW,padx=10,pady=10)
        # qn.set("*")

        preparedVar = tk.StringVar()
        preparedBox = ttk.Entry(entryFrame2, textvariable=preparedVar, background = 'white',width = 20)
        preparedBox.grid(row=1,column=0,sticky=tk.EW,padx=10,pady=10)
        preparedVar.set("*")
        cxDatadict["preparedby"].append((preparedBox, preparedVar))

        salespersonVar = tk.StringVar()
        salespersonBox = ttk.Entry(entryFrame3, textvariable=salespersonVar, background = 'white',width = 20)
        salespersonBox.grid(row=1,column=0,sticky=tk.EW,padx=10,pady=10)
        salespersonVar.set("*")
        cxDatadict["sales_person"].append((salespersonBox, salespersonVar))
        salespersonBox.bind("<Tab>",tabFunc)
        
        

        TimeVar = tk.StringVar()
        TimeBox = ttk.Entry(entryFrame2, textvariable=TimeVar, background = 'white',width = 20)
        TimeBox.grid(row=3,column=0,sticky=tk.EW,padx=10,pady=10)
        TimeVar.set("*")
        TimeBox.bind("<Tab>",tabFuncTime)

        DeliveryVar = tk.StringVar()
        DeliveryBox = ttk.Entry(entryFrame3, textvariable=DeliveryVar, background = 'white',width = 20)
        # DeliveryBox = tk.Text(entryFrame3, height = 5, width = 5,bg='white')
        # DeliveryBox.pack(side='bottom',padx=2)
        DeliveryBox.grid(row=3,column=0,sticky=tk.EW,padx=10,pady=10)
        DeliveryVar.set("*")

        no_feedbackVar = tk.StringVar()
        # no_feedbackBox = ttk.Entry(labelFrame, textvariable=no_feedbackVar, background = 'white',width = 10)
        no_feedbackBox = ttk.Entry(labelFrame2, textvariable= no_feedbackVar, width = 30,background='white')
        no_feedbackBox.grid(row=5,column=0,sticky=tk.EW,padx=10,pady=10)
        no_feedbackVar.set("*")

        add_commtVar = tk.StringVar()
        # add_commtBox = ttk.Entry(labelFrame, textvariable=add_commtVar, background = 'white',width = 10)
        add_commtBox = ttk.Entry(labelFrame2, textvariable= add_commtVar, width = 30,background='white')
        add_commtBox.grid(row=5,column=1,sticky=tk.EW,padx=10,pady=10)
        add_commtVar.set("*")

        data = {'QUOTE_NUMBER': [], 'PREPAREDBY': [], 'SALES_PERSON': [], 'PRICE': [], 'TIME': [], 'DELIVERY': [], 'NO_FEEDBACK': [], 'ADDITIONAL_COMMENT':''}
        df = pd.DataFrame(data)
        def generate():
            submitButton.configure(state='disable')
            priceBox.configure(state='disable')
            preparedBox.configure(state='disable')
            salespersonBox.configure(state='disable')
            TimeBox.configure(state='disable')
            DeliveryBox.configure(state='disable')
            no_feedbackBox.configure(state='disable')
            add_commtBox.configure(state='disable')

            # qn_list.append(qn.get())
            # preparedVar_list.append(preparedVar.get())
            # salespersonVar_list.append(salespersonVar.get())
            # priceVar_list.append(priceVar.get())
            # TimeVar_list.append(TimeVar.get())
            # DeliveryVar_list.append(DeliveryVar.get())
            # no_feedbackVar_list.append(no_feedbackVar.get())
            # add_commtVar_list.append(add_commtVar.get())

            value1=cxDatadict["quoteno"][0][0][0].get()
            value2=preparedVar.get()
            value3=salespersonVar.get()
            value4=priceVar.get()
            value5=TimeVar.get()
            value6=DeliveryVar.get()
            value7=no_feedbackVar.get()
            value8=add_commtVar.get()

            df.loc[len(df)] = [value1, value2, value3, value4, value5, value6,value7,value8]

        # df['QUOTE_NUMBER','PREPAREDBY','SALES_PERSON','PRICE','TIME','DELIVERY','NO_FEEDBACK','ADDITIONAL']= qn_list,preparedVar_list,salespersonVar_list,priceVar_list,TimeVar_list,DeliveryVar_list ,no_feedbackVar_list,add_commtVar_list
        
            print(df)

            upload_no_order_why(conn, df)
            messagebox.showinfo("Info", "Data Uploaded Successfully")
            submitButton.configure(state='normal')
            priceBox.configure(state='normal')
            preparedBox.configure(state='normal')
            salespersonBox.configure(state='normal')
            TimeBox.configure(state='normal')
            DeliveryBox.configure(state='normal')
            no_feedbackBox.configure(state='normal')
            add_commtBox.configure(state='normal')

        submitButton = ttk.Button(submitFrame,text="Submit",command=generate)
        # submitButton.place(relx=.5, rely=.5, anchor="center")
        submitButton.grid(row=0,column=1,pady=40)
        quote_no_var[0][0].focus_set()
        toproot.focus()
        # gradeBox.focus()

        # gradeVar.set('*')
        # yieldVar.set('*')
        # odVar.set('*')
        # idVar.set('*')

        def on_closing():
            try:
                toproot.grab_release()
                toproot.destroy()
            except Exception as e:
                raise e
        
        
        toproot.protocol("WM_DELETE_WINDOW", on_closing)
    except Exception as e:
        raise e


##########################################################################################################

def reportGenerator(root, conn):
    def intFloat(inStr,acttyp):
        try:
            # if acttyp == '1': #insert
            if inStr == '' or inStr == "NA":
                return True
            try:
                float(inStr)
                # print('value:', inStr)
            except ValueError:
                toproot.attributes('-topmost', True)
                messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in Integer or Decimal format only",parent=toproot)
                toproot.attributes('-topmost', False)
                return False
            return True
        except Exception as e:
            raise e
    try:
        global df
        df = get_master_df(conn,table = MASTER_TABLE)
        toproot = tk.Toplevel(root, bg = "#9BC2E6")
        toproot.title('EAGS Report Generator')
        screen_width = toproot.winfo_screenwidth()
        screen_height = toproot.winfo_screenheight()
        toproot.resizable(False,False)

        width = 415
        height = 380
        # calculate position x and y coordinates
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        toproot.geometry('%dx%d+%d+%d' % (width, height, x, y))

        toproot.grab_set()

        labelFrame = tk.Frame(toproot, bg= "#9BC2E6")
        labelFrame.grid(row=0, column=1)
        entryFrame1 = tk.Frame(labelFrame, bg= "#9BC2E6")
        entryFrame1.grid(row=1, column=0)
        entryFrame2 = tk.Frame(labelFrame, bg= "#9BC2E6")
        entryFrame2.grid(row=1, column=1)
        quoteFrame = tk.Frame(toproot, bg= "#9BC2E6")
        quoteFrame.grid(row=1, column=1)
        submitFrame = tk.Frame(toproot, bg= "#9BC2E6")
        submitFrame.grid(row=2, column=1)
        #Declaring Labels
        # tobePrinted = tk.Label(labelFrame, text="To be Printed", bg = "#9BC2E6")
        # tobePrinted.grid(row=0, column=0)

        # tobeCalculated = tk.Label(labelFrame, text="To be Calculated", bg = "#9BC2E6")
        # tobeCalculated.grid(row=0,column=1)

        userLabel = tk.Label(entryFrame1, text="Username", bg = "#9BC2E6", font=("Segoe UI", 10))
        userLabel.grid(row=0, column=0)

        loactionLabel = tk.Label(entryFrame1, text="Location", bg = "#9BC2E6", font=("Segoe UI", 10))
        loactionLabel.grid(row=0, column=1)

        gradeLabel = tk.Label(entryFrame2, text="Grade", bg = "#9BC2E6", font=("Segoe UI", 10))
        gradeLabel.grid(row=0, column=0)

        yieldLabel = tk.Label(entryFrame2, text="Yield", bg = "#9BC2E6", font=("Segoe UI", 10))
        yieldLabel.grid(row=0, column=1)

        fromOdLabel = tk.Label(entryFrame1, text="From OD", bg = "#9BC2E6", font=("Segoe UI", 10))
        fromOdLabel.grid(row=2,column=0)

        toOdLabel = tk.Label(entryFrame1, text="To OD", bg = "#9BC2E6", font=("Segoe UI", 10))
        toOdLabel.grid(row=2, column=1)

        fromIdLabel = tk.Label(entryFrame2, text="From ID", bg = "#9BC2E6", font=("Segoe UI", 10))
        fromIdLabel.grid(row=2, column=0)

        toIdLabel = tk.Label(entryFrame2, text="To ID", bg = "#9BC2E6", font=("Segoe UI", 10))
        toIdLabel.grid(row=2, column=1)

        fDateLabel = tk.Label(labelFrame, text="From Date", bg = "#9BC2E6", font=("Segoe UI", 10))
        fDateLabel.grid(row=2, column=0)

        tDateLabel = tk.Label(labelFrame, text="To Date", bg = "#9BC2E6")
        tDateLabel.grid(row=2, column=1)
        #Configuring frame sizes based on screen size
        toproot.grid_rowconfigure(0, weight=1)
        toproot.grid_columnconfigure(0, weight=1)
        toproot.grid_rowconfigure(1, weight=1)
        toproot.grid_columnconfigure(1, weight=1)
        toproot.grid_rowconfigure(2, weight=1)

        labelFrame.grid_rowconfigure(0, weight=1)
        labelFrame.grid_columnconfigure(0, weight=1)
        labelFrame.grid_rowconfigure(1, weight=1)
        labelFrame.grid_columnconfigure(1, weight=1)
        labelFrame.grid_rowconfigure(2, weight=1)
        labelFrame.grid_columnconfigure(2, weight=1)
        labelFrame.grid_rowconfigure(3, weight=1)
        labelFrame.grid_columnconfigure(3, weight=1)

        entryFrame1.grid_rowconfigure(0, weight=1)
        entryFrame1.grid_columnconfigure(0, weight=1)
        entryFrame1.grid_rowconfigure(1, weight=1)
        entryFrame1.grid_columnconfigure(1, weight=1)

        entryFrame2.grid_rowconfigure(0, weight=1)
        entryFrame2.grid_columnconfigure(0, weight=1)
        entryFrame2.grid_rowconfigure(1, weight=1)
        entryFrame2.grid_columnconfigure(1, weight=1)

        submitFrame.grid_rowconfigure(0, weight=1)
        submitFrame.grid_columnconfigure(0, weight=1)
        submitFrame.grid_rowconfigure(1, weight=1)
        submitFrame.grid_columnconfigure(1, weight=1)
        
        #Declaring Entry Boxes
        #Manual Entry boxes
        
        
        
        userVar = tk.StringVar()
        
        userBox = ttk.Entry(entryFrame1, textvariable=userVar, background = 'white',width = 10)
        userBox.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        userVar.set("*")

        locationVar = tk.StringVar()
        locationBox = ttk.Entry(entryFrame1, textvariable=locationVar, background = 'white',width = 10)
        locationBox.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        locationVar.set("*")

        gradeVar = tk.StringVar()
        
        gradeBox = ttk.Entry(entryFrame2, textvariable=gradeVar, background = 'white',width = 10)
        gradeBox.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        gradeVar.set("*")

        yieldVar = tk.StringVar()
        yieldBox = ttk.Entry(entryFrame2, textvariable=yieldVar, background = 'white',width = 10)
        yieldBox.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        yieldVar.set("*")

        vcmd = toproot.register(intFloat)
        fromOdVar = tk.StringVar()
        
        fromOd = ttk.Entry(entryFrame1, textvariable=fromOdVar, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        fromOd['validatecommand'] = (fromOd.register(intFloat),'%P','%d')
        fromOd.grid(row=3,column=0,sticky=tk.EW,padx=5,pady=5)
        # fromOdVar.set("*")

        toOdVar = tk.StringVar()
        toOd = ttk.Entry(entryFrame1, textvariable=toOdVar, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        toOd['validatecommand'] = (toOd.register(intFloat),'%P','%d')
        toOd.grid(row=3,column=1,sticky=tk.EW,padx=5,pady=5)
        # yieldVar.set("*")

        fromIdVar = tk.StringVar()
        
        fromId = ttk.Entry(entryFrame2, textvariable=fromIdVar, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        fromId['validatecommand'] = (fromId.register(intFloat),'%P','%d')

        fromId.grid(row=3,column=0,sticky=tk.EW,padx=5,pady=5)
        # odVar.set("*")

        toIdVar = tk.StringVar()
        toId = ttk.Entry(entryFrame2, textvariable=toIdVar, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        toId['validatecommand'] = (toId.register(intFloat),'%P','%d')

        toId.grid(row=3,column=1,sticky=tk.EW,padx=5,pady=5)
        
        # fDateVar = tk.StringVar()
        fDateBox = MyDateEntry(master=labelFrame, width=10, selectmode='day', font=("Segoe UI", 10)) #ttk.Entry(labelFrame, textvariable=fDateVar, background = 'white',width = 10)
        fDateBox.grid(row=3,column=0,sticky=tk.EW,padx=5,pady=5)
        # fDateVar.set("*")

        # tDateVar = tk.StringVar()
        tDateBox = MyDateEntry(master=labelFrame, width=10, selectmode='day', font=("Segoe UI", 10))#ttk.Entry(labelFrame, textvariable=tDateVar, background = 'white',width = 10)
        tDateBox.grid(row=3,column=1,sticky=tk.EW,padx=5,pady=5)
        # tDateVar.set("*")

        s = ttk.Style()
        s.configure("TRadiobutton", background=toproot["bg"])
        s.configure("TLabel", font=("Segoe UI", 10))

        quoteLabel = ttk.Label(quoteFrame, text="Quote Yes/No", background = toproot["bg"])#, font=("Segoe UI", 10))
        quoteLabel.grid(row=0, column=0,sticky="nw", padx=(0,30))
        yesVar = tk.IntVar()
        yesBox = ttk.Radiobutton(quoteFrame, text='Yes',variable=yesVar, value=1)#, onvalue=1, offvalue=0)#, font=("Segoe UI", 10))
        yesBox.grid(row=0, column=1, sticky="ne")
        yesVar.set(1)
        noVar = tk.IntVar()
        noBox = ttk.Radiobutton(quoteFrame, text='No',variable=yesVar, value=2)#, onvalue=1, offvalue=0)#, font=("Segoe UI", 10))
        noBox.grid(row=0, column=2, sticky="ne")
        # noVar.set(1)
        otherVar = tk.IntVar()
        otherBox = ttk.Radiobutton(quoteFrame, text='Other',variable=yesVar ,value=3)#, onvalue=1, offvalue=0)#, font=("Segoe UI", 10))
        otherBox.grid(row=0, column=3, sticky="ne")
        # otherVar.set(1)
        allVar = tk.IntVar()
        allBox = ttk.Radiobutton(quoteFrame, text='All',variable=yesVar, value=4)#, onvalue=1, offvalue=0)#, font=("Segoe UI", 10))
        allBox.grid(row=0, column=4, sticky="ne")
        # allVar.set(1)

        

        def starSearcher():
            """Filtering downloaded master dataframe one by one"""
            try:
                global df
                print(yesVar.get())
                

                user = userVar.get()
                locationValue = locationVar.get()
                grade = gradeVar.get()
                yieldValue = yieldVar.get()
                from_od_value = fromOdVar.get()
                to_od_value = toOdVar.get()
                from_id_value = fromIdVar.get()
                to_id_value = toIdVar.get()
                tDate = tDateBox.get()
                fDate = fDateBox.get()
                quoteYesNo = yesVar.get()
                


                filtered_df = df.copy()

                #Filtering based on Username
                if user == '' or user == '*':
                    filtered_df = df.copy()
                elif '*' in user:
                    filtered_df = filtered_df.loc[df["PREPAREDBY"].str.startswith((user.replace('*','')))]
                else:
                    filtered_df = df[(df['PREPAREDBY']==user)]

                noCheck= False
                allCheck = False
                # #Filtering Based on Quote Yes, No and Other
                if quoteYesNo==1:#Yes
                    filtered_df = filtered_df[(filtered_df["C_QUOTE_YES/NO"]=="Yes")]
                elif quoteYesNo==2:#No
                    filtered_df = filtered_df[(filtered_df["C_QUOTE_YES/NO"]=="No")]
                    noCheck = True
                elif quoteYesNo==3:#Other
                    filtered_df = filtered_df[(filtered_df["C_QUOTE_YES/NO"]=="Other")]
                elif quoteYesNo==4:#All
                    allCheck = True
                else:
                    toproot.attributes('-topmost', True)
                    messagebox.showerror("Error", f"Please check Quote Yes or No Checkboxes",parent=toproot)
                    toproot.attributes('-topmost', False)
                    return
                
                

                na_df = []

                #Filtering NA Values if required for No and All case
                if noCheck or allCheck:
                    na_df = filtered_df[
                            ("NA" == filtered_df['E_OD1'])  & (filtered_df['E_OD1'] == "NA")
                            ]
                    if allCheck:
                        filtered_df = filtered_df[
                            ("NA" != filtered_df['E_OD1'])  & (filtered_df['E_OD1'] != "NA")
                            ]
                if not noCheck:
                    #Filtering based on Location
                    if locationValue == '' or locationValue == '*':
                        pass
                    elif '*' in locationValue:
                        filtered_df = filtered_df.loc[df["E_LOCATION"].str.startswith(locationValue.replace('*',''))]
                    else:
                        filtered_df = filtered_df.loc[(df['E_LOCATION']==locationValue)]
                    #Filtering based on Grade
                    if grade == "*":
                            pass
                    elif "*" in grade:
                        filtered_df = filtered_df.loc[df["E_GRADE"].str.startswith(grade.replace('*',''))]
                    else: # yieldValue != "*":
                        filtered_df  = filtered_df[ (filtered_df["E_GRADE"]==grade)]

                    #Filtering based on Yield
                    if yieldValue == "*":
                            pass
                    elif "*" not in yieldValue:
                        filtered_df = filtered_df.loc[df["E_YIELD"].str.startswith(yieldValue.replace('*',''))]
                    else: # yieldValue != "*":
                        filtered_df  = filtered_df[ (filtered_df["E_YIELD"]==yieldValue)]

                    #Filtering OD
                    if from_od_value == "" and to_od_value == "":
                        pass
                    elif from_od_value != "" and to_od_value != "":
                        filtered_df = filtered_df[
                                (float(from_od_value) <= filtered_df['E_OD1'].astype(float))  & (filtered_df['E_OD1'].astype(float) <= float(to_od_value))
                                ]
                    else:
                        toproot.attributes('-topmost', True)
                        messagebox.showerror("Error", f"Please check OD search query and try again",parent=toproot)
                        toproot.attributes('-topmost', False)
                        return


                    #Filtering ID
                    if from_id_value == "" and to_id_value == "":
                        pass
                    elif from_id_value != "" and to_id_value != "":
                        filtered_df = filtered_df[
                                (float(from_id_value) <= filtered_df['E_ID1'].astype(float))  & (filtered_df['E_ID1'].astype(float) <= float(to_id_value))
                                ]
                    else:
                        toproot.attributes('-topmost', True)
                        messagebox.showerror("Error", f"Please check ID search query and try again",parent=toproot)
                        toproot.attributes('-topmost', False)
                        return

                if len(na_df):
                    filtered_df = pd.concat([filtered_df,na_df],ignore_index=True)

                if len(filtered_df):
                #    df['DATE'] = df['DATE'].dt.strftime('%d.%m.%Y')
                    filtered_df=  df[(df['DATE'] >= datetime.strptime(fDate,'%d.%m.%Y').date())]
                    filtered_df=  filtered_df[(filtered_df['DATE'] <= datetime.strptime(tDate,'%d.%m.%Y').date())]
                
                
                if len(filtered_df):
                    
                    # filtered_df = filtered_df[["site", "grade", "heat_condition", "od_in","od_in_2",'age' ,'date_last_receipt','onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in', 'heat_number', 'lot_serial_number']]
                    # filtered_df = filtered_df.sort_values(["site","grade", "heat_condition", "od_in","od_in_2", "age"], ascending=[True, True, True, True, True, False])
                    

                    colList = ['QUOTE NO', 'PREPARED BY', 'SALES_PERSON', 'DATE', 'CUSTOMER_NAME', 'PAYMENT_TERM', 'CURRENCY', 'CUSTOMER_ADDRESS', 'CUSTOMER_PHONE', 'CUSTOMER_EMAIL', 'CUSTOMER_CITY_ZIP', 'WORK_ORDER', 'DELIVERY_DATE',
                     'MATERIALNUMBER', 'MATERIALDESCRIPTION', 'QTY', 'RM', 'RMQTY', 'SAW_CUT', 'CUSTOMER_SPECIFICATION', 'CUSTOMER_TYPE', 'CUSTOMER_GRADE', 'CUSTOMER_YIELD', 'CUSTOMER_OD', 'CUSTOMER_ID', 'CUSTOMER_LENGTH', 
                     'CUSTOMER_QTY', 'CUSTOMER_QUOTE_YES/NO', 'E_LOCATION', 'E_TYPE', 'E_SPEC', 'E_GRADE', 'E_YIELD', 'E_OD1', 'E_ID1', 'E_OD2', 'E_ID2', 'E_LENGTH', 'E_QTY', 'E_COST', 
                     'E_SELLING_COST/LBS', 'E_MARGIN_LBS', 'E_UOM', 'E_SELLING_COST/UOM', 'E_ADDITIONAL_COST', 'LEAD_TIME', 'E_FINAL_PRICE', 'E_FREIGHT_INCURED', 
                     'E_FREIGHT_CHARGED', 'E_MARGIN_FREIGHT', 'LOT_SERIAL_NUMBER', 'VALIDITY', 'ADD_COMMENTS', 'PREVIOUS_QUOTE', 'REV_CHECKER']
                    filtered_df.drop('INSERT_DATE', axis=1, inplace=True)
                    filtered_df.columns = colList

                    desktopDir = os.path.join(os.environ["HOMEPATH"], "Desktop\\EAGS_Reports")
                    desktopDir = os.path.join('C:', desktopDir)
                    if not os.path.exists(desktopDir):
                        os.mkdir(desktopDir)
                    #Report name format Report_today_date[1],
                    reportName = f'Report_{datetime.strftime(datetime.today(), "%m%d%Y_%H%M%S")}.xlsx'
                    reportPath = desktopDir+ "\\" +reportName
                    filtered_df.to_excel(reportPath, index=False)
                    toproot.attributes('-topmost', True)
                    messagebox.showinfo("Successfull", f"{reportName} has been generated and parked in your Desktop EAGS_Reports Folder",parent=toproot)
                    toproot.attributes('-topmost', False)
                else:
                    toproot.attributes('-topmost', True)
                    messagebox.showerror("Error", f"Please check search query and try again",parent=toproot)
                    toproot.attributes('-topmost', False)
                    return
            except Exception as e:
                raise e
                
               
            
        submitButton = ttk.Button(submitFrame,text="Submit", command=starSearcher)
        # submitButton.place(relx=.5, rely=.5, anchor="center")
        submitButton.grid(row=0,column=1,pady=40)
        toproot.focus()
        gradeBox.focus()

        gradeVar.set('*')
        yieldVar.set('*')
        # odVar.set('*')
        # idVar.set('*')

        def on_closing():
            try:
                toproot.grab_release()
                toproot.destroy()
            except Exception as e:
                raise e
        
        
        toproot.protocol("WM_DELETE_WINDOW", on_closing)
    except Exception as e:
        raise e



# conn = get_connection()
# # conn=None
# root = tk.Tk()
# user = "Imam"
# df = pd.read_excel("sampleInventory.xlsx")

# reportGenerator(root)
# root.mainloop()








