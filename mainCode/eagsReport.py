import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import date
from sfTool import get_connection, get_inv_df, get_master_df
from datetime import datetime
import os
import pandas as pd
# import ChecklistCombobox

MASTER_TABLE = "EAGS_MASTER"


#Calendar
class MyDateEntry(DateEntry):
    try:
        def __init__(self, master=None, **kw):
            DateEntry.__init__(self, master=master, date_pattern='mm.dd.yyyy',**kw)
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

        width = 415
        height = 380
        # calculate position x and y coordinates
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        toproot.geometry('%dx%d+%d+%d' % (width, height, x, y))

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
                    root.attributes('-topmost', True)
                    messagebox.showerror("Error", f"Please check Quote Yes or No Checkboxes",parent=root)
                    root.attributes('-topmost', False)
                    return
                
                #Filtering based on Location
                if locationValue == '' or locationValue == '*':
                    pass
                elif '*' in locationValue:
                    filtered_df = filtered_df.loc[df["E_LOCATION"].str.startswith(locationValue.replace('*',''))]
                else:
                    filtered_df = df[(df['E_LOCATION']==locationValue)]


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
                        root.attributes('-topmost', True)
                        messagebox.showerror("Error", f"Please check OD search query and try again",parent=root)
                        root.attributes('-topmost', False)
                        return


                    #Filtering ID
                    if from_id_value == "" and to_id_value == "":
                        pass
                    elif from_id_value != "" and to_id_value != "":
                        filtered_df = filtered_df[
                                (float(from_id_value) <= filtered_df['E_ID1'].astype(float))  & (filtered_df['E_ID1'].astype(float) <= float(to_id_value))
                                ]
                    else:
                        root.attributes('-topmost', True)
                        messagebox.showerror("Error", f"Please check ID search query and try again",parent=root)
                        root.attributes('-topmost', False)
                        return

                if len(na_df):
                    filtered_df = pd.concat([filtered_df,na_df],ignore_index=True)


                
                
                if len(filtered_df):
                    
                    # filtered_df = filtered_df[["site", "grade", "heat_condition", "od_in","od_in_2",'age' ,'date_last_receipt','onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in', 'heat_number', 'lot_serial_number']]
                    # filtered_df = filtered_df.sort_values(["site","grade", "heat_condition", "od_in","od_in_2", "age"], ascending=[True, True, True, True, True, False])
                    

                    colList = ['QUOTE NO', 'PREPARED BY', 'DATE', 'CUSTOMER_NAME', 'PAYMENT_TERM', 'CURRENCY', 'CUSTOMER_ADDRESS', 'CUSTOMER_PHONE', 'CUSTOMER_EMAIL', 'CUSTOMER_CITY_ZIP', 'WORK_ORDER', 'DELIVERY_DATE',
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
                    root.attributes('-topmost', True)
                    messagebox.showinfo("Successfull", f"{reportName} has been generated and parked in your Desktop EAGS_Reports Folder",parent=root)
                    root.attributes('-topmost', False)
                else:
                    root.attributes('-topmost', True)
                    messagebox.showerror("Error", f"Please check search query and try again",parent=root)
                    root.attributes('-topmost', False)
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
    except Exception as e:
        raise e



# conn = get_connection()
# # conn=None
# root = tk.Tk()
# user = "Imam"
# df = pd.read_excel("sampleInventory.xlsx")

# reportGenerator(root)
# root.mainloop()








