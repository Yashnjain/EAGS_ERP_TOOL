import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from customComboboxV2 import myCombobox
from tkcalendar import DateEntry
from datetime import date
import sys
import pandas as pd
from pandastable import Table
from Tools import dfMaker, resource_path
from sfTool import get_connection,get_cx_df, get_inv_df
from final_pdf_creator import pdf_generator
from sfTool import eagsQuotationuploader
import os

UNITS = "units"
INV_TABLE = "EAGS_INVENTORY"
CX_TABLE = "EAGS_CUSTOMER"

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


class ResizingCanvas(tk.Canvas):
    try:
        def __init__(self,parent,**kwargs):
            tk.Canvas.__init__(self,parent,**kwargs)
            self.bind("<Configure>", self.on_resize)
            self.height = self.winfo_reqheight()
            self.width = self.winfo_reqwidth()

        def on_resize(self,event):
            # determine the ratio of old width/height to new width/height
            wscale = float(event.width)/self.width
            hscale = float(event.height)/self.height
            self.width = event.width
            self.height = event.height
            # resize the canvas 
            self.config(width=self.width, height=self.height)
            # rescale all the objects tagged with the "all" tag
            self.scale("all",0,0,wscale,hscale)
    except Exception as e:
        raise e





def quoteGenerator(mainRoot,user,conn):
    try:
        def set_mousewheel(widget, command):
            try:
                """Activate / deactivate mousewheel scrolling when 
                cursor is over / not over the widget respectively."""
                widget.bind("<Enter>", lambda _: widget.bind_all('<MouseWheel>', command))
                widget.bind("<Leave>", lambda _: widget.unbind_all('<MouseWheel>'))
            except Exception as e:
                raise e
        def OnMouseWheel(event):
            try:
                # print(yscrollbar.get())
                if yscrollbar.get() != (0.0, 1.0):
                # if yscrollbar.get()[-1]!=1.0 or yscrollbar.get()[0]!=0.0:
                    entryCanvas.yview_scroll(int(-1*(event.delta/120)), "units")
                # show bottom of canvas
                # entryCanvas.yview_moveto('1.0')
            except Exception as e:
                raise e

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
 
        def intFloat(inStr,acttyp):
            try:
                # if acttyp == '1': #insert
                if inStr == '' or inStr == "NA":
                    return True
                try:
                    float(inStr)
                    # print('value:', inStr)
                except ValueError:
                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value",parent=entryFrame)
                    return False
                return True
            except Exception as e:
                raise e

        def intChecker(inStr,acttyp):
            try:
                # if acttyp == '1': #insert
                if inStr == '' or inStr == "NA":
                    return True
                if not inStr.isdigit():
                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value",parent=entryFrame)
                    return False
                return True
            except Exception as e:
                raise e
        def on_configure(event):
            try:
                # update scrollregion after starting 'mainloop'
                # when all widgets are in canvas
                entryCanvas.configure(scrollregion=entryCanvas.bbox('all'))#,width=1890,height=380)#(0,0,300,200)
            except Exception as e:
                raise e
        def returnTohome():
            try:
                root.withdraw()
                mainRoot.deiconify()
                mainRoot.state('zoomed')
            except Exception as e:
                raise e

        def tabFunc(e):
            try:
                cx_spec[0][0].focus_set()
                return "break"
            
            except Exception as e:
                raise e

        def addRow():
            try:
                row_num = len(quoteYesNo)
                
                cx_spec.append((ttk.Entry(entryFrame,width=15),None))
                cx_spec[-1][0].grid(row=2+row_num,column=0,padx=(15,0))

                
                # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                cx_grade.append((ttk.Entry(entryFrame,width=15),None))
                cx_grade[-1][0].grid(row=2+row_num,column=1)
                # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                cx_yield.append((ttk.Entry(entryFrame,width=15),None))
                cx_yield[-1][0].grid(row=2+row_num,column=2)
                # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=3,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                
                vcmd = tab1.register(intFloat)
                cx_od.append((ttk.Entry(entryFrame, width=5,validate = "key",
                        validatecommand=(vcmd, '%P','%d')),None))
                cx_od[-1][0].grid(row=2+row_num,column=3)
                # cx_od['validatecommand'] = (cx_od.register(intFloat),'%P','%d')



                # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=4,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                cx_id.append((ttk.Entry(entryFrame, width=5, validate = "key"),None))
                cx_id[-1][0].grid(row=2+row_num,column=4)
                cx_id[-1][0]['validatecommand'] = (cx_id[-1][0].register(intFloat),'%P','%d')
                # myCombobox(df,tab1,cx_list,frame=entryFrame,row=1,column=5,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                cx_len.append((ttk.Entry(entryFrame, width=5, validate = "key"),None))
                cx_len[-1][0].grid(row=2+row_num,column=5)
                cx_len[-1][0]['validatecommand'] = (cx_len[-1][0].register(intFloat),'%P','%d')
                # myCombobox(df,tab1,cx_list,frame=entryFrame,row=1,column=6,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                cx_qty.append((ttk.Entry(entryFrame, width=5, validate = "key"), None))
                cx_qty[-1][0].grid(row=2+row_num,column=6)
                cx_qty[-1][0]['validatecommand'] = (cx_qty[-1][0].register(intChecker),'%P','%d')
                quoteYesNo.append(myCombobox(df,tab1,["Yes","No","Other"],frame=entryFrame,row=2+row_num,column=7,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                # quoteYesNo[-1]['validate']='focusout'
                # quoteYesNo[-1]['validatecommand'] = (quoteYesNo[-1].register(yesNo),'%P','%W')
                e_location.append(myCombobox(df,tab1,["Dubai","Singapore","USA","UK"],frame=entryFrame,row=2+row_num,column=8,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                # e_location[-1].config(textvariable="NA", state='disabled')
                e_type.append(myCombobox(df,tab1,["THF","BR"],frame=entryFrame,row=2+row_num,column=9,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # e_type[-1].config(textvariable="NA", state='disabled')
                e_grade.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=10,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # e_grade[-1].config(textvariable="NA", state='disabled')
                e_yield.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=11,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # e_yield[-1].config(textvariable="NA", state='disabled')
                e_od.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=12,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # e_od[-1].config(textvariable="NA", state='disabled')
                e_od[-1][0]['validate']='key'
                e_od[-1][0]['validatecommand'] = (e_od[-1][0].register(intFloat),'%P','%d')
                e_id.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=13,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                
                # e_id[-1].config(textvariable="NA", state='disabled')
                e_len.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=14,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                e_len[-1][0]['validate']='key'
                e_len[-1][0]['validatecommand'] = (e_len[-1][0].register(intFloat),'%P','%d')
                # e_len[-1].config(textvariable="NA", state='disabled')
                e_qty.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=15,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                e_qty[-1][0]['validate']='key'
                e_qty[-1][0]['validatecommand'] = (e_qty[-1][0].register(intChecker),'%P','%d')
                # e_qty[-1].config(textvariable="NA", state='disabled')
                sellCostLBS.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=16,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                sellCostLBS[-1][0]['validate']='key'
                sellCostLBS[-1][0]['validatecommand'] = (sellCostLBS[-1][0].register(intFloat),'%P','%d')
                # sellCostLBS[-1].config(textvariable="NA", state='disabled')
                e_uom.append(myCombobox(df,tab1,["Inch","Each"],frame=entryFrame,row=2+row_num,column=17,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                # e_uom[-1].config(textvariable="NA", state='disabled')
                sellCostUOM.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=18,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                sellCostUOM[-1][0]['validate']='key'
                sellCostUOM[-1][0]['validatecommand'] = (sellCostUOM[-1][0].register(intFloat),'%P','%d')
                # sellCostUOM[-1].config(textvariable="NA", state='disabled')
                addCost.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=19,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                addCost[-1][0]['validate']='key'
                addCost[-1][0]['validatecommand'] = (addCost[-1][0].register(intFloat),'%P','%d')
                # addCost[-1].config(textvariable="NA", state='disabled')
                leadTime.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=20,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                # leadTime[-1].config(textvariable="NA", state='disabled')
                finalCost.append(myCombobox(df,tab1,item_list,frame=entryFrame,row=2+row_num,column=21,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                finalCost[-1][0]['validate']='key'
                finalCost[-1][0]['validatecommand'] = (finalCost[-1][0].register(intFloat),'%P','%d')
            except Exception as e:
                raise e

        def cxListCalc():
            try:
                # finalCost[-1].config(textvariable="NA", state='disabled')
                cxList = [cxDatadict["Prepared_By"],cxDatadict["Date"],cxDatadict["cus_long_name"][0][0][0].get(), cxDatadict["payment_term"][0][0].get(), cxDatadict["cus_address"][0][0].get(),
                    cxDatadict["cus_phone"],cxDatadict["cus_email"][0][0].get(),cxDatadict["cus_city_zip"]]
                return cxList
            except Exception as e:
                raise e
        def otherListCalc():
            try:
                otherList = [validityVar.get(), addCommVar.get()]
                return otherList
                # row_num+=1
            except Exception as e:
                raise e
        def deleteRow():
            try:
                global quoteDf
                quoteDf = []
                submitButton.configure(state='disable')
                for key in specialList.keys():
                    
                    # specialList[key][0][-1][1].destroy()
                    if len(specialList[key][0])==1:
                        specialList[key][0][0][0].delete(0, tk.END)
                        
                        # time.sleep(1)
                        # addRow()
                    else:
                        specialList[key][0][-1][0].destroy()
                        specialList[key][0].pop()
                # show bottom of canvas
                entryCanvas.yview_moveto('1.0')
            except Exception as e:
                raise e
        
        def create_pdf():
            try:
                global quoteDf
                quoteDf = dfMaker(specialList,cxListCalc(),otherListCalc(),pt,conn)
                if len(quoteDf):
                    pt.model.df = quoteDf
                    pt.redraw()
                    
                    global pdf_path
                    pdf_path = pdf_generator(quoteDf)
                    submitButton.configure(state='normal')
                else:
                    messagebox.showerror("Error", "Empty dataframe was given in input")
            except Exception as e:
                raise e

        def uploadDf(conn, quoteDf):
            try:
                # pt.model.df = quoteDf
                # pt.redraw()
                if messagebox.askyesno("Upload to Database", "Are sure that you want to generate quote and upload Data?"):
                    eagsQuotationuploader(conn, quoteDf)
                    
                    messagebox.showinfo("Info", "Data uploaded Successfully!")

                    current_work_dir = os.getcwd()
                    cx_init_name = str(quoteDf['QUOTENO'][0]).split("_")[-1]
                    filename = str(quoteDf['QUOTENO'][0])+".pdf"
                    save_dir = current_work_dir+"\\"+cx_init_name
                    if not os.path.exists(save_dir):
                        os.mkdir(save_dir)
                    os.rename(pdf_path,save_dir+"\\"+filename)
                    
                else:
                    os.remove(pdf_path)
                    submitButton.configure(state='disable')
            except Exception as e:
                raise e

        mainRoot.withdraw()
        global row_num
        row_num=0

        #Getting invoentory dataframe
        df = get_inv_df(conn,table = INV_TABLE)
        
        
        # df = pd.read_excel("sampleInventory.xlsx")
        #Getting Cx Dataframe
        cx_df = get_cx_df(conn,table = CX_TABLE)
        # cx_df = pd.read_excel("cxDatabase.xlsx")

        count = 0
        root = tk.Toplevel(mainRoot, bg = "#9BC2E6")
        root.state('zoomed')
        tabControl = ttk.Notebook(root)
        s = ttk.Style(tabControl)
        s.configure("TFrame", background=root["bg"])
        tab1 = ttk.Frame(tabControl)
        tab2 = ttk.Frame(tabControl)
        tab3 = ttk.Frame(tabControl)

        tabControl.add(tab1, text='Quote Generator')
        tabControl.add(tab2, text='Machining')
        tabControl.add(tab3, text='Quote Generator + Machining')

        tabControl.pack(expand=1, fill='both')
        cxFrame = tk.Frame(tab1, bg = "#9BC2E6")#,highlightbackground="blue", highlightthickness=2)
        cxFrame2 = tk.Frame(tab1, bg = "#9BC2E6")#,highlightbackground="blue", highlightthickness=2)

        m_entryFrame = tk.Frame(tab1, bg= "#DDEBF7",highlightbackground="black", highlightthickness=2)#width=1700,height=300
        entryCanvas = tk.Canvas(m_entryFrame, bg= "#DDEBF7")#,width=1930,height=400)
        xscrollbar=ttk.Scrollbar(m_entryFrame,orient=tk.HORIZONTAL, command=entryCanvas.xview)
        
        entryCanvas.config(xscrollcommand = xscrollbar.set)

        #Defining frame inside canvas
        entryFrame = tk.Frame(entryCanvas, bg= "#DDEBF7")
        
        entryFrame.bind('<Configure>', on_configure)

        yscrollbar=ttk.Scrollbar(m_entryFrame,orient="vertical", command=entryCanvas.yview)
        set_mousewheel(entryCanvas, OnMouseWheel)
        # entryCanvas.bind_all("<MouseWheel>", OnMouseWheel)
        
        entryCanvas.config(yscrollcommand = yscrollbar.set)
        databaseFrame = tk.Frame(tab1,height=500, bg= "#DDEBF7")

        controlFrame = tk.Frame(tab1, bg= "#DDEBF7")


        cxFrame.grid(row=0, column=0,pady=(24,0), padx=(30,0),sticky="new")
        cxFrame2.grid(row=0, column=1,pady=(24,0), padx=(30,120),sticky="new")
        # headerFrame.grid(row=1, column=0,sticky="sew",columnspan=2)
        m_entryFrame.grid(row=1, column=0,sticky="nsew", columnspan=2)
        xscrollbar.grid(row=1,column=0,sticky=tk.NSEW)
        yscrollbar.grid(row=0,column=1,sticky=tk.NSEW)
        entryCanvas.grid(row=0,column=0, sticky=tk.NSEW)
        databaseFrame.grid(row=2,column=0, sticky="nsew")
        controlFrame.grid(row=2,column=1, sticky="nsew")
        
        entryCanvas.create_window((0,0),window=entryFrame,tags='expand')
        
        # entryFrame.grid(row=0,column=0)
        tab1.grid_rowconfigure(0, weight=1) # For row 0
        tab1.grid_rowconfigure(1, weight=1) # For row 1
        tab1.grid_rowconfigure(2, weight=1) # For row 1

        tab1.grid_columnconfigure(0, weight=1) # For column 0
        tab1.grid_columnconfigure(1, weight=1) # For column 1


        cxFrame.grid_rowconfigure(0, weight=1) # For row 0
        cxFrame.grid_rowconfigure(1, weight=1) # For row 1
        cxFrame.grid_rowconfigure(2, weight=1) # For row 2
        cxFrame.grid_rowconfigure(3, weight=1) # For row 3
        cxFrame.grid_rowconfigure(4, weight=1) # For row 4

        cxFrame.grid_columnconfigure(0, weight=1) # For column 0
        cxFrame.grid_columnconfigure(1, weight=1) # For column 1
        cxFrame.grid_columnconfigure(2, weight=1) # For column 2
        cxFrame.grid_columnconfigure(3, weight=1) # For column 3
        
        cxFrame2.grid_rowconfigure(0, weight=1) # For row 0
        cxFrame2.grid_rowconfigure(1, weight=1) # For row 1
        cxFrame2.grid_rowconfigure(2, weight=1) # For row 1

        cxFrame2.grid_columnconfigure(0, weight=1) # For column 0
        cxFrame2.grid_columnconfigure(1, weight=1) # For column 1
        cxFrame2.grid_columnconfigure(2, weight=1) # For column 1


        m_entryFrame.grid_rowconfigure(0, weight=1) # For row 0
        # m_entryFrame.grid_rowconfigure(1, weight=1) # For row 1

        m_entryFrame.grid_columnconfigure(0, weight=1) # For column 0
        # m_entryFrame.grid_columnconfigure(1, weight=1) # For column 1

        entryFrame.grid_rowconfigure(0, weight=1) # For row 0
        entryFrame.grid_rowconfigure(1, weight=1) # For row 1
        entryFrame.grid_rowconfigure(2, weight=1) # For row 1

        entryFrame.grid_columnconfigure(0, weight=1) # For column 0
        entryFrame.grid_columnconfigure(1, weight=1) # For column 1
        entryFrame.grid_columnconfigure(2, weight=1) # For column 2
        entryFrame.grid_columnconfigure(3, weight=1) # For column 3
        entryFrame.grid_columnconfigure(4, weight=1) # For column 4
        entryFrame.grid_columnconfigure(5, weight=1) # For column 5
        entryFrame.grid_columnconfigure(6, weight=1) # For column 6
        entryFrame.grid_columnconfigure(7, weight=1) # For column 7
        entryFrame.grid_columnconfigure(8, weight=1) # For column 8
        entryFrame.grid_columnconfigure(9, weight=1) # For column 9
        entryFrame.grid_columnconfigure(10, weight=1) # For column 10
        entryFrame.grid_columnconfigure(11, weight=1) # For column 11
        entryFrame.grid_columnconfigure(12, weight=1) # For column 12
        entryFrame.grid_columnconfigure(13, weight=1) # For column 13
        entryFrame.grid_columnconfigure(14, weight=1) # For column 14
        entryFrame.grid_columnconfigure(15, weight=1) # For column 15
        entryFrame.grid_columnconfigure(16, weight=1) # For column 16
        entryFrame.grid_columnconfigure(17, weight=1) # For column 17
        entryFrame.grid_columnconfigure(18, weight=1) # For column 18
        entryFrame.grid_columnconfigure(19, weight=1) # For column 19
        entryFrame.grid_columnconfigure(20, weight=1) # For column 20
        entryFrame.grid_columnconfigure(21, weight=1) # For column 21
        
        controlFrame.grid_rowconfigure(0, weight=1) # For column 21
        controlFrame.grid_rowconfigure(1, weight=1) # For column 21

        
        controlFrame.grid_rowconfigure(2, weight=1) # For column 21
    
        
        controlFrame.grid_rowconfigure(3, weight=1) # For column 21
        controlFrame.grid_columnconfigure(1, weight=1) # For column 21

        databaseFrame.grid_rowconfigure(1, weight=1) # For column 21
        databaseFrame.grid_columnconfigure(1, weight=1) # For column 21
      ###############Importing Images fir buttons #####################################
        home_path = resource_path("home.png")
        add_img_path = resource_path("addRowS.png")
        add_img2_path = resource_path("addRow2S.png")
        delete_img_path = resource_path("deleteRowS.png")
        delete_img2_path = resource_path("deleteRow2S.png")
        preview_img_path = resource_path("previewButtonS.png")
        preview_img2_path = resource_path("previewButton2S.png")
        submit_img_path = resource_path("submitButtonS.png")
        submit_img2_path = resource_path("submitButton2S.png")
        home_img = tk.PhotoImage(master=tab1, file=home_path)
        add_img = tk.PhotoImage(master=controlFrame, file=add_img_path)
        add_img2 = tk.PhotoImage(master=controlFrame, file=add_img2_path)
        delete_img = tk.PhotoImage(master=controlFrame, file=delete_img_path)
        delete_img2 = tk.PhotoImage(master=controlFrame, file=delete_img2_path)
        preview_img = tk.PhotoImage(master=controlFrame, file=preview_img_path)
        preview_img2 = tk.PhotoImage(master=controlFrame, file=preview_img2_path)
        submit_img = tk.PhotoImage(master=controlFrame, file=submit_img_path)
        submit_img2 = tk.PhotoImage(master=controlFrame, file=submit_img2_path)
        

        #Creating list to be sent fro df creation 
        #df = pd.read_clipboard(sep=',',on_bad_lines='skip')
        nonList = [[None,None,None,None,None,None]]
        pandasDf = pd.DataFrame(nonList,columns=['onhand_pieces', 'onhand_length_in', 'reserved_pieces', 'reserved_length_in', 'available_pieces', 'available_length_in'])
        # pandasDf = pd.DataFrame(cx_df)
        pt = Table(databaseFrame, dataframe=pandasDf,showtoolbar=True, showstatusbar=True, maxcellwidth=1500)
        pt.cellwidth=135
        pt.show()

        global cxDatadict
        cxDatadict = {}
        #Cx data Varilables

        cxDatadict["Prepared_By"] = []
        prepByVar = []

        cxDatadict["Date"] = []
        inpDateVar = []
        

        cxDatadict["cus_long_name"] = []
        cxNameVar = []
        cxDatadict["cus_long_name"].append(cxNameVar)
        
        cxDatadict["payment_term"] = []

        cxDatadict["cus_address"] = []

        cxDatadict["cus_phone"] = []

        cxDatadict["cus_city_zip"] = []

        cxDatadict["cus_email"] = []


        item_list = ('A4140', 'A4140M', 'A4330V', 'A4715', 'BS708M40', 'A4145M', '4542','4462')

        cxLabel = tk.Label(cxFrame, text="Customer Details", bg = "#9BC2E6")
        lb1 = tk.Label(cxFrame,text="Prepared By", bg = "#9BC2E6")
        lb2 = tk.Label(cxFrame,text="Date", bg = "#9BC2E6")
        lb3 = tk.Label(cxFrame,text="Customer Name", bg = "#9BC2E6")
        lb4 = tk.Label(cxFrame,text="Location/Address", bg = "#9BC2E6")
        lb5 = tk.Label(cxFrame,text="Email", bg = "#9BC2E6")
        lb6 = tk.Label(cxFrame,text="Payment Terms", bg = "#9BC2E6")
        blanckLabel = tk.Label(cxFrame2,text="", bg = "#9BC2E6")
        lb7 = tk.Label(cxFrame2,text="Validity", bg = "#9BC2E6")
        lb8 = tk.Label(cxFrame2,text="Additional Comments", bg = "#9BC2E6")
        cxLabel.grid(row=0,column=0)
        lb1.grid(row=1,column=0)
        lb2.grid(row=1,column=1)
        lb3.grid(row=3,column=0)
        lb4.grid(row=3,column=1)
        lb5.grid(row=3,column=2)
        lb6.grid(row=3,column=3)
        blanckLabel.grid(row=0,column=0)
        lb7.grid(row=1,column=0,padx=(100,5))
        lb8.grid(row=1,column=1)

        
        prep_by = ttk.Entry(cxFrame)
        prep_by.insert(tk.END, user)
        prep_by.grid(row=2,column=0)
        prep_by.config(state= "disabled")
        cxDatadict["Prepared_By"] = prep_by.get()

        
        inpDate = MyDateEntry(master=cxFrame, width=17, selectmode='day')
        inpDate.grid(row=2, column=1)
        cxDatadict["Date"] = inpDate.get()
        

        
        #Validity
        validityVar = tk.StringVar()
        validity = ttk.Entry(cxFrame2, textvariable=validityVar, foreground='blue', background = 'white',width = 15)
        validity.grid(row=2,column=0,padx=(100,5),pady=5)
        
        #Additional Comments
        addCommVar = tk.StringVar()
        addComm = ttk.Entry(cxFrame2, textvariable=addCommVar, foreground='blue', background = 'white',width = 15)
        addComm.grid(row=2,column=1,sticky=tk.EW,padx=5,pady=5)

        addComm.bind("<Tab>",tabFunc)
        
        


        #Customer Name Entry Box
        cxNameVar.append(myCombobox(cx_df,tab1,item_list,frame=cxFrame,row=4,column=0,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",cxDict= cxDatadict,val=validity))
        #location Address entry box
        locAddVar = tk.StringVar()
        locAdd = ttk.Entry(cxFrame, textvariable=locAddVar, foreground='blue', background = 'white',width = 20)
        locAdd.grid(row=4,column=1,sticky=tk.EW,padx=5,pady=5)
        # cxLocVar = []
        cxDatadict["cus_address"].append((locAdd, locAddVar))
        # myCombobox(df,root,item_list,frame=cxFrame,row=4,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")

        #Email
        emailAddVar = tk.StringVar()
        emailAdd = ttk.Entry(cxFrame, textvariable=emailAddVar, foreground='blue', background = 'white',width = 20)
        emailAdd.grid(row=4,column=2,sticky=tk.EW,padx=5,pady=5)
        # cxemailAddVar = []
        cxDatadict["cus_email"].append((emailAdd, emailAddVar))
        # myCombobox(df,root,item_list,frame=cxFrame,row=4,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")

        #Payment Terms Entry
        payTermVar = tk.StringVar()
        payTerm = ttk.Entry(cxFrame, textvariable=payTermVar, foreground='blue', background = 'white',width = 20)
        payTerm.grid(row=4,column=3,sticky=tk.EW,padx=5,pady=5)
        cxDatadict["payment_term"].append((payTerm, payTermVar))
        

        home_button = tk.Button(cxFrame2, image=home_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"],command=returnTohome)
        home_button.image = home_img #Preventing image to go into garbage
        home_button.grid(row=0,column=2,sticky="ne")
        # home_button.place(x=1600,y=-10,relx=0.1,rely=0.1,anchor="sw")
        #######################################
        
        ########################################
        
        

        #################Entry Form Section##############################################
        ######################defining labels############################################
        specLabel = tk.Label(entryFrame, text="Specification", bg= "#DDEBF7")
        gradeLabel = tk.Label(entryFrame, text="Grade", bg= "#DDEBF7")
        yieldLabel = tk.Label(entryFrame, text="Yield", bg= "#DDEBF7")
        odLabel = tk.Label(entryFrame, text="OD", bg= "#DDEBF7")
        idLabel = 	tk.Label(entryFrame, text="ID", bg= "#DDEBF7")
        lengthLabel = tk.Label(entryFrame, text="Length", bg= "#DDEBF7")
        qtyLabel = tk.Label(entryFrame, text="Qty", bg= "#DDEBF7")
        quoteLabel1 = tk.Label(entryFrame, text="Quote", bg= "#DDEBF7")
        quoteLabel2 = tk.Label(entryFrame, text="Yes/No", bg= "#DDEBF7")
        locationLabel = tk.Label(entryFrame, text="Location", bg= "#DDEBF7")
        typeLabel = tk.Label(entryFrame, text="Type", bg= "#DDEBF7")
        e_gradeLabel = tk.Label(entryFrame, text="Grade", bg= "#DDEBF7")
        e_yieldLabel = tk.Label(entryFrame, text="Yield", bg= "#DDEBF7")
        e_odLabel = tk.Label(entryFrame, text="OD", bg= "#DDEBF7")
        e_idLabel = tk.Label(entryFrame, text="ID", bg= "#DDEBF7")
        e_Length = tk.Label(entryFrame, text="Length", bg= "#DDEBF7")
        e_Qty = tk.Label(entryFrame, text="Qty", bg= "#DDEBF7")
        sellcostLbsLabel1 = tk.Label(entryFrame, text="Selling", bg= "#DDEBF7")
        sellcostLbsLabel2 = tk.Label(entryFrame, text="Cost/LBS", bg= "#DDEBF7")
        uom = tk.Label(entryFrame, text="UOM", bg= "#DDEBF7")
        sellcostUOMLabel1 = tk.Label(entryFrame, text="Selling", bg= "#DDEBF7")
        sellcostUOMLabel2 = tk.Label(entryFrame, text="Cost/UOM", bg= "#DDEBF7")
        addCostLabel1 = tk.Label(entryFrame, text="Additional", bg= "#DDEBF7")
        addCostLabel2 = tk.Label(entryFrame, text="Cost", bg= "#DDEBF7")
        leadTimeLAbel = tk.Label(entryFrame, text="Lead Time", bg= "#DDEBF7")
        finalPriceLabel = tk.Label(entryFrame, text="Final Price", bg= "#DDEBF7")



        specLabel.grid(row=0,column=0,padx=(15,0), sticky="ew")
        gradeLabel.grid(row=0,column=1, sticky="ew")
        yieldLabel.grid(row=0,column=2, sticky="ew")
        odLabel.grid(row=0,column=3, sticky="ew")
        idLabel.grid(row=0,column=4, sticky="ew")
        lengthLabel.grid(row=0,column=5, sticky="ew")
        qtyLabel.grid(row=0,column=6, sticky="ew")
        quoteLabel1.grid(row=0,column=7, sticky="ew")
        quoteLabel2.grid(row=1,column=7, sticky="ew")
        locationLabel.grid(row=0,column=8, sticky="ew")
        typeLabel.grid(row=0,column=9, sticky="ew")
        e_gradeLabel.grid(row=0,column=10, sticky="ew")
        e_yieldLabel.grid(row=0,column=11, sticky="ew")
        e_odLabel.grid(row=0,column=12, sticky="ew")
        e_idLabel.grid(row=0,column=13, sticky="ew")
        e_Length.grid(row=0,column=14, sticky="ew")
        e_Qty.grid(row=0,column=15, sticky="ew")
        sellcostLbsLabel1.grid(row=0,column=16, sticky="ew")
        sellcostLbsLabel2.grid(row=1,column=16, sticky="ew")
        uom.grid(row=0,column=17, sticky="ew")
        sellcostUOMLabel1.grid(row=0,column=18, sticky="ew")
        sellcostUOMLabel2.grid(row=1,column=18, sticky="ew")
        addCostLabel1.grid(row=0,column=19, sticky="ew")
        addCostLabel2.grid(row=1,column=19, sticky="ew")
        leadTimeLAbel.grid(row=0,column=20, sticky="ew")
        finalPriceLabel.grid(row=0,column=21,padx=(0,10), sticky="ew")
        ###################################################################
        ######################Defining List variables for various entry boxes######################
        global specialList
        
        specialList = {}
        


        #General Quote Form Variables
        cx_spec = []
        specialList["C_Specification"] = []
        specialList["C_Specification"].append(cx_spec)

        cx_grade = []
        specialList["C_Grade"] = []
        specialList["C_Grade"].append(cx_grade)

        cx_yield = []
        specialList["C_Yield"] = []
        specialList["C_Yield"].append(cx_yield
        )
        cx_od = []
        specialList["C_OD"] = []
        specialList["C_OD"].append(cx_od)

        cx_id = []
        specialList["C_ID"] = []
        specialList["C_ID"].append(cx_id)

        cx_len = []
        specialList["C_Length"] = []
        specialList["C_Length"].append(cx_len)
        
        cx_qty = []
        specialList["C_Qty"] = []
        specialList["C_Qty"].append(cx_qty)
        
        
        quoteYesNo = []
        specialList["C_Quote Yes/No"] = []
        specialList["C_Quote Yes/No"].append(quoteYesNo)

        e_location = []
        specialList["E_Location"] = []
        specialList["E_Location"].append(e_location)

        e_type = []
        specialList["E_Type"] = []
        specialList["E_Type"].append(e_type)

        e_grade = []
        specialList["E_Grade"] = []
        specialList["E_Grade"].append(e_grade)

        e_yield = []
        specialList["E_Yield"] = []
        specialList["E_Yield"].append(e_yield)

        e_od = []
        specialList["E_OD"] = []
        specialList["E_OD"].append(e_od)

        e_id = []
        specialList["E_ID"] = []
        specialList["E_ID"].append(e_id)

        e_len = []
        specialList["E_Length"] = []
        specialList["E_Length"].append(e_len)

        e_qty = []
        specialList["E_Qty"] = []
        specialList["E_Qty"].append(e_qty)

        sellCostLBS = []
        specialList["E_Selling Cost/LBS"] = []
        specialList["E_Selling Cost/LBS"].append(sellCostLBS)

        e_uom = []
        specialList["E_UOM"] = []
        specialList["E_UOM"].append(e_uom)
        
        sellCostUOM = []
        specialList["E_Selling Cost/UOM"] = []
        specialList["E_Selling Cost/UOM"].append(sellCostUOM)

        addCost = []
        specialList["E_Additional_Cost"] = []
        specialList["E_Additional_Cost"].append(addCost)

        leadTime = []
        specialList["E_LeadTime"] = []
        specialList["E_LeadTime"].append(leadTime)

        finalCost = []
        specialList["E_Final Price"] = []
        specialList["E_Final Price"].append(finalCost)

        # specialList = [[quoteYesNo],[e_location], [e_type], [e_grade], [e_yield], [e_od], [e_id], [e_len], [e_qty], [sellCostLBS], [sellCostUOM],
        # [e_uom], [addCost], [leadTime], [finalCost]]
        ###########################################################################################
        
        # var = tk.StringVar()
        # spec = ttk.Entry(entryFrame,textvariable=var, foreground='blue',background='white',width=5)
        # spec.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=0,width=2,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        
        while len(quoteYesNo)<1:
            addRow()
        
        button_dict = {}

        addRowbut = tk.Button(controlFrame, image=add_img, command=addRow,borderwidth=0, background=controlFrame["bg"])
        addRowbut.image = add_img
        addRowbut.grid(row=0,column=1)
        button_dict[addRowbut] = [add_img, add_img2]
        addRowbut.bind("<Enter>", on_enter)
        addRowbut.bind("<Leave>", on_leave)

        deleteRowbut = tk.Button(controlFrame, image=delete_img, text="Delete Row",command=deleteRow,borderwidth=0, background=controlFrame["bg"])
        deleteRowbut.image = delete_img
        deleteRowbut.grid(row=1,column=1)
        button_dict[deleteRowbut] = [delete_img, delete_img2]
        deleteRowbut.bind("<Enter>", on_enter)
        deleteRowbut.bind("<Leave>", on_leave)
        
            
        Previewbut = tk.Button(controlFrame, image=preview_img,text="Preview",command=create_pdf,borderwidth=0, background=controlFrame["bg"])
        Previewbut.image = preview_img
        Previewbut.grid(row=2,column=1)
        button_dict[Previewbut] = [preview_img, preview_img2]
        Previewbut.bind("<Enter>", on_enter)
        Previewbut.bind("<Leave>", on_leave)

        submitButton = tk.Button(controlFrame, image=submit_img, text="Submit",command=lambda: uploadDf(conn, quoteDf),borderwidth=0, background=controlFrame["bg"])
        submitButton.image = submit_img
        submitButton.configure(state='disable')
        submitButton.grid(row=3,column=1)
        button_dict[submitButton] = [submit_img, submit_img2]
        submitButton.bind("<Enter>", on_enter)
        submitButton.bind("<Leave>", on_leave)

        def on_closing():
            if messagebox.askokcancel("Quit", "Do you want to quit?"):
                # mainRoot.destroy()
                conn.close()
                root.destroy()
                sys.exit()
        mainRoot.protocol("WM_DELETE_WINDOW", on_closing)
        root.protocol("WM_DELETE_WINDOW", on_closing)
    except Exception as e:
        raise e
    
    # root.mainloop()

# conn = get_connection()
# # conn=None
# mainRoot = tk.Tk()
# user = "Imam"
# quoteGenerator(mainRoot, user, conn)
# mainRoot.mainloop()
