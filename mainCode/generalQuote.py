import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from customComboboxV2 import myCombobox
from tkcalendar import DateEntry
from datetime import date
import sys
import pandas as pd





#Calendar


class MyDateEntry(DateEntry):
    def __init__(self, master=None, **kw):
        DateEntry.__init__(self, master=master, date_pattern='mm.dd.yyyy',**kw)
        # add black border around drop-down calendar
        self._top_cal.configure(bg='black', bd=1)
        # add label displaying today's date below
        tk.Label(self._top_cal, bg='gray90', anchor='w',
                 text='Today: %s' % date.today().strftime('%x')).pack(fill='both', expand=1)


class ResizingCanvas(tk.Canvas):
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





def quoteGenerator(mainRoot,user):
    def OnMouseWheel(event):
        entryCanvas.yview_scroll(int(-1*(event.delta/120)), "units")
    # def yesNo(inStr,wName):
    #     if inStr == "No":
    #         print(wName)
    #         print(specialList[0][-1])
    #         specialList[0][-1].configure(validate = 'focusout')
    #         for i in range(1,len(specialList)):
    #             specialList[i][-1]['validate'] = 'none'
    #             specialList[i][-1].delete(0, tk.END)
    #             specialList[i][-1].insert(0, "NA")
    #             specialList[i][-1].configure(state='disabled')
    #     elif specialList[0][-1]['state'] == 'disabled':
    #         for i in range(1,len(specialList)):
    #             specialList[i][-1].delete(0, tk.END)
    #             specialList[i][-1].insert(0, "NA")
    #             specialList[i][-1].configure(state='normal')
    def intFloat(inStr,acttyp):
        # if acttyp == '1': #insert
        if inStr == '' or inStr == "NA":
            return True
        try:
            float(inStr)
            print('value:', inStr)
        except ValueError:
            messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value",parent=entryFrame)
            return False
        return True

    def intChecker(inStr,acttyp):
        # if acttyp == '1': #insert
        if inStr == '' or inStr == "NA":
            return True
        if not inStr.isdigit():
            messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value",parent=entryFrame)
            return False
        return True
    def on_configure(event):
        # update scrollregion after starting 'mainloop'
        # when all widgets are in canvas
        entryCanvas.configure(scrollregion=entryCanvas.bbox('all'),width=1700,height=400)
    def returnTohome():
        root.withdraw()
        mainRoot.deiconify()
        mainRoot.state('zoomed')
    mainRoot.withdraw()
    global row_num
    row_num=0


    df = pd.read_excel("sampleInventory.xlsx")

    print(df)



    count = 0
    root = tk.Toplevel(mainRoot)
    root.state('zoomed')
    cxFrame = tk.Frame(root)#,highlightbackground="blue", highlightthickness=2)
    cxFrame2 = tk.Frame(root)#,highlightbackground="blue", highlightthickness=2)
    m_entryFrame = tk.Frame(root,highlightbackground="black", highlightthickness=2)#width=1700,height=300
    entryCanvas = tk.Canvas(m_entryFrame)#,width=1930,height=400)
    # entryCanvas = ResizingCanvas(m_entryFrame,width=1700, height=400)
    xscrollbar=ttk.Scrollbar(m_entryFrame,orient=tk.HORIZONTAL, command=entryCanvas.xview)
    
    entryCanvas.config(xscrollcommand = xscrollbar.set)

    #Defining frame inside canvas
    entryFrame = tk.Frame(entryCanvas)
    
    entryFrame.bind('<Configure>', on_configure)

    yscrollbar=ttk.Scrollbar(m_entryFrame,orient="vertical", command=entryCanvas.yview)
    entryCanvas.bind_all("<MouseWheel>", OnMouseWheel)
    
    entryCanvas.config(yscrollcommand = yscrollbar.set)
    databaseFrame = tk.Frame(root,height=500)


    # 
    
    
    # dbFrame = ttk.Frame(root)
    # canvas =tk.Canvas(root)
    # canvas.pack()
    # cxFrame.place(in_=root, anchor="nw",x=10,y=10,relwidth=0.55)
    root.grid_rowconfigure(0,weight=1)
    root.grid_rowconfigure(1,weight=1)
    root.grid_rowconfigure(2,weight=1)
    root.grid_columnconfigure(0,weight=1)
    root.grid_columnconfigure(1,weight=1)

    cxFrame.grid(row=0, column=0,pady=(24,0), padx=(30,0),rowspan=10,sticky="new")
    cxFrame2.grid(row=0, column=1,pady=(24,10),columnspan=3, padx=(30,120),rowspan=10,sticky="nw")
    m_entryFrame.grid(row=0, column=0,pady=(160,0),columnspan=3, padx=(0,0),sticky="nsew")
    xscrollbar.grid(row=1,column=0,sticky="nsew")
    yscrollbar.grid(row=0,column=1,sticky="nsew")
    entryCanvas.grid(row=0,column=0, sticky="nsew")
    databaseFrame.grid(row=1,column=0, sticky=tk.NSEW)
    
    # myscrollbar.grid(row=1,column=0,sticky="ns")
    # dbFrame.grid(row=2, column=0,pady=(24,0),columnspan=3, padx=(30,0),sticky="nsew")

    #Configuing Canvas
    # entryCanvas.configure(xscrollcommand=myscrollbar.set)
    entryCanvas.create_window((0,0),window=entryFrame,tags='expand')
    
    # entryFrame.grid(row=0,column=0)
    
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

    cxFrame2.grid_columnconfigure(0, weight=1) # For column 0
    cxFrame2.grid_columnconfigure(1, weight=1) # For column 1


    m_entryFrame.grid_rowconfigure(0, weight=1) # For row 0
    m_entryFrame.grid_rowconfigure(1, weight=1) # For row 1

    m_entryFrame.grid_columnconfigure(0, weight=1) # For column 0
    m_entryFrame.grid_columnconfigure(1, weight=1) # For column 1

    entryFrame.grid_rowconfigure(0, weight=1) # For row 0
    entryFrame.grid_rowconfigure(1, weight=1) # For row 1

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
    

    
    home_img = tk.PhotoImage(master=root, file="home.png")
    
    cx_list = ('Perfect Tools Factory LLC', 'Accurate Edge Manufacturin & Coating LLC', 'High precision Manufacturing LLC', 
    'NTS Middle East FZCO', 'Ultra Corpotech', 'Falcon Group of Companies')

    item_list = ('A4140', 'A4140M', 'A4330V', 'A4715', 'BS708M40', 'A4145M', '4542','4462')

    cxLabel = tk.Label(cxFrame, text="Customer Details")
    lb1 = tk.Label(cxFrame,text="Prepared By")
    lb2 = tk.Label(cxFrame,text="Date")
    lb3 = tk.Label(cxFrame,text="Customer Name")
    lb4 = tk.Label(cxFrame,text="Location/Address")
    lb5 = tk.Label(cxFrame,text="Email")
    lb6 = tk.Label(cxFrame,text="Payment Terms")
    lb7 = tk.Label(cxFrame2,text="Validity")
    lb8 = tk.Label(cxFrame2,text="Additional Comments")
    cxLabel.grid(row=0,column=0)
    lb1.grid(row=1,column=0)
    lb2.grid(row=1,column=1)
    lb3.grid(row=3,column=0)
    lb4.grid(row=3,column=1)
    lb5.grid(row=3,column=2)
    lb6.grid(row=3,column=3)
    lb7.grid(row=0,column=0)
    lb8.grid(row=0,column=1)

    
    prep_by = ttk.Entry(cxFrame)
    prep_by.insert(tk.END, user)
    prep_by.grid(row=2,column=0)
    prep_by.config(state= "disabled")
    # myCombobox(df,root,item_list,frame=cxFrame,row=2,column=1,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
    cal = MyDateEntry(master=cxFrame, width=17, selectmode='day')
    cal.grid(row=2, column=1)
    myCombobox(df,root,item_list,frame=cxFrame,row=4,column=0,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
    myCombobox(df,root,item_list,frame=cxFrame,row=4,column=1,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
    myCombobox(df,root,item_list,frame=cxFrame,row=4,column=2,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
    myCombobox(df,root,item_list,frame=cxFrame,row=4,column=3,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew")

    myCombobox(df,root,item_list,frame=cxFrame2,row=1,column=0,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
    myCombobox(df,root,item_list,frame=cxFrame2,row=1,column=1,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
    
    # myCombobox(df,root,item_list,frame=entryFrame,row=1,column=1,width=10,list_bd = 0,foreground='blue', background='white',sticky = tk.EW)


    home_button = tk.Button(root, image=home_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"],command=returnTohome)
    home_button.image = home_img #Preventing image to go into garbage
    home_button.grid(row=0,column=2,stick="nw",padx=(0,50),pady=(10,0))
    # home_button.place(x=1600,y=-10,relx=0.1,rely=0.1,anchor="sw")
    #######################################
    
    ########################################
    
    # canvas.create_line(200, 100, 1800, 100, fill="green")

    #################Entry Form Section##############################################
    ######################defining labels############################################
    specLabel = tk.Label(entryFrame, text="Specification")
    gradeLabel = tk.Label(entryFrame, text="Grade")
    yieldLabel = tk.Label(entryFrame, text="Yield")
    odLabel = tk.Label(entryFrame, text="OD")
    idLabel = 	tk.Label(entryFrame, text="ID")
    lengthLabel = tk.Label(entryFrame, text="Length")
    qtyLabel = tk.Label(entryFrame, text="Qty")
    quoteLabel = tk.Label(entryFrame, text="Quote Yes/No")
    locationLabel = tk.Label(entryFrame, text="Location")
    typeLabel = tk.Label(entryFrame, text="Type")
    e_gradeLabel = tk.Label(entryFrame, text="Grade")
    e_yieldLabel = tk.Label(entryFrame, text="Yield")
    e_odLabel = tk.Label(entryFrame, text="OD")
    e_idLabel = tk.Label(entryFrame, text="ID")
    e_Length = tk.Label(entryFrame, text="Length")
    e_Qty = tk.Label(entryFrame, text="Qty")
    sellcostLbsLabel = tk.Label(entryFrame, text="Selling Cost/LBS")
    uom = tk.Label(entryFrame, text="UOM")
    sellcostUOMLabel = tk.Label(entryFrame, text="Selling Cost/UOM")
    addCostLabel = tk.Label(entryFrame, text="Additional Cost")
    leadTimeLAbel = tk.Label(entryFrame, text="Lead Time")
    finalPriceLabel = tk.Label(entryFrame, text="Final Price")



    specLabel.grid(row=0,column=0,padx=(15,0))
    gradeLabel.grid(row=0,column=1)
    yieldLabel.grid(row=0,column=2)
    odLabel.grid(row=0,column=3)
    idLabel.grid(row=0,column=4)
    lengthLabel.grid(row=0,column=5)
    qtyLabel.grid(row=0,column=6)
    quoteLabel.grid(row=0,column=7)
    locationLabel.grid(row=0,column=8)
    typeLabel.grid(row=0,column=9)
    e_gradeLabel.grid(row=0,column=10)
    e_yieldLabel.grid(row=0,column=11)
    e_odLabel.grid(row=0,column=12)
    e_idLabel.grid(row=0,column=13)
    e_Length.grid(row=0,column=14)
    e_Qty.grid(row=0,column=15)
    sellcostLbsLabel.grid(row=0,column=16)
    uom.grid(row=0,column=17)
    sellcostUOMLabel.grid(row=0,column=18)
    addCostLabel.grid(row=0,column=19)
    leadTimeLAbel.grid(row=0,column=20)
    finalPriceLabel.grid(row=0,column=21)
    ###################################################################
    ######################Defining List variables for various entry boxes######################
    global specialList
    specialList = {}
    cx_spec = []
    cx_grade = []
    specialList["cx_grade"] = []
    specialList["cx_grade"].append(cx_grade)
    cx_yield = []
    specialList["cx_yield"] = []
    specialList["cx_yield"].append(cx_yield)
    cx_od = []
    specialList["cx_od"] = []
    specialList["cx_od"].append(cx_od)
    cx_id = []
    specialList["cx_id"] = []
    specialList["cx_id"].append(cx_id)
    cx_len = []
    cx_qty = []
    
    
    quoteYesNo = []
    specialList["quoteYesNo"] = []
    specialList["quoteYesNo"].append(quoteYesNo)
    e_location = []
    specialList["e_location"] = []
    specialList["e_location"].append(e_location)

    e_type = []
    specialList["e_type"] = []
    specialList["e_type"].append(e_type)

    e_grade = []
    specialList["e_grade"] = []
    specialList["e_grade"].append(e_grade)

    e_yield = []
    specialList["e_yield"] = []
    specialList["e_yield"].append(e_yield)

    e_od = []
    specialList["e_od"] = []
    specialList["e_od"].append(e_od)

    e_id = []
    specialList["e_id"] = []
    specialList["e_id"].append(e_id)

    e_len = []
    specialList["e_len"] = []
    specialList["e_len"].append(e_len)

    e_qty = []
    specialList["e_qty"] = []
    specialList["e_qty"].append(e_qty)

    sellCostLBS = []
    specialList["sellCostLBS"] = []
    specialList["sellCostLBS"].append(sellCostLBS)

    sellCostUOM = []
    specialList["sellCostUOM"] = []
    specialList["sellCostUOM"].append(sellCostUOM)

    e_uom = []
    specialList["e_uom"] = []
    specialList["e_uom"].append(e_uom)

    addCost = []
    specialList["addCost"] = []
    specialList["addCost"].append(addCost)

    leadTime = []
    specialList["leadTime"] = []
    specialList["leadTime"].append(leadTime)

    finalCost = []
    specialList["finalCost"] = []
    specialList["finalCost"].append(finalCost)

    # specialList = [[quoteYesNo],[e_location], [e_type], [e_grade], [e_yield], [e_od], [e_id], [e_len], [e_qty], [sellCostLBS], [sellCostUOM],
    # [e_uom], [addCost], [leadTime], [finalCost]]
    ###########################################################################################
    
    # var = tk.StringVar()
    # spec = ttk.Entry(entryFrame,textvariable=var, foreground='blue',background='white',width=5)
    # spec.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
    # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=0,width=2,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
    def addRow():
        
        row_num = len(quoteYesNo)
        cx_spec.append(ttk.Entry(entryFrame,width=15))
        cx_spec[-1].grid(row=1+row_num,column=0,padx=(15,0))
        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        cx_grade.append(ttk.Entry(entryFrame,width=15))
        cx_grade[-1].grid(row=1+row_num,column=1)
        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        cx_yield.append(ttk.Entry(entryFrame,width=15))
        cx_yield[-1].grid(row=1+row_num,column=2)
        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=3,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        
        vcmd = root.register(intFloat)
        cx_od.append(ttk.Entry(entryFrame, width=10,validate = "key",
                validatecommand=(vcmd, '%P','%d')))
        cx_od[-1].grid(row=1+row_num,column=3)
        # cx_od['validatecommand'] = (cx_od.register(intFloat),'%P','%d')



        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=4,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        cx_id.append(ttk.Entry(entryFrame, width=10, validate = "key"))
        cx_id[-1].grid(row=1+row_num,column=4)
        cx_id[-1]['validatecommand'] = (cx_id[-1].register(intFloat),'%P','%d')
        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=5,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        cx_len.append(ttk.Entry(entryFrame, width=10, validate = "key"))
        cx_len[-1].grid(row=1+row_num,column=5)
        cx_len[-1]['validatecommand'] = (cx_len[-1].register(intFloat),'%P','%d')
        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=6,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        cx_qty.append(ttk.Entry(entryFrame, width=10, validate = "key"))
        cx_qty[-1].grid(row=1+row_num,column=6)
        cx_qty[-1]['validatecommand'] = (cx_qty[-1].register(intChecker),'%P','%d')
        quoteYesNo.append(myCombobox(df,root,["Yes","No","Other"],frame=entryFrame,row=1+row_num,column=7,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # quoteYesNo[-1]['validate']='focusout'
        # quoteYesNo[-1]['validatecommand'] = (quoteYesNo[-1].register(yesNo),'%P','%W')
        e_location.append(myCombobox(df,root,["Dubai","Singapore","USA","UK"],frame=entryFrame,row=1+row_num,column=8,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_location[-1].config(textvariable="NA", state='disabled')
        e_type.append(myCombobox(df,root,["THF","BR"],frame=entryFrame,row=1+row_num,column=9,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_type[-1].config(textvariable="NA", state='disabled')
        e_grade.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=10,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_grade[-1].config(textvariable="NA", state='disabled')
        e_yield.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=11,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_yield[-1].config(textvariable="NA", state='disabled')
        e_od.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=12,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_od[-1].config(textvariable="NA", state='disabled')
        e_od[-1][0]['validate']='key'
        e_od[-1][0]['validatecommand'] = (e_od[-1][0].register(intFloat),'%P','%d')
        e_id.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=13,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_id[-1].config(textvariable="NA", state='disabled')
        e_len.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=14,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_len[-1].config(textvariable="NA", state='disabled')
        e_qty.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=15,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_qty[-1].config(textvariable="NA", state='disabled')
        sellCostLBS.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=16,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # sellCostLBS[-1].config(textvariable="NA", state='disabled')
        e_uom.append(myCombobox(df,root,["Inch","Each"],frame=entryFrame,row=1+row_num,column=17,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # e_uom[-1].config(textvariable="NA", state='disabled')
        sellCostUOM.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=18,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # sellCostUOM[-1].config(textvariable="NA", state='disabled')
        addCost.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=19,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # addCost[-1].config(textvariable="NA", state='disabled')
        leadTime.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=20,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # leadTime[-1].config(textvariable="NA", state='disabled')
        finalCost.append(myCombobox(df,root,item_list,frame=entryFrame,row=1+row_num,column=21,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
        # finalCost[-1].config(textvariable="NA", state='disabled')
        
        
        # row_num+=1
    while len(quoteYesNo)<1:
        addRow()
        
        



    addRowbut = tk.Button(databaseFrame, text="Add Row",command=addRow)
    addRowbut.grid(row=0,column=1, sticky=tk.NW)
    

    # scrollbar = tk.Scrollbar(entryFrame, orient='horizontal')
    # scrollbar.pack(side = tk.BOTTOM, fill = tk.X)





    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            # mainRoot.destroy()
            root.destroy()
            sys.exit()
    mainRoot.protocol("WM_DELETE_WINDOW", on_closing)
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # root.mainloop()


# mainRoot = tk.Tk()
# user = "Imam"
# quoteGenerator(mainRoot, user)
# mainRoot.mainloop()
