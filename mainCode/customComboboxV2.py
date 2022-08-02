
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.tix import ButtonBox
import Tools

def formulaCalc(boxList, index):
    try:
        if boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" or boxList['E_Type'][0][index][0].get()=="HM":
            e_od = float(boxList["E_OD2"][0][index][1])
            e_id = float(boxList["E_ID2"][0][index][1])
        else:
            e_od = float(boxList["E_OD1"][0][index][1].get())
            e_id = float(boxList["E_ID1"][0][index][1].get())
        uom = boxList["E_UOM"][0][index][1].get()
        e_length = float(boxList["E_Length"][0][index][1].get())#int(boxList["E_Length"][0][index][1].get())
        sellCostLBS = float(boxList["E_Selling Cost/LBS"][0][index][1].get())
        wt = (e_od - e_id)/2
        # THF
        if boxList["E_Type"][0][index][1].get() == "THF" or boxList["E_Type"][0][index][1].get() == "TUI" or boxList["E_Type"][0][index][1].get() == "HT":
            # wt = (od-id)/2
            
            # mid_formula = ((od-wt)*wt*10.68)/12
            mid_formula = ((e_od-wt)*wt*10.68)/12 
            # For Each:
            if uom == "Each":
            # Selling cost/UOM = "SellingCost/LBS" * mid_formula * Length (rounded upto 2 decimal places)
                sellCostUOM = round((sellCostLBS * mid_formula * e_length),2)
            # For Inch
            else:
                # Selling cost/UOM = "SellingCost/LBS" * mid_formula (rounded upto 2 decimal places)
                sellCostUOM = round((sellCostLBS * mid_formula),2)
        elif boxList["E_Type"][0][index][1].get() == "BR" or boxList["E_Type"][0][index][1].get() == "HR" or boxList["E_Type"][0][index][1].get() == "HB" or boxList["E_Type"][0][index][1].get() == "HM":
            # BR
            # mid_formula = (od*od*2.71)/12
            mid_formula = ((e_od-wt)*wt*10.68)/12
            if uom == "Each":
                # For Each: 
                # Selling cost/UOM ="SellingCost/LBS" * mid_formula * Length 
                sellCostUOM = round((sellCostLBS * mid_formula * e_length),2)
            else:
                # For Inch: 
                # Selling cost/UOM ="SellingCost/LBS" * mid_formula	
                sellCostUOM = round((sellCostLBS * mid_formula),2)
        else:
            sellCostUOM = 0
        return sellCostUOM
    except Exception as ex:
        raise ex


def myCombobox(df,root,frame,row,column,width=10,list_bd = 0,foreground='blue', background='white',sticky = tk.EW,item_list=[],boxList={},cxDict={},val = None, pt=None, entpady=(5, 5), bakerDf = []):
    try:
        # def __init__(self,item_list,frame,row,column,width=10,list_bd = 0,foreground='blue', background='white',sticky = tk.EW):
        global checker
        checker = True
        
        # def paste(event):
        #     rows = root.clipboard_get().split('\n')
        #     for r, row in enumerate(rows):
        #         values = row.split('\t')
        #         for c, value in enumerate(values):
        #             table[r][c].set(value)
        def set_mousewheel(widget, command):
            try:
                """Activate / deactivate mousewheel scrolling when 
                cursor is over / not over the widget respectively."""
                widget.bind("<Enter>", lambda _: widget.bind_all('<MouseWheel>', command))
                widget.bind("<Leave>", lambda _: widget.unbind_all('<MouseWheel>'))
            except Exception as ex:
                raise ex

        def onMouseWheel(event):
            try:
                lbframe.list.yview_scroll(int(-1*(event.delta/220)), "units")
            except Exception as ex:
                raise ex
        if len(item_list) and type(item_list[0])==str:
            max_len = len(max(item_list, key=len))
            if max_len > width:
                width = max_len
        var = tk.StringVar(frame)
        ent = ttk.Entry(frame,textvariable=var, font=('Segoe UI', 10),foreground=foreground,background=background,width=width)
        lst = tk.Listbox(frame, bd=list_bd, background=background,width=width,height=20)

        lbframe = tk.Frame(root)
        lbframe.list_values = tk.StringVar(frame) 
        lbframe.list_values.set(item_list) 
        lbframe.list= tk.Listbox(lbframe, height=5, width= 10,
                                        listvariable=lbframe.list_values)
        lbframe.list.pack(side='left', fill="both", expand=True)
        # s = ttk.Scrollbar(lbframe, orient=tk.VERTICAL, command=lbframe.list.yview)
        # s.pack(side='right', fill = "y")
        # lbframe.list['yscrollcommand'] = s.set
        set_mousewheel(lbframe.list, onMouseWheel)
        # lbframe.list.bind("<MouseWheel>", onMouseWheel)
        ent.grid(row=row,column=column,sticky=tk.EW,padx=5,pady=entpady)



        def nextKey(key,i):
            try:
                keyList=boxList.keys()
                for k,v in enumerate(keyList):
                    if v==key:
                        newKey = list(keyList)[k+1]
                return newKey
            except Exception as ex:
                    raise ex
            
        
        # def entryUpdater(check,i,value,key, index):
        #     initValue = int(ent._name.split("y")[-1])
        #     ent._name = f'!entry{initValue+int(i//2)}'
        #     ent._w = f'.!toplevel.!frame3.!canvas.!frame.!entry{initValue+int(i//2)}'
        #     if check:
        #         ent.configure(state='disabled')
        #     else:
        #         ent.configure(state='normal')
        #     if value != "No" and value != "Yes":
        #         # if i ==2:
        #         #     key, index = keyFinder(boxList,(ent,var))
                
        #         newKey = nextKey(key,int(i//2))
        #         newList = filterList(boxList,newKey,index,df)
        #         # if value in newList:
        #         #     var.set(value)
        #         if len(newList)==1:
        #             var.set(newList[0])
        #         else:
        #             var.set("")
        #     ent._name = f'!entry{initValue}'
        #     ent._w = f'.!toplevel.!frame3.!canvas.!frame.!entry{initValue}'
                # ent.delete(0, tk.END)
        def addCostCalc(boxList,index):
            try:
                addCost = float(boxList["E_Additional_Cost"][0][index][1].get())
                sellCost = float(boxList["E_Selling Cost/UOM"][0][index][1].get())
                finalPrice = addCost+sellCost
                boxList["E_Final Price"][0][index][1].set(finalPrice)
            except Exception as ex:
                raise ex


        
        def list_hide(e=None):
            try:
                breakCheck = False
                # global get_input_checker
                global checker
                checker = False
                if lbframe.list.size()==1:
                    value = lbframe.list.get(0)
                    var.set(value)
                    #finding current key and index
                    if len(boxList):
                        key, index = keyFinder(boxList,(ent,var))
                        
                    else:
                        key, index = keyFinder(cxDict,(ent,var))
                    if key == "cus_long_name":
                        dataList = df[(df["cus_long_name"] == value)].values.tolist()[0]
                        cxDict["payment_term"][index][1].set(dataList[3])
                        # cxDict["payment_term"] = dataList[2]

                        cxDict["cus_address"][index][1].set(dataList[4])
                        # cxDict["cus_address"] = dataList[3]

                        cxDict["cus_email"][index][1].set(dataList[6])
                        # cxDict["cus_email"] = dataList[5]

                        cxDict["cus_phone"][index][1].set(dataList[5])

                        cxDict["cus_city_zip"] = dataList[7]
                        # val.lift()
                        val.focus()
                        breakCheck = True

                    elif key=='E_UOM' and boxList['C_Quote Yes/No'][0][index][0].get() != "Other":
                        sellCostUOM = formulaCalc(boxList, index)
                        boxList["E_Selling Cost/UOM"][0][index][1].set(sellCostUOM)
                        boxList["E_Additional_Cost"][0][index][0].focus()
                    elif key == "E_Additional_Cost":
                        addCost = float(boxList["E_Additional_Cost"][0][index][1].get())
                        sellCost = float*(boxList["E_Selling Cost/UOM"][0][index][1].get())
                        finalPrice = addCost+sellCost
                        boxList["E_Final Price"][0][index][0].set(finalPrice)

                    elif key == "E_Type":
                        boxList['E_Length'][0][index][0].focus()
                        breakCheck = True

                        
                    elif value != "Other" and value != "Yes" and value != "No"  and key != "E_UOM": #and key!='E_Location'
                        current_key = key
                        while True:
                            next_key = nextKey(current_key, index)
                            if next_key == "E_Length":
                                if boxList['E_ID1'][0][index][0].get() != '' and boxList['E_OD1'][0][index][0].get() != '':
                                    #Adding Baker Type Handler
                                    if boxList['E_Type'][0][index][0].get() == "HT":
                                        e_type_var = "THF"
                                    elif boxList['E_Type'][0][index][0].get() == "HB":
                                        e_type_var = "BR"
                                    else:
                                        e_type_var = boxList['E_Type'][0][index][0].get()
                                    if (boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR") and (boxList['E_ID2'][0][index][0].get() != '' and boxList['E_OD2'][0][index][0].get() != ''):
                                        # newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)
                                        newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())
                                                & (df["global_grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                                                & (df["od_in"]==float(boxList['E_OD2'][0][index][0].get())) & (df["od_in_2"]==float(boxList['E_ID2'][0][index][0].get()))]
                                        
                                    else:
                                        # newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)
                                        newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())
                                                & (df["global_grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                                                & (df["od_in"]==float(boxList['E_OD1'][0][index][0].get())) & (df["od_in_2"]==float(boxList['E_ID1'][0][index][0].get()))]
                                    # newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds','reserved_pieces', 'reserved_length_in', 'available_pieces', 'available_length_in']]
                                    newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in']]
                                    newDf = newDf[newDf['available_pieces']>0]
                                    boxList['E_Type'][0][index][0].focus()
                                    if pt is not None:
                                        pt.model.df = newDf
                                        pt.redraw()
                                        breakCheck = True
                                break
                            new_list = filterList(boxList,next_key,index,df)
                            if len(bakerDf) and next_key == 'E_Type':
                                new_list = [w.replace('THF', 'HT') for w in new_list]
                                new_list = [w.replace('BR', 'HB') for w in new_list]
                                new_list = [w.replace('HR', 'HM') for w in new_list]
                            cx_eags = {'E_Grade':'C_Grade','E_Yield':'C_Yield','E_OD1':'C_OD','E_ID1':'C_ID','E_OD2':'C_OD','E_ID2':'C_ID','E_Type':"C_Type"}
                            #if cx value in current list, set next key value = cx value
                            # print(bakerDf)
                            if next_key in ["E_OD1","E_ID1"] and len(new_list):
                                if boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" or boxList['E_Type'][0][index][0].get()=="HM":
                                    newList_id = filterList(boxList, 'E_ID2', index, df)
                                    try:
                                        boxList["E_OD2"][0][index][0].get()
                                        if len(newList_id)==1:
                                            boxList["E_ID2"][0][index][1].set(newList_id[0])
                                    except:
                                        try:
                                            od1,id1,od2,id2 = Tools.specialCase(root, boxList, pt, df, index, item_list=new_list)

                                            if od1 is not None and id1 is not None:
                                                boxList['E_OD1'][0][index][1].set(od1)
                                                boxList['E_ID1'][0][index][1].set(id1)

                                        except Exception as ex:
                                            raise ex
                                    break
                                # new_list2 = filterList(boxList,'E_ID',index,df)
                                #function call to pop up windows for 2 different od and id's and let it fill od id in main page writern variables for calculation
                                elif len(new_list)==1:
                                    boxList[next_key][0][index][1].set(new_list[0])
                                elif boxList['E_ID1'][0][index][0].get() == '' or boxList['E_OD1'][0][index][0].get() == '':
                                    if not len(bakerDf):
                                        if float(boxList[cx_eags[next_key]][0][index][0].get()) in list(map(lambda x: float(x),new_list)):
                                            if next_key != 'E_OD2' and next_key != 'E_ID2':
                                                boxList[next_key][0][index][1].set(float(boxList[cx_eags[next_key]][0][index][0].get()))
                                    else:#removing .get() from customer side 
                                        if float(boxList[cx_eags[next_key]][0][index][0]) in list(map(lambda x: float(x),new_list)):
                                            if next_key != 'E_OD2' and next_key != 'E_ID2':
                                                boxList[next_key][0][index][1].set(float(boxList[cx_eags[next_key]][0][index][0]))
                                
                                else:
                                    break
                            elif not len(bakerDf) and next_key == 'E_Type':#2n case will be handled len(bakerDf) and str(boxList[cx_eags[next_key]][0][index][0]).upper() in list(map(lambda x: str(x).upper(),new_list))
                                boxList[next_key][0][index][1].set("")
        
                            elif next_key == 'E_Spec' and len(bakerDf):
                                boxList[next_key][0][index][1].set("")
                            
                            elif not len(bakerDf) and next_key != 'E_Spec' and str(boxList[cx_eags[next_key]][0][index][0].get()).upper() in list(map(lambda x: str(x).upper(),new_list)):
                                if not(next_key == 'E_OD2' or next_key == 'E_ID2'):
                                    boxList[next_key][0][index][1].set(str(boxList[cx_eags[next_key]][0][index][0].get()).upper())
                                    
                            elif len(bakerDf) and str(boxList[cx_eags[next_key]][0][index][0]).upper() in list(map(lambda x: str(x).upper(),new_list)): #Removing .get() to get direct customer data
                                if not(next_key == 'E_OD2' or next_key == 'E_ID2'):
                                    boxList[next_key][0][index][1].set(str(boxList[cx_eags[next_key]][0][index][0]).upper())
                            elif len(new_list)==1:
                                if not(next_key == 'E_OD2' or next_key == 'E_ID2'):
                                    boxList[next_key][0][index][1].set(new_list[0])
                            
                            elif next_key == list(boxList.keys())[0]:
                                break
                            else:
                                if next_key != 'E_OD2' and next_key != 'E_ID2':
                                    if next_key=='E_Spec' and not len(bakerDf):
                                        pass
                                    else:
                                        boxList[next_key][0][index][1].set("")
                            
                            current_key = next_key
                            
                    elif value == "Yes" or value == "No"or value == "Other":
                        keyIndex= list(boxList.keys()).index('E_Location')
                        for i in range(keyIndex,len(list(boxList.keys()))):
                            newKey = list(boxList.keys())[i]
                            if value == "Yes" or value == "Other":
                                if newKey != 'E_OD2' and newKey != 'E_ID2':
                                    if newKey=='E_Spec' and not len(bakerDf):
                                        pass
                                    else:
                                        boxList[newKey][0][index][1].set("")
                                        boxList[newKey][0][index][0].configure(state='normal')
                            else:
                                if newKey != 'E_OD2' and newKey != 'E_ID2' and newKey != 'E_Spec':
                                    boxList[newKey][0][index][1].set("NA")
                                    boxList[newKey][0][index][0].configure(state='disabled')
                                elif newKey == 'E_Spec' and len(bakerDf):
                                    boxList[newKey][0][index][1].set("NA")
                                    boxList[newKey][0][index][0].configure(state='disabled')
                        
                else:
                    if len(boxList):
                        key, index = keyFinder(boxList,(ent,var))
                        if key=='E_Additional_Cost':
                            if boxList["E_Additional_Cost"][0][index][1].get()!='' and boxList["E_Selling Cost/UOM"][0][index][1].get() != '':
                                addCostCalc(boxList,index)
                                # addCost = float(boxList["E_Additional_Cost"][0][index][1].get())
                                # sellCost = float(boxList["E_Selling Cost/UOM"][0][index][1].get())
                                # finalPrice = addCost+sellCost
                                # boxList["E_Final Price"][0][index][1].set(finalPrice)
                                        
                    

                            
                            

                    # # if value=="No":
                    # # if (ent,var) in boxList['quoteYesNo'][0]:
                    # init_name = var._name
                checker = True
                
                # var.set(value)
                    # key, index = keyFinder(boxList,(ent,var))
                    # start_num = int(init_name.split("VAR")[-1])
                    # for i in range(start_num+2,start_num+30,2):
                    #     # var._name = 
                    #     print(i)
                    #     var._name = f'PY_VAR{i}'
                    #     if value == "No":
                    #         var.set("NA")
                    #         entryUpdater(True,i-start_num,value,key, index)
                    #     elif value == "Yes":
                    #         entryUpdater(False,i-start_num,value,key, index)
                    #         var.set("")
                    #     # else:
                                        
                    #     #     entryUpdater(False,i-start_num,value,key, index)
                            
                                
                    #     #         pass
                    # var._name = init_name
                    # var.set(value)
                    
                lbframe.list.delete(0, tk.END)
                lbframe.place_forget()
                if breakCheck:# or e.keysym=="Tab":
                    return "break"
            except Exception as ex:
                raise ex

        def list_input(e):
            try:
                print(e.type)
                print(e.num)
                listCheck = True
                if ent.get()=='':
                    if len(boxList):
                        key, index = keyFinder(boxList,(ent,var))
                        if boxList['C_Quote Yes/No'][0][index][1].get() == "Other":
                            listCheck = False
                        elif key=='E_UOM':
                            newList = item_list
                        else:
                            newList = filterList(boxList,key,index,df)
                        
                    else:
                        key, index = keyFinder(cxDict,(ent,var))
                        newList = df["cus_long_name"].values.tolist() 
                    lbframe.list.delete(0, tk.END)
                    if listCheck and len(newList):
                        for item in newList:
                            lbframe.list.insert(tk.END, item)
                            lbframe.list.itemconfigure(tk.END, foreground="black")
                        if not newList[0]=='':  
                            if (key=='E_OD1' or key =='E_ID1'):
                                if boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR":
                                    pass
                                else:
                                    lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                                    lbframe.list.focus()
                                    lbframe.list.select_set(0)
                            else:
                                lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                                lbframe.list.focus()
                                lbframe.list.select_set(0)
                elif len(boxList):
                    key, index = keyFinder(boxList,(ent,var))
                    if len(lbframe.list_values.get()) != 2 and boxList['C_Quote Yes/No'][0][index][0].get()=="Yes" and len(lbframe.list_values.get()) != 0:
                        lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                        lbframe.list.focus()
                        lbframe.list.select_set(0)
                elif len(cxDict):
                    key, index = keyFinder(cxDict,(ent,var))
                    if key == "cus_long_name":
                        lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                        lbframe.list.focus()
                        lbframe.list.select_set(0)
                if e.num == 1 or e.num=='??' and e.keysym == '??':
                    root.update()
                    ent.focus()
                    print("Activating Entry box")
                    # time.sleep(0.1)
                    # ent.focus_set()
                    # return "break"
            except Exception as ex:
                raise ex
            
            
                    

        def list_up(_):
            try:
                if not lbframe.list.curselection()[0]:
                    
                    ent.focus()
                    list_hide()
            except Exception as e:
                raise e
                    


        def get_selection(e):
            try:
                value = lbframe.list.get(lbframe.list.curselection())
                var.set(value)
                list_hide(e)
                ent.focus()
                ent.icursor(tk.END)
            except Exception as ex:
                raise ex
        
        ent.bind("<1>", list_input)
        ent.bind('<Down>', list_input)
        ent.bind('<Return>', list_hide)
        ent.bind('<Tab>',list_hide)
        # ent.bind("<Leave>",list_up)
        # ent.bind('<Escape>', list_hide)

        lbframe.list.bind("<Enter>", list_input)
        lbframe.list.bind('<Up>', list_up)
        lbframe.list.bind('<Return>', get_selection)
        lbframe.list.bind('<Double-Button-1>', get_selection)
        lbframe.list.bind('<Tab>',get_selection)
        lbframe.list.bind('<Escape>',list_hide)
        lbframe.list.bind("<Leave>", list_hide)
        # ent.bind('<Down>', list_input)
        # # ent.bind('<Return>', list_hide)

        # lbframe.list.bind('<Up>', list_up)
        # lbframe.list.bind('<Return>', get_selection)

        # return ent,lbframe.list,var

        def keyFinder(dict, tupleValue):
            try:
                for key, value in dict.items():
                    for index in range(len(value[0])):
                        
                        if value[0][index] == tupleValue:
                            return key,index
            except Exception as e:
                raise e
        def filterList(boxList,key,index,df):
            try:
                #handling baker mapping for thf ht and br hb
                if boxList['E_Type'][0][index][0].get() == "HT":
                    e_type_var = "THF"
                elif boxList['E_Type'][0][index][0].get() == "HB":
                    e_type_var = "BR"
                else:
                    e_type_var = boxList['E_Type'][0][index][0].get()
                newList = []
                if key == 'C_Quote Yes/No':
                    newList = ["Yes", "No", "Other"]
                elif key== 'E_Location':
                    if boxList['C_Quote Yes/No'][0][index][0].get() != '':
                        newList = list(df["site"].unique())
                elif key == 'E_Type':
                    newList = ["THF","BR", "TUI", "HR"]
                    # #filter df based on e_location and make unique column of e_type
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())]
                    # newList = list(new_df["material_type"].unique())

                elif key == 'E_Grade':
                    #filter df based on e_location and make unique column of e_type
                    
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)]#boxList['E_Type'][0][index][0].get())]
                    new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())]#boxList['E_Type'][0][index][0].get())]
                    
                    newList = list(new_df["global_grade"].unique())
                elif key == 'E_Yield':
                    #filter df based on e_location and make unique column of e_type
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)#boxList['E_Type'][0][index][0].get())
                    new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())#boxList['E_Type'][0][index][0].get())
                    & (df["global_grade"]==boxList['E_Grade'][0][index][0].get())]
                    newList = list(new_df["heat_condition"].unique())
                elif key == 'E_OD1':
                    #filter df based on e_location and make unique column of e_type
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)#==boxList['E_Type'][0][index][0].get())
                    new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())#==boxList['E_Type'][0][index][0].get())
                    & (df["global_grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())]
                    newList = list(new_df["od_in"].unique())
                elif key == 'E_OD2':
                    #filter df based on e_location and make unique column of e_type
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)#==boxList['E_Type'][0][index][0].get())
                    new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())#==boxList['E_Type'][0][index][0].get())
                    & (df["global_grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())]
                    newList = list(new_df["od_in"].unique())
                elif key == 'E_ID1':
                    try:
                        #filter df based on e_location and make unique column of e_type
                        # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)#boxList['E_Type'][0][index][0].get())
                        new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())#boxList['E_Type'][0][index][0].get())
                        & (df["global_grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                        & (df["od_in"]==float(boxList['E_OD1'][0][index][0].get()))]
                        newList = list(new_df["od_in_2"].unique())
                
                    except Exception as ex:
                        newList = []
                elif key == 'E_ID2':
                    try:
                        #filter df based on e_location and make unique column of e_type
                        # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==boxList['E_Type'][0][index][0].get())
                        new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())
                        & (df["global_grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                        & (df["od_in"]==float(boxList['E_OD2'][0][index][0].get()))]
                        newList = list(new_df["od_in_2"].unique())
                
                    except Exception as ex:
                        newList = []

                elif key == "cus_long_name":
                    #filter df based on e_location and make unique column of e_type
                    new_df = df[(df["cus_long_name"] == boxList["cus_long_name"][0][index][0].get())] #
                    newList = new_df.values.tolist()
                try:
                    if key != 'C_Quote Yes/No':
                        newList.sort()
                except:
                    pass
                return newList
            except Exception as e:
                raise e




        def get_input(*args):
            try:
                if checker:
                    newList = item_list
                    check = True
                    # print(boxList)
                    if len(boxList):# and len(lbframe.list.get(0,tk.END))!=1
                        # boxList["quoteYesNo"]
                        key, index = keyFinder(boxList,(ent,var))   
                        # print(key, index)
                        if key != "C_Quote Yes/No" and boxList['E_Location'][0][index][0].get()!='' and key != "cus_long_name":# and key != "E_UOM":#and key != 'E_Location' 
                            if not str(boxList["C_Quote Yes/No"][0][index][0].get().upper()).startswith("Y"):
                                check = False
                                if key == "E_Additional_Cost" and boxList["E_Additional_Cost"][0][index][1].get()!='':
                                    addCostCalc(boxList, index)
                            elif (key == 'E_ID1' or key == 'E_OD1') and (boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR"):
                                check = False
                                if key == 'E_ID1':
                                    # newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==boxList['E_Type'][0][index][0].get())
                                    newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())
                                                & (df["global_grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                                                & (df["od_in"]==float(boxList['E_OD2'][0][index][0])) & (df["od_in_2"]==float(boxList['E_ID2'][0][index][0]))]
                                    # newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds','reserved_pieces', 'reserved_length_in', 'available_pieces', 'available_length_in']]
                                    newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in']]
                                    newDf = newDf[newDf['available_pieces']>0]
                                    boxList['E_Length'][0][index][0].focus()
                                    if pt is not None:
                                        pt.model.df = newDf
                                        pt.redraw()
                                        breakCheck = True
                            
                            else:
                                if key != "E_UOM":
                                    newList = filterList(boxList,key,index,df)
                        
                    elif len(cxDict):
                        key, index = keyFinder(cxDict,(ent,var))
                        newList = df["cus_long_name"].values.tolist()
                    
                    # lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                    
                    lbframe.list.delete(0, tk.END)
                    if len(newList)==1:
                        if (key == 'E_ID1' or key == 'E_OD1') and (boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR"):
                            pass
                        else:
                            var.set(newList[0])
                    else:
                        string = var.get()

                        if string and check:
                            for item in newList:
                                if str(item).upper().startswith(str(string).upper()):
                                    lbframe.list.insert(tk.END, item)
                                    lbframe.list.itemconfigure(tk.END, foreground="black")
                            errMessage = False
                            for item in newList:
                                if str(item).upper().startswith(str(string).upper()):
                                    lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                                    # lbframe.list.grid(row=1,column=1,sticky=tk.NSEW)
                                elif not lbframe.list.get(0):
                                    errMessage = True
                                    lbframe.place_forget()
                                    
                            if errMessage:
                                if not len(cxDict) and not str(boxList["C_Quote Yes/No"][0][index][0].get()).upper().startswith("N"):
                                    messagebox.showerror(title="Wrong Value",message="Please enter value from list only!")
                                    ent.delete(0, tk.END)
                                    ent.focus()
                                elif len(cxDict):
                                    messagebox.showerror(title="Wrong Value",message="Please enter value from list only!")
                                    ent.delete(0, tk.END)
                                    ent.focus()

                        else:
                            lbframe.place_forget()
                            if string!='' and check:
                                messagebox.showerror(title="Wrong Value",message="Please enter value from list only!222")
                                ent.focus()
            except Exception as e:
                raise e
                        
                    # lst.grid_remove()

        var.trace('w', get_input)

        return (ent,var)
    except Exception as ex:
        raise ex
    # ent.bind("<Button-1>", get_input)
    # ent.bind('<Down>', list_input)
    # ent.bind('<Return>', list_hide)

    # lst.bind('<Up>', list_up)
    # lst.bind('<Return>', get_selection)
    
# def mylistBox(ent, item_list, lst, var):
#     var.trace('w', get_input(ent, item_list, lst, var))
    