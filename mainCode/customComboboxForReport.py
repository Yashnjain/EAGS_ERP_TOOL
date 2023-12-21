import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.tix import ButtonBox
import Tools
import pandas as pd


def myCombobox(df,root,frame,row,column,width=10,list_bd = 0,foreground='blue', background='white',sticky = tk.EW,item_list=[],boxList={},cxDict={},salesDict={},val = None, pt=None, bakerDf = [],is_on=None):
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
        ent = ttk.Entry(frame,textvariable=var, font=('Segoe UI', 10),foreground=foreground,background=background,width=10)
        lst = tk.Listbox(frame, bd=list_bd, background=background,width=width,height=20)

        lbframe = tk.Frame(root)
        lbframe.list_values = tk.StringVar(frame) 
        lbframe.list_values.set(item_list) 
        lbframe.list= tk.Listbox(lbframe, height=5, width= 20,
                                        listvariable=lbframe.list_values)
        lbframe.list.pack(side='left', fill="both", expand=True)
        # s = ttk.Scrollbar(lbframe, orient=tk.VERTICAL, command=lbframe.list.yview)
        # s.pack(side='right', fill = "y")
        # lbframe.list['yscrollcommand'] = s.set
        set_mousewheel(lbframe.list, onMouseWheel)
        # lbframe.list.bind("<MouseWheel>", onMouseWheel)
        ent.grid(row=row,column=column,sticky=tk.EW,padx=10,pady=10)



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
        # def addCostCalc(boxList,index):
        #     try:
        #         if boxList["E_Additional_Cost"][0][index][1].get() != '':
        #             if boxList["E_Additional_Cost"][0][index][1].get() != 'None':
        #                 addCost = float(boxList["E_Additional_Cost"][0][index][1].get())
        #         else:
        #             root.attributes('-topmost', True)
        #             messagebox.showerror(title="Wrong Value",message="Please enter value in Additional Cost",parent=root)
        #             root.attributes('-topmost', False)
        #             return
        #         if boxList["E_Selling Cost/UOM"][0][index][1].get() != '':
        #             if boxList["E_Selling Cost/UOM"][0][index][1].get() != 'None':
        #                 sellCost = float(boxList["E_Selling Cost/UOM"][0][index][1].get())
        #         else:
        #             root.attributes('-topmost', True)   
        #             messagebox.showerror(title="Wrong Value",message="Please enter value in Selling Cost/UOM",parent=root)
        #             root.attributes('-topmost', False)
        #             return
        #         finalPrice = round((addCost+sellCost),2)
        #         boxList["E_Final Price"][0][index][1].set(finalPrice)
        #     except Exception as ex:
        #         raise ex


        
        def list_hide(e=None):
            try:
                breakCheck = False
                # global get_input_checker
                global checker
                checker = False
                print("List box size is ")
                print(lbframe.list.size())
                if lbframe.list.size()!=0:
                    value = lbframe.list.get(0)
                    var.set(value)
                    #finding current key and index
                    if len(boxList):
                        key, index = keyFinder(boxList,(ent,var))
                        
                    # elif len(salesDict):
                    #     key, index = keyFinder(salesDict,(ent,var))
                    # else:
                    key, index = keyFinder(cxDict,(ent,var))
                    if key == "quoteno":
                        # if is_on is not None:
                            # if is_on.get()=='On':
                        # df["quote"] = df['cus_long_name'].astype(str) +" | "+ df["cus_id"]
                        # dataList = df[(df["cus_long_name"] == value)].values.tolist()[0]
                        dataList = df[(df["quoteno"] == value)].values.tolist()[0]
                        try:
                            cxDict["sales_person"][index][1].set(dataList[2])
                        except:
                            pass
                        print(dataList[1])
                        print(f"Index is {index}")
                        cxDict["preparedby"][index][1].set(dataList[1])
                        # cxDict["payment_term"] = dataList[2]

                        # cxDict["cus_address"][index][1].set(dataList[4])
                        # # cxDict["cus_address"] = dataList[3]

                        # cxDict["cus_email"][index][1].set(dataList[6])
                        # # cxDict["cus_email"] = dataList[5]

                        # cxDict["cus_phone"][index][1].set(dataList[5])

                        # cxDict["cus_city_zip"] = dataList[7]
                        # val.lift()
                        # val.focus_set()
                        breakCheck = True
                    

               
                
                

                checker = True
                

                    
                lbframe.list.delete(0, tk.END)
                lbframe.place_forget()
                if breakCheck:# or e.keysym=="Tab":
                    return "break"
            except Exception as ex:
                if len(bakerDf):
                    if bakerDf.iloc[0,0] is None:
                        root.attributes('-topmost', True)
                        messagebox.showerror(title="Wrong Value",message="Please paste Baker input data first",parent=root)
                        root.attributes('-topmost', False)
                    else:
                        raise ex
                else:
                    raise ex

        def list_input(e):
            try:
                print(e.type)
                print(e.num)
                print("abc")
                listCheck = True
                if ent.get()=='':
                    
                    key, index = keyFinder(cxDict,(ent,var))
                    # df["cx_name_id"] = df['cus_long_name'].astype(str) +" | "+ df["cus_id"]
                    newList = list(set(df["quoteno"].values.tolist()[:1000]))

                    lbframe.list.delete(0, tk.END)
                    if listCheck and len(newList):
                        for item in newList:
                            lbframe.list.insert(tk.END, item)
                            lbframe.list.itemconfigure(tk.END, foreground="black")
                        if not newList[0]=='' or (newList[0]=='' and len(newList)!=1):  
                           
                            lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                            lbframe.list.focus()
                            lbframe.list.select_set(0)

                
                elif len(cxDict):
                    key, index = keyFinder(cxDict,(ent,var))
                    if key == "quoteno":
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
                if len(lbframe.list.curselection()):
                    if not lbframe.list.curselection()[0]:
                        
                        ent.focus()
                        list_hide()
            except Exception as e:
                raise e
                    


        def get_selection(e):
            try:
                if len(lbframe.list.curselection()):
                    value = lbframe.list.get(lbframe.list.curselection())
                    var.set(value)
                    list_hide(e)
                    if not salesDict and not cxDict:
                        ent.focus()
                        ent.icursor(tk.END)
                else:
                    pass
            except Exception as ex:
                raise ex
        

        # def handle_left_click(e):
        #     try:
        #         rowclicked_single = pt.get_row_clicked(e)
        #         print(f"Row clicked is {rowclicked_single+1}")
        #         if len(boxList) and rowclicked_single < len(pt.model.df):
        #             if list(pt.model.df.columns) == ['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in','date_last_receipt','age', 'heat_number', 'lot_serial_number']:
        #                 if list(pt.model.df.iloc[0]) != [None, None, None, None, None, None, None, None, None]:
        #                     # pt=pt.getCurrentTable()
        #                     varname = focused_entry.cget("textvariable")
        #                     focused_var = focused_entry.getvar(varname)
        #                     key, index = keyFinder2(boxList,(focused_entry,varname))
        #                     if key ==  None and index == None:
        #                         pass
        #                     else:
        #                         pt.model.df.reset_index(inplace = True,drop = True)
        #                         print(key, index)
        #                         boxList['Lot_Serial_Number'][0][index] = (pt.model.df['lot_serial_number'][rowclicked_single], None)
        #                         boxList['E_COST'][0][index][1].set(float(pt.model.df['onhand_dollars_per_pounds'][rowclicked_single]))
        #                         if not len(bakerDf):
        #                             boxList['E_freightIncured'][0][index][1].set(0)
        #                             boxList['E_freightCharged'][0][index][1].set(0)
        #                             boxList['E_Margin_Freight'][0][index][1].set(0)
        #                         boxList['E_Additional_Cost'][0][index][1].set(0)
        #                 else:
        #                     pass
        #             else:
        #                 pass
        #         pt.setSelectedRow(rowclicked_single)
        #         pt.redraw()
        #     except Exception as ex:
        #         raise ex
        def remember_focus(event):
            global focused_entry
            focused_entry = event.widget
            



        ent.bind("<1>", list_input)
        ent.bind('<Down>', list_input)
        ent.bind('<Return>', list_hide)
        ent.bind('<Tab>',list_hide)
        ent.bind('<FocusIn>',remember_focus)
        # ent.bind("<Leave>",list_up)
        # ent.bind('<Escape>', list_hide)

        lbframe.list.bind("<Enter>", list_input)
        lbframe.list.bind('<Up>', list_up)
        lbframe.list.bind('<Return>', get_selection)
        lbframe.list.bind('<Double-Button-1>', get_selection)
        lbframe.list.bind('<Tab>',get_selection)
        lbframe.list.bind('<Escape>',list_hide)
        lbframe.list.bind("<Leave>", list_hide)

        # if pt is not None:
        
        #     pt.bind('<Button-1>',handle_left_click)
        
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

        def keyFinder2(dict, tupleValue):
            try:
                for key, value in dict.items():
                    for index in range(len(value[0])):
                        
                        if value[0][index][0] == tupleValue[0] and value[0][index][1]._name == tupleValue[1]:
                            return key,index
                else:
                    return None,None
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
                elif key== 'E_Location' or key == 'searchLocation':
                    if key == 'searchLocation':
                        newList = list(df["site"].unique())
                    if boxList['C_Quote Yes/No'][0][index][0].get() != '':
                        if boxList['C_Quote Yes/No'][0][index][0].get() != 'None':
                            newList = list(df["site"].unique())
                elif key == 'E_Type':
                    if len(bakerDf) and key == 'E_Type':
                        newList = ["HT","HB", "HM"]
                        
                    else:
                        newList = ["THF","BR", "TUI", "HR"]
                    # #filter df based on e_location and make unique column of e_type
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())]
                    # newList = list(new_df["material_type"].unique())

                elif key == 'E_Grade':
                    
                    #filter df based on e_location and make unique column of e_type
                    
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)]#boxList['E_Type'][0][index][0].get())]
                    # if key == 'searchGrade':
                    #     new_df = df[(df["site"] == boxList['searchLocation'][0][index][0].get())]#boxList['E_Type'][0][index][0].get())]
                    # else:
                    new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())]#boxList['E_Type'][0][index][0].get())]
                    
                    newList = list(new_df["grade"].unique())
                elif key == 'E_Yield':
                    #filter df based on e_location and make unique column of e_type
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)#boxList['E_Type'][0][index][0].get())
                    # if key == 'searchYield':
                    #     new_df = df[(df["site"] == boxList['searchLocation'][0][index][0].get())#boxList['E_Type'][0][index][0].get())
                    #     & (df["grade"]==boxList['searchGrade'][0][0][0].get())]
                    # else:
                    new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())#boxList['E_Type'][0][index][0].get())
                    & (df["grade"]==boxList['E_Grade'][0][index][0].get())]
                    newList = list(new_df["heat_condition"].unique())
                elif key == 'E_OD1':
                    #filter df based on e_location and make unique column of e_type
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)#==boxList['E_Type'][0][index][0].get())
                    new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())#==boxList['E_Type'][0][index][0].get())
                    & (df["grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())]
                    newList = list(new_df["od_in"].unique())
                elif key == 'E_OD2':
                    #filter df based on e_location and make unique column of e_type
                    # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)#==boxList['E_Type'][0][index][0].get())
                    new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())#==boxList['E_Type'][0][index][0].get())
                    & (df["grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())]
                    newList = list(new_df["od_in"].unique())
                elif key == 'E_ID1':
                    try:
                        #filter df based on e_location and make unique column of e_type
                        # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)#boxList['E_Type'][0][index][0].get())
                        new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())#boxList['E_Type'][0][index][0].get())
                        & (df["grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                        & (df["od_in"]==float(boxList['E_OD1'][0][index][0].get()))]
                        newList = list(new_df["od_in_2"].unique())
                
                    except Exception as ex:
                        newList = []
                elif key == 'E_ID2':
                    try:
                        #filter df based on e_location and make unique column of e_type
                        # new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==boxList['E_Type'][0][index][0].get())
                        new_df = df[(df["site"] == boxList['E_Location'][0][index][0].get())
                        & (df["grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                        & (df["od_in"]==float(boxList['E_OD2'][0][index][0].get()))]
                        newList = list(new_df["od_in_2"].unique())
                
                    except Exception as ex:
                        newList = []

                elif key == "quoteno":
                    #filter df based on e_location and make unique column of e_type
                    new_df = df[(df["quoteno"] == boxList["quoteno"][0][index][0].get())] #
                    newList = list(set(new_df.values.tolist()))
                try:
                    if key != 'C_Quote Yes/No':
                        newList.sort()
                except:
                    pass
                return newList
            except Exception as e:
                raise e


        ###########Function for checking value entered via double click from table or not################
        def toplevels(cur_root):
            for k, v in cur_root.children.items():
                if isinstance(v, tk.Toplevel):
                    print('Toplevel:', k, v.title())
                    if v.title() == "Range Search Table":
                        return True
            return False
                
                

        def get_input(*args):
            try:
                if checker:
                    rangeSearcherValue = toplevels(root)

                    newList = item_list
                    check = True
                    # print(boxList)
                   
                    if len(cxDict):
                        key, index = keyFinder(cxDict,(ent,var))
                        # df["cx_name_id"] = df['cus_long_name'].astype(str) +" | "+ df["cus_id"]
                        newList = list(set(df["quoteno"].values.tolist()))
                        # if key == "cus_long_name":
                        #     if is_on is not None:
                        #         if is_on.get()=='Off':
                        #             newList = []
                    
                    # lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                    
                    lbframe.list.delete(0, tk.END)
                    if len(newList)==1:
                        # if (key == 'E_ID1' or key == 'E_OD1') and (boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" or boxList['E_Type'][0][index][0].get()=="HM"):
                        #     pass
                        # else:
                        var.set(newList[0])
                    else:
                        string = var.get()

                        if string and check and not rangeSearcherValue:
                            for item in newList:
                                if str(item).upper().startswith(str(string).upper()):
                                    lbframe.list.insert(tk.END, item)
                                    lbframe.list.itemconfigure(tk.END, foreground="black")
                            errMessage = False
                            for item in newList:
                                if str(item).upper().startswith(str(string).upper()):
                                    lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                                    # lbframe.list.grid(row=1,column=1,sticky=tk.NSEW)
                                elif not lbframe.list.get(0) and lbframe.list.get(0) != 0.0:
                                    errMessage = True
                                    lbframe.place_forget()
                                    
                            if errMessage:
                                if not len(salesDict) and not len(cxDict) and not str(boxList["C_Quote Yes/No"][0][index][0].get()).upper().startswith("N"):
                                    root.attributes('-topmost', True)
                                    messagebox.showerror(title="Wrong Value",message="Please enter value from list only!",parent=root)
                                    root.attributes('-topmost', False)
                                    ent.delete(0, tk.END)
                                    ent.focus()
                                elif len(salesDict):
                                    root.attributes('-topmost', True)
                                    messagebox.showerror(title="Wrong Value",message="Please enter value from list only!",parent=root)
                                    root.attributes('-topmost', False)
                                    ent.delete(0, tk.END)
                                    ent.focus()
                                elif len(cxDict):
                                    root.attributes('-topmost', True)
                                    messagebox.showerror(title="Wrong Value",message="Please enter value from list only!",parent=root)
                                    root.attributes('-topmost', False)
                                    ent.delete(0, tk.END)
                                    ent.focus()

                        else:
                            lbframe.place_forget()
                            if string!='' and check and string!='None' and not rangeSearcherValue:
                                root.attributes('-topmost', True)
                                messagebox.showerror(title="Wrong Value",message="Please enter value from list only!",parent=root)
                                root.attributes('-topmost', False)
                                ent.focus()
            except Exception as e:
                if ent.winfo_exists()==0:
                    root.attributes('-topmost', True)
                    messagebox.showerror(title="Error",message="Current sub window closed, please reopen it",parent=root)
                    root.attributes('-topmost', False)
                else:
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
    