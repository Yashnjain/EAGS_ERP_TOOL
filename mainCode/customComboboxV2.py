import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.tix import ButtonBox
import Tools
import pandas as pd

def formulaCalc(boxList, index, root):
    try:
        if (boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" or boxList['E_Type'][0][index][0].get()=="HM") and boxList['C_Quote Yes/No'][0][index][0].get() != "Other" :
            if boxList["E_OD2"][0][index][1]!='' and boxList["E_ID2"][0][index][1]!='':
                if boxList["E_OD2"][0][index][1]!='None' and boxList["E_ID2"][0][index][1]!='None':
                    e_od = float(boxList["E_OD2"][0][index][1])
                    e_id = float(boxList["E_ID2"][0][index][1])
            else:
                root.attributes('-topmost', True)   
                messagebox.showerror(title="Wrong Value",message="Please enter value To be Calculated OD and ID",parent=root)
                root.attributes('-topmost', False)
                return
        else:
            if boxList["E_OD1"][0][index][1].get() != '' and boxList["E_ID1"][0][index][1].get() != '':
                if boxList["E_OD1"][0][index][1].get() != 'None' and boxList["E_ID1"][0][index][1].get() != 'None':
                    e_od = float(boxList["E_OD1"][0][index][1].get())
                    e_id = float(boxList["E_ID1"][0][index][1].get())
            else:
                root.attributes('-topmost', True)
                messagebox.showerror(title="Wrong Value",message="Please enter value in OD and ID",parent=root)
                root.attributes('-topmost', False)
                return
        uom = boxList["E_UOM"][0][index][1].get()
        if boxList["E_Length"][0][index][1].get() != '':
            if boxList["E_Length"][0][index][1].get() != 'None':
                e_length = float(boxList["E_Length"][0][index][1].get())#int(boxList["E_Length"][0][index][1].get())
        else:
            root.attributes('-topmost', True)   
            messagebox.showerror(title="Wrong Value",message="Please enter value in Length Entry Box", parent=root)
            root.attributes('-topmost', False)
            return
        if boxList["E_Selling Cost/LBS"][0][index][1].get() != '':
            if boxList["E_Selling Cost/LBS"][0][index][1].get() != 'None':
                sellCostLBS = float(boxList["E_Selling Cost/LBS"][0][index][1].get())
        else:
            root.attributes('-topmost', True)    
            messagebox.showerror(title="Wrong Value",message="Please enter value in sellCost/LBS", parent=root)
            root.attributes('-topmost', False)
            return
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
            elif uom == "Inch":
                # Selling cost/UOM = "SellingCost/LBS" * mid_formula (rounded upto 2 decimal places)
                sellCostUOM = round((sellCostLBS * mid_formula),2)
            elif uom =="Meter":
                sellCostUOM = round(((e_od-wt) * wt * 0.0544),2)
            #For Foot
            else:
                sellCostUOM = round((sellCostLBS * mid_formula),2)
                sellCostUOM = sellCostUOM *12
                
        elif boxList["E_Type"][0][index][1].get() == "BR" or boxList["E_Type"][0][index][1].get() == "HR" or boxList["E_Type"][0][index][1].get() == "HB" or boxList["E_Type"][0][index][1].get() == "HM":
            # BR
            mid_formula = (e_od*e_od*2.71)/12
            # mid_formula = ((e_od-wt)*wt*10.68)/12
            if uom == "Each":
                # For Each: 
                # Selling cost/UOM ="SellingCost/LBS" * mid_formula * Length 
                sellCostUOM = round((sellCostLBS * mid_formula * e_length),2)
            elif uom == "Inch":
                # For Inch: 
                # Selling cost/UOM ="SellingCost/LBS" * mid_formula	
                sellCostUOM = round((sellCostLBS * mid_formula),2)
            elif uom =="Meter":
                sellCostUOM = round(((e_od * e_od * 2.71) / 196.64),2)
            #For Foot
            else:
                sellCostUOM = round((sellCostLBS * mid_formula),2)
                sellCostUOM = sellCostUOM *12
        else:
            sellCostUOM = 0
        return sellCostUOM
    except Exception as ex:
        raise ex


def myCombobox(df,root,frame,row,column,width=10,list_bd = 0,foreground='blue', background='white',sticky = tk.EW,item_list=[],boxList={},cxDict={},salesDict={},val = None, pt=None, entpady=0, bakerDf = [],is_on=None):
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
                if boxList["E_Additional_Cost"][0][index][1].get() != '':
                    if boxList["E_Additional_Cost"][0][index][1].get() != 'None':
                        addCost = float(boxList["E_Additional_Cost"][0][index][1].get())
                else:
                    root.attributes('-topmost', True)
                    messagebox.showerror(title="Wrong Value",message="Please enter value in Additional Cost",parent=root)
                    root.attributes('-topmost', False)
                    return
                if boxList["E_Selling Cost/UOM"][0][index][1].get() != '':
                    if boxList["E_Selling Cost/UOM"][0][index][1].get() != 'None':
                        sellCost = float(boxList["E_Selling Cost/UOM"][0][index][1].get())
                else:
                    root.attributes('-topmost', True)   
                    messagebox.showerror(title="Wrong Value",message="Please enter value in Selling Cost/UOM",parent=root)
                    root.attributes('-topmost', False)
                    return
                finalPrice = round((addCost+sellCost),2)
                boxList["E_Final Price"][0][index][1].set(finalPrice)
            except Exception as ex:
                raise ex


        
        def list_hide(e=None):
            try:
                breakCheck = False
                # global get_input_checker
                global checker
                checker = False
                print("List box size is ")
                print(lbframe.list.size())
                if lbframe.list.size()==1:
                    value = lbframe.list.get(0)
                    var.set(value)
                    #finding current key and index
                    if len(boxList):
                        key, index = keyFinder(boxList,(ent,var))
                        
                    elif len(salesDict):
                        key, index = keyFinder(salesDict,(ent,var))
                    else:
                        key, index = keyFinder(cxDict,(ent,var))
                    if key == "cus_long_name":
                        if is_on is not None:
                            if is_on.get()=='On':
                                df["cx_name_id"] = df['cus_long_name'].astype(str) +" | "+ df["cus_id"]
                                # dataList = df[(df["cus_long_name"] == value)].values.tolist()[0]
                                dataList = df[(df["cx_name_id"] == value)].values.tolist()[0]
                                cxDict["payment_term"][index][1].set(dataList[3])
                                # cxDict["payment_term"] = dataList[2]

                                cxDict["cus_address"][index][1].set(dataList[4])
                                # cxDict["cus_address"] = dataList[3]

                                cxDict["cus_email"][index][1].set(dataList[6])
                                # cxDict["cus_email"] = dataList[5]

                                cxDict["cus_phone"][index][1].set(dataList[5])

                                cxDict["cus_city_zip"] = dataList[7]
                                # val.lift()
                                val.focus_set()
                        breakCheck = True
                    elif key == "sales_person":
                        df["s_person_name_id"] = df['sales_person'] +" | "+ df["code_no"].astype(str) +" | "+ df["location"]
                        val.focus_set()
                        breakCheck = True

                    elif key=='E_UOM' and boxList['C_Quote Yes/No'][0][index][0].get() != "No":
                        sellCostUOM = formulaCalc(boxList, index, root)
                        boxList["E_Selling Cost/UOM"][0][index][1].set(sellCostUOM)
                        boxList["E_Additional_Cost"][0][index][0].focus()
                    elif key == "E_Additional_Cost":
                        # addCost = float(boxList["E_Additional_Cost"][0][index][1].get())
                        # sellCost = float*(boxList["E_Selling Cost/UOM"][0][index][1].get())
                        if boxList["E_Additional_Cost"][0][index][1].get() != '':
                            if boxList["E_Additional_Cost"][0][index][1].get() != 'None':
                                addCost = float(boxList["E_Additional_Cost"][0][index][1].get())
                        else:
                            root.attributes('-topmost', True)
                            messagebox.showerror(title="Wrong Value",message="Please enter value in Additional Cost",parent=root)
                            root.attributes('-topmost', False)
                        if boxList["E_Selling Cost/UOM"][0][index][1].get() != '':
                            if boxList["E_Selling Cost/UOM"][0][index][1].get() != 'None':
                                sellCost = float(boxList["E_Selling Cost/UOM"][0][index][1].get())
                        else:
                            root.attributes('-topmost', True)
                            messagebox.showerror(title="Wrong Value",message="Please enter value in Selling Cost/UOM",parent=root)
                            root.attributes('-topmost', False)
                        finalPrice = round((addCost+sellCost),2)
                    
                        boxList["E_Final Price"][0][index][0].set(finalPrice)

                    
                    # elif key == "E_Type":
                    #     boxList['E_Length'][0][index][0].focus()E_OD2
                    #     breakCheck = True

                        
                    elif value != "Other" and value != "Yes" and value != "No"  and key != "E_UOM" and boxList['C_Quote Yes/No'][0][index][0].get() != "Other" and boxList['C_Quote Yes/No'][0][index][0].get() != "No": #and key!='E_Location'
                        current_key = key
                        while True:
                            if current_key=="searchLocation":
                                break
                            next_key = nextKey(current_key, index)
                            if next_key == "E_Length":
                                if boxList['E_ID1'][0][index][0].get() != '' and boxList['E_OD1'][0][index][0].get() != '' \
                                    and boxList['E_ID1'][0][index][0].get() != 'None' and boxList['E_OD1'][0][index][0].get() != 'None':
                                    #Adding Baker Type Handler
                                    if boxList['E_Type'][0][index][0].get() == "HT":
                                        e_type_var = "THF"
                                    elif boxList['E_Type'][0][index][0].get() == "HB":
                                        e_type_var = "BR"
                                    else:
                                        e_type_var = boxList['E_Type'][0][index][0].get()
                                    if (boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" \
                                         or boxList['E_Type'][0][index][0].get()=="HM") and (boxList['E_ID2'][0][index][0].get() != '' \
                                             and boxList['E_OD2'][0][index][0].get() != ''):
                                        # newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)
                                        if boxList['E_ID2'][0][index][0].get() != 'None' and boxList['E_OD2'][0][index][0].get() != 'None':
                                            newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())
                                                    & (df["grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                                                    & (df["od_in"]==float(boxList['E_OD2'][0][index][0].get())) & (df["od_in_2"]==float(boxList['E_ID2'][0][index][0].get()))]
                                        
                                    else:
                                        if boxList['E_OD1'][0][index][0].get() != '' and boxList['E_ID1'][0][index][0].get() != '':
                                            if boxList['E_OD1'][0][index][0].get() != 'None' and boxList['E_ID1'][0][index][0].get() != 'None':
                                            # newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==e_type_var)
                                                newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())
                                                        & (df["grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                                                        & (df["od_in"]==float(boxList['E_OD1'][0][index][0].get())) & (df["od_in_2"]==float(boxList['E_ID1'][0][index][0].get()))]
                                        else:
                                            root.attributes('-topmost', True)
                                            messagebox.showerror(title="Wrong Value",message="Please enter value in OD and ID",parent=root)
                                            root.attributes('-topmost', False)
                                            break
                                    # newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds','reserved_pieces', 'reserved_length_in', 'available_pieces', 'available_length_in']]
                                    newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in','date_last_receipt','age', 'heat_number', 'lot_serial_number']]
                                    newDf['date_last_receipt'] = pd.to_datetime(newDf['date_last_receipt'])
                                    newDf['date_last_receipt'] = newDf['date_last_receipt'].dt.date
                                    newDf = newDf[newDf['available_pieces']>0]
                                    newDf = newDf.sort_values('age', ascending=False).sort_values('date_last_receipt', ascending=True)
                                    boxList['E_Length'][0][index][0].focus()
                                    #Resetting Index
                                    newDf.reset_index(inplace=True, drop=True)
                                    if pt is not None:
                                        pt.model.df = newDf
                                        pt.redraw()
                                        pt.rowheader.bind('<Button-1>',handle_left_click)
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
                                boxList['E_OD1'][0][index][0].configure(state='normal')
                                boxList['E_ID1'][0][index][0].configure(state='normal')
                                if boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" or boxList['E_Type'][0][index][0].get()=="HM":
                                    newList_id = filterList(boxList, 'E_ID2', index, df)
                                    try:
                                        boxList["E_OD2"][0][index][0].get()
                                        if len(newList_id)==1:
                                            boxList["E_ID2"][0][index][1].set(newList_id[0])
                                    except:
                                        try:
                                            od1,id1,od2,id2 = Tools.specialCase(root, boxList, pt, df, index, item_list=new_list, bakerDf=bakerDf,cxDict=cxDict)

                                            if od1 is not None and id1 is not None:
                                                boxList['E_OD1'][0][index][1].set(od1)
                                                boxList['E_ID1'][0][index][1].set(id1)
                                                boxList['E_OD1'][0][index][0].configure(state='disabled')
                                                boxList['E_ID1'][0][index][0].configure(state='disabled')

                                        except Exception as ex:
                                            raise ex
                                    break
                                # new_list2 = filterList(boxList,'E_ID',index,df)
                                #function call to pop up windows for 2 different od and id's and let it fill od id in main page writern variables for calculation
                                elif len(new_list)==1:
                                    boxList[next_key][0][index][1].set(new_list[0])
                                elif boxList['E_ID1'][0][index][0].get() == '' or boxList['E_OD1'][0][index][0].get() == '':
                                    if not len(bakerDf):
                                        if boxList[cx_eags[next_key]][0][index][0].get() != '':
                                            if boxList[cx_eags[next_key]][0][index][0].get() != 'None':
                                                if float(boxList[cx_eags[next_key]][0][index][0].get()) in list(map(lambda x: float(x),new_list)):
                                                    if next_key != 'E_OD2' and next_key != 'E_ID2':
                                                        boxList[next_key][0][index][1].set(float(boxList[cx_eags[next_key]][0][index][0].get()))
                                        else:
                                            root.attributes('-topmost', True)
                                            messagebox.showerror(title="CX Requiremnet not Filled",message="Please fill customer requirements first",parent=root)
                                            root.attributes('-topmost', False)
                                            for keys in cx_eags.keys():
                                                if not(keys == 'E_OD2' or keys == 'E_ID2'):
                                                    boxList[keys][0][index][1].set('')
                                            boxList[key][0][index][0].focus()
                                            breakCheck = True
                                            break
                                    else:#removing .get() from customer side 
                                        if boxList[cx_eags[next_key]][0][index][0] is not None and boxList[cx_eags[next_key]][0][index][0] != '' and boxList[cx_eags[next_key]][0][index][0] != 'None':
                                            if float(boxList[cx_eags[next_key]][0][index][0]) in list(map(lambda x: float(x),new_list)):
                                                if next_key != 'E_OD2' and next_key != 'E_ID2':
                                                    boxList[next_key][0][index][1].set(float(boxList[cx_eags[next_key]][0][index][0]))
                                        else:
                                            root.attributes('-topmost', True)
                                            messagebox.showerror(title="CX Requiremnet not Filled",message="Please fill customer requirements first",parent=root)
                                            root.attributes('-topmost', False)
                                            for keys in cx_eags.keys():
                                                if not(keys == 'E_OD2' or keys == 'E_ID2'):
                                                    boxList[keys][0][index][1].set('')
                                            boxList[key][0][index][0].focus()
                                            breakCheck = True
                                            break
                                
                                else:
                                    break
                            elif not len(bakerDf) and next_key == 'E_Type':#2n case will be handled len(bakerDf) and str(boxList[cx_eags[next_key]][0][index][0]).upper() in list(map(lambda x: str(x).upper(),new_list))
                                if boxList[next_key][0][index][1].get()=='':
                                    boxList[next_key][0][index][1].set("")
                                    break
        
                            elif next_key == 'E_Spec' and len(bakerDf):
                                boxList[next_key][0][index][1].set("")
                            
                            elif not len(bakerDf) and next_key not in ['E_Spec', 'Lot_Serial_Number', 'searchLocation'] and \
                                str(boxList[cx_eags[next_key]][0][index][0].get()).upper() in list(map(lambda x: str(x).upper(),new_list)):
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
                            
                            elif next_key != 'E_Spec' and not len(bakerDf) and (str(boxList[cx_eags[next_key]][0][index][0].get()) == '' or str(boxList[cx_eags[next_key]][0][index][0].get()) is None):
                                if next_key != 'E_Yield':
                                    root.attributes('-topmost', True)
                                    messagebox.showerror(title="Customer Requiremnet not Filled",message=f"Please fill customer {next_key} requirements first",parent=root)
                                    root.attributes('-topmost', False)
                                    if not(key == 'E_OD2' or key == 'E_ID2'):
                                        for keys in cx_eags.keys():
                                            if not(keys == 'E_OD2' or keys == 'E_ID2'):
                                                boxList[keys][0][index][1].set('')
                                        boxList[key][0][index][0].focus()
                                        breakCheck = True
                                        break
                            elif next_key != 'E_Type' and next_key != 'E_Spec' and len(bakerDf) and (str(boxList[cx_eags[next_key]][0][index][0]) == '' or str(boxList[cx_eags[next_key]][0][index][0]) is None):
                                root.attributes('-topmost', True)
                                messagebox.showerror(title="Customer Requiremnet not Filled",message=f"Please fill customer {next_key} requirements first",parent=root)
                                root.attributes('-topmost', False)
                                if not(key == 'E_OD2' or key == 'E_ID2'):
                                    for keys in cx_eags.keys():
                                        if not(keys == 'E_OD2' or keys == 'E_ID2'):
                                            boxList[keys][0][index][1].set('')
                                    boxList[key][0][index][0].focus()
                                    breakCheck = True
                                    break
                            else:
                                if next_key != 'E_OD2' and next_key != 'E_ID2' and next_key != 'Lot_Serial_Number':
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
                                # if newKey != 'E_OD2' and newKey != 'E_ID2' and newKey != 'Lot_Serial_Number'and newKey != 'searchLocation':
                                if newKey != 'E_OD2' and newKey != 'E_ID2' and newKey != 'E_Spec' and newKey != 'Lot_Serial_Number' and not len(bakerDf) and newKey != 'searchLocation':
                                    # if newKey=='E_Spec' and not len(bakerDf):
                                    #     pass
                                    # else:
                                    #     if newKey != 'E_freightIncured' and newKey != 'E_freightCharged' and newKey != 'E_Margin_Freight':
                                    boxList[newKey][0][index][1].set("")
                                    boxList[newKey][0][index][0].configure(state='normal')
                                elif newKey != 'E_OD2' and newKey != 'E_ID2' and newKey != 'Lot_Serial_Number' and newKey != 'searchLocation' and len(bakerDf):
                                    if newKey != 'E_freightIncured' and newKey != 'E_freightCharged' and newKey != 'E_Margin_Freight':
                                        boxList[newKey][0][index][1].set("")
                                        boxList[newKey][0][index][0].configure(state='normal')
                            else:
                                if newKey != 'E_OD2' and newKey != 'E_ID2' and newKey != 'E_Spec' and newKey != 'Lot_Serial_Number' and newKey != 'searchLocation' and newKey !='E_Location' and not len(bakerDf):
                                    boxList[newKey][0][index][1].set("NA")
                                    boxList[newKey][0][index][0].configure(state='disabled')
                                
                                elif newKey != 'E_OD2' and newKey != 'E_ID2' and newKey != 'Lot_Serial_Number' and newKey != 'searchLocation' and newKey !='E_Location' and len(bakerDf):
                                    if newKey != 'E_freightIncured' and newKey != 'E_freightCharged' and newKey != 'E_Margin_Freight':
                                        boxList[newKey][0][index][1].set("NA")
                                        boxList[newKey][0][index][0].configure(state='disabled')
                            # if value == "No":
                            #     print("no")
                            #     newKey == 'E_OD2' and newKey == 'E_ID2' and newKey == 'Lot_Serial_Number' and newKey == 'searchLocation' and newKey !='E_Location' and len(bakerDf)
                            #     boxList[newKey][0][index][1].set("NA")
                            #     boxList[newKey][0][index][0].configure(state='disabled')
                                

                            
                        
                
                
                else:
                    if len(boxList):
                        cx_eags = {'E_Grade':'C_Grade','E_Yield':'C_Yield','E_OD1':'C_OD','E_ID1':'C_ID','E_OD2':'C_OD','E_ID2':'C_ID'}
                        key, index = keyFinder(boxList,(ent,var))
                        if key=='E_Additional_Cost':
                            if boxList["E_Additional_Cost"][0][index][1].get()!='' and boxList["E_Selling Cost/UOM"][0][index][1].get() != '':
                                if boxList["E_Additional_Cost"][0][index][1].get()!='None' and boxList["E_Selling Cost/UOM"][0][index][1].get() != 'None':
                                    addCostCalc(boxList,index)
                                    # addCost = float(boxList["E_Additional_Cost"][0][index][1].get())
                                    # sellCost = float(boxList["E_Selling Cost/UOM"][0][index][1].get())
                                    # finalPrice = addCost+sellCost
                                    # boxList["E_Final Price"][0][index][1].set(finalPrice)

                        elif key=='E_Selling Cost/UOM':
                            if boxList["E_Additional_Cost"][0][index][0].get() == '0':
                                if boxList["E_Selling Cost/UOM"][0][index][1].get() != '' and boxList["E_Selling Cost/UOM"][0][index][1].get() != 'None':
                                    boxList["E_Final Price"][0][index][1].set(float(boxList["E_Selling Cost/UOM"][0][index][1].get()))
                                    boxList["E_Additional_Cost"][0][index][0].focus()
                                    breakCheck = True
                                else:
                                    root.attributes('-topmost', True)
                                    messagebox.showerror(title="Wrong Value",message="Please enter value in Selling Cost/UOM",parent=root)
                                    root.attributes('-topmost', False)
                                    return
                        
                        elif key=='E_LeadTime' and not len(bakerDf):

                            boxList["E_freightIncured"][0][index][0].focus()
                            breakCheck = True

                        #Adding condition for margin
                        elif key=='E_Qty':
                            if boxList["E_COST"][0][index][0].get()=='':
                                root.attributes('-topmost', True)
                                messagebox.showerror(title="COST not Selected",message="Please select cost price(on_dollars_per_pound) row from table present below",parent=root)
                                root.attributes('-topmost', False)
                                boxList["E_Qty"][0][index][0].focus()
                                breakCheck = True
                                
                            else:

                                boxList["E_Selling Cost/LBS"][0][index][0].focus()
                                breakCheck = True

                        #Adding condition for margin
                        elif key=='E_Selling Cost/LBS':
                            if boxList["E_Selling Cost/LBS"][0][index][0].get() == '':
                                root.attributes('-topmost', True)
                                messagebox.showerror(title="Sale Price Blank",message="Please enter value in Selling Cost/LBS box.",parent=root)
                                root.attributes('-topmost', False)
                                boxList["E_Selling Cost/LBS"][0][index][0].focus()
                                breakCheck = True

                            elif boxList["E_COST"][0][index][0].get() == '':
                                root.attributes('-topmost', True)
                                messagebox.showerror(title="COST not Selected",message="Please select cost price(on_dollars_per_pound) row from table present below",parent=root)
                                root.attributes('-topmost', False)
                                boxList["E_COST"][0][index][0].focus()
                                breakCheck = True
                            else:
                                #Calculating Margin forper LBS value of Raw Material
                                if boxList["E_Selling Cost/LBS"][0][index][0].get() != '' and boxList["E_COST"][0][index][0].get() != '':
                                    if boxList["E_Selling Cost/LBS"][0][index][0].get() != 'None' and boxList["E_COST"][0][index][0].get() != 'None':
                                        salePrice = float(boxList["E_Selling Cost/LBS"][0][index][0].get())
                                        costPrice = float(boxList["E_COST"][0][index][0].get())
                                        if salePrice == 0.0 or costPrice == 0.0:
                                            root.attributes('-topmost', True)
                                            messagebox.showerror(title="Wrong Value",message="Selling Cost/LBS or COST is 0, please check and retry", parent=root)
                                            root.attributes('-topmost', False)
                                            return
                                        margin_per_lbs = round(((salePrice - costPrice)/salePrice) * 100, 2)
                                        boxList["E_MarginLBS"][0][index][1].set(margin_per_lbs)
                                        boxList["E_UOM"][0][index][0].focus()
                                        breakCheck = True
                                else:
                                    root.attributes('-topmost', True)
                                    messagebox.showerror(title="Wrong Value",message="Selling Cost/LBS or COST is blank, please fill their respective boxes",parent=root)
                                    root.attributes('-topmost', False)
                                    return
                        #Adding Condition for Freight Calculation
                        elif key=='E_freightCharged':
                            if boxList["E_freightIncured"][0][index][0].get() == '':
                                root.attributes('-topmost', True)
                                messagebox.showerror(title="Fright Price Blank",message="Please enter value in Freight Incured box.",parent=root)
                                root.attributes('-topmost', False)
                                boxList["E_freightIncured"][0][index][0].focus()
                                breakCheck = True

                            elif boxList["E_freightCharged"][0][index][0].get() == '':
                                root.attributes('-topmost', True)
                                messagebox.showerror(title="Fright Price Blank",message="Please enter value in Freight Charged box.",parent=root)
                                root.attributes('-topmost', False)
                                boxList["E_freightCharged"][0][index][0].focus()
                                breakCheck = True
                            
                            else:
                                if boxList["E_freightIncured"][0][index][0].get() != '' and boxList["E_freightCharged"][0][index][0].get() != '':
                                    if boxList["E_freightIncured"][0][index][0].get() != 'None' and boxList["E_freightCharged"][0][index][0].get() != 'None':
                                        #Calculating Freight Margin
                                        fcostPrice = float(boxList["E_freightIncured"][0][index][0].get())
                                        fsalePrice = float(boxList["E_freightCharged"][0][index][0].get())
                                        if fsalePrice == 0.0 or fcostPrice == 0.0:
                                            root.attributes('-topmost', True)
                                            messagebox.showerror(title="Wrong Value",message="Freight Charged or Freight Incured is 0, please check and retry",parent=root)
                                            root.attributes('-topmost', False)
                                            return
                                        if fcostPrice != 0.0 and fsalePrice != 0.0:
                                            margin_freight = round(((fsalePrice - fcostPrice)/fcostPrice) * 100, 2)
                                            boxList["E_Margin_Freight"][0][index][1].set(margin_freight)
                                        
                                            breakCheck = True
                                else:
                                    root.attributes('-topmost', True)
                                    messagebox.showerror(title="Wrong Value",message="FreightIncured or FreightCharged is blank, please fill their respective boxes",parent=root)
                                    root.attributes('-topmost', False)
                                    return

                        elif key in cx_eags and boxList["C_Quote Yes/No"][0][index][0].get() != "Other":
                            new_list = filterList(boxList,key,index,df)
                            if not len(bakerDf):
                                if boxList[cx_eags[key]][0][index][0].get() != '' and boxList[cx_eags[key]][0][index][0].get() != 'None':
                                    if key not in ['E_Grade', 'E_Yield']:
                                        if float(boxList[cx_eags[key]][0][index][0].get()) in list(map(lambda x: float(x),new_list)):
                                            if key != 'E_OD2' and key != 'E_ID2':
                                                boxList[key][0][index][1].set(float(boxList[cx_eags[key]][0][index][0].get()))
                                    else:
                                        if boxList[cx_eags[key]][0][index][0].get() in new_list:
                                            if key != 'E_OD2' and key != 'E_ID2':
                                                boxList[key][0][index][1].set(boxList[cx_eags[key]][0][index][0].get())
                                else:
                                    if key != 'E_Yield':
                                        root.attributes('-topmost', True)
                                        messagebox.showerror(title="CX Requiremnet not Filled",message="Please fill customer requirements first",parent=root)
                                        root.attributes('-topmost', False)
                                        for keys in cx_eags.keys():
                                            if not(keys == 'E_OD2' or keys == 'E_ID2'):
                                                boxList[keys][0][index][1].set('')
                                        boxList[key][0][index][0].focus()
                                        breakCheck = True
                                        # break
                            else:#removing .get() from customer side 
                                if boxList[cx_eags[key]][0][index][0] is not None and boxList[cx_eags[key]][0][index][0] != ''  and boxList[cx_eags[key]][0][index][0] != 'None' and key not  in ['E_Grade', 'E_Yield']:
                                    # if key not in ['E_Grade', 'E_Yield']:
                                    if float(boxList[cx_eags[key]][0][index][0]) in list(map(lambda x: float(x),new_list)):
                                        if key != 'E_OD2' and key != 'E_ID2':
                                            boxList[key][0][index][1].set(float(boxList[cx_eags[key]][0][index][0]))
                                    # else:
                                    #     if boxList[cx_eags[key]][0][index][0].get() in new_list:
                                    #         if key != 'E_OD2' and key != 'E_ID2':
                                    #             boxList[key][0][index][1].set(boxList[cx_eags[key]][0][index][0].get())
                                else:
                                     if key not in ['E_Grade', 'E_Yield']:
                                        root.attributes('-topmost', True)
                                        messagebox.showerror(title="CX Requiremnet not Filled",message="Please fill customer requirements first", parent=root)
                                        root.attributes('-topmost', False)
                                        for keys in cx_eags.keys():
                                            if not(keys == 'E_OD2' or keys == 'E_ID2'):
                                                boxList[keys][0][index][1].set('')
                                        boxList[key][0][index][0].focus()
                                        breakCheck = True
                                    # break
                                
                            


                                        
                    

                            
                            

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
                listCheck = True
                if ent.get()=='':
                    if len(boxList):
                        key, index = keyFinder(boxList,(ent,var))
                        if boxList['C_Quote Yes/No'][0][index][1].get() == "Other" and key != 'E_UOM' and key != 'E_Location' and key!='E_Type':
                            listCheck = False
                        elif key=='E_UOM':
                            newList = item_list
                        else:
                            newList = filterList(boxList,key,index,df)
                        
                    elif len(salesDict):
                        key, index = keyFinder(salesDict,(ent,var))
                        df["s_person_name_id"] = df['sales_person'] +" | "+ df["code_no"].astype(str) +" | "+ df["location"]
                        newList = df["s_person_name_id"].values.tolist() 

                    else:
                        key, index = keyFinder(cxDict,(ent,var))
                        df["cx_name_id"] = df['cus_long_name'].astype(str) +" | "+ df["cus_id"]
                        newList = df["cx_name_id"].values.tolist() 
                        if key == "cus_long_name":
                            if is_on is not None:
                                if is_on.get()=='Off':
                                    newList = []
                    lbframe.list.delete(0, tk.END)
                    if listCheck and len(newList):
                        for item in newList:
                            lbframe.list.insert(tk.END, item)
                            lbframe.list.itemconfigure(tk.END, foreground="black")
                        if not newList[0]=='' or (newList[0]=='' and len(newList)!=1):  
                            if (key=='E_OD1' or key =='E_ID1'):
                                if boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" or boxList['E_Type'][0][index][0].get()=="HM":
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
                    if len(lbframe.list_values.get()) != 2 and (boxList['C_Quote Yes/No'][0][index][0].get()=="Yes" or key in ['searchLocation']) and len(lbframe.list_values.get()) != 0:
                        lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                        lbframe.list.focus()
                        lbframe.list.select_set(0)
                    #Adding for Freight
                    elif key == 'E_freightIncured' and boxList["E_freightIncured"][0][index][0].get()=='0':
                        boxList["E_freightIncured"][0][index][1].set('')
                    elif key == 'E_freightCharged' and boxList["E_freightCharged"][0][index][0].get()=='0':
                        boxList["E_freightCharged"][0][index][1].set('')
                    elif (boxList['C_Quote Yes/No'][0][index][0].get()=="Other" and key == 'E_UOM'):
                        lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                        lbframe.list.focus()
                        lbframe.list.select_set(0)

                elif len(salesDict):
                    key, index = keyFinder(salesDict,(ent,var))
                    if key == "sales_person":
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
        

        def handle_left_click(e):
            try:
                rowclicked_single = pt.get_row_clicked(e)
                print(f"Row clicked is {rowclicked_single+1}")
                if len(boxList) and rowclicked_single < len(pt.model.df):
                    if list(pt.model.df.columns) == ['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in','date_last_receipt','age', 'heat_number', 'lot_serial_number']:
                        if list(pt.model.df.iloc[0]) != [None, None, None, None, None, None, None, None, None]:
                            # pt=pt.getCurrentTable()
                            varname = focused_entry.cget("textvariable")
                            focused_var = focused_entry.getvar(varname)
                            key, index = keyFinder2(boxList,(focused_entry,varname))
                            if key ==  None and index == None:
                                pass
                            else:
                                pt.model.df.reset_index(inplace = True,drop = True)
                                print(key, index)
                                boxList['Lot_Serial_Number'][0][index] = (pt.model.df['lot_serial_number'][rowclicked_single], None)
                                boxList['E_COST'][0][index][1].set(float(pt.model.df['onhand_dollars_per_pounds'][rowclicked_single]))
                                if not len(bakerDf):
                                    boxList['E_freightIncured'][0][index][1].set(0)
                                    boxList['E_freightCharged'][0][index][1].set(0)
                                    boxList['E_Margin_Freight'][0][index][1].set(0)
                                boxList['E_Additional_Cost'][0][index][1].set(0)
                        else:
                            pass
                    else:
                        pass
                pt.setSelectedRow(rowclicked_single)
                pt.redraw()
            except Exception as ex:
                raise ex
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

        if pt is not None:
        
            pt.bind('<Button-1>',handle_left_click)
        
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
                    if len(boxList):# and len(lbframe.list.get(0,tk.END))!=1
                        # boxList["quoteYesNo"]
                        key, index = keyFinder(boxList,(ent,var))   
                        # print(key, index)
                        if key != "C_Quote Yes/No" and (boxList['E_Location'][0][index][0].get()!='' or key in ['searchLocation']) and key != "cus_long_name" and (boxList['E_Location'][0][index][0].get()!='None' or key in ['searchLocation']):# and key != "E_UOM":#and key != 'E_Location' 
                            # if key in ['searchLocation', 'searchYield']:
                            #     newList = filterList(boxList,key,index,df)
                            if not str(boxList["C_Quote Yes/No"][0][index][0].get().upper()).startswith("Y") and key != 'searchLocation':
                                check = False
                                if key == "E_Additional_Cost" and boxList["E_Additional_Cost"][0][index][1].get()!='' and boxList["E_Additional_Cost"][0][index][1].get()!='None':
                                    addCostCalc(boxList, index)
                            elif (key == 'E_ID1' or key == 'E_OD1') and (boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" or boxList['E_Type'][0][index][0].get()=="HM"):
                                check = False
                                if key == 'E_ID1':
                                    if not type(boxList['E_OD2'][0][index][0]) == str and (isinstance(boxList['E_OD2'][0][index][0].get(), str) and isinstance(boxList['E_ID2'][0][index][0].get(), str))and(boxList['E_OD2'][0][index][0].get() != '' and boxList['E_ID2'][0][index][0].get() != '') and (boxList['E_OD2'][0][index][0].get() != 'None' and boxList['E_ID2'][0][index][0].get() != 'None'):
                                    
                                        # newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())& (df["material_type"]==boxList['E_Type'][0][index][0].get())
                                        newDf = df[(df["site"] == boxList['E_Location'][0][index][0].get())
                                                    & (df["grade"]==boxList['E_Grade'][0][index][0].get())& (df["heat_condition"]==boxList['E_Yield'][0][index][0].get())
                                                    & (df["od_in"]==float(boxList['E_OD2'][0][index][0].get())) & (df["od_in_2"]==float(boxList['E_ID2'][0][index][0].get()))]
                                        # newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds','reserved_pieces', 'reserved_length_in', 'available_pieces', 'available_length_in']]
                                        newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in','date_last_receipt','age', 'heat_number', 'lot_serial_number']]
                                        newDf['date_last_receipt'] = pd.to_datetime(newDf['date_last_receipt'])
                                        newDf['date_last_receipt'] = newDf['date_last_receipt'].dt.date
                                        newDf = newDf[newDf['available_pieces']>0]
                                        newDf = newDf.sort_values('age', ascending=False).sort_values('date_last_receipt', ascending=True)
                                        boxList['E_Length'][0][index][0].focus()
                                        #Resetting Index
                                        newDf.reset_index(inplace=True, drop=True)
                                        if pt is not None:
                                            pt.model.df = newDf
                                            pt.redraw()
                                            pt.rowheader.bind('<Button-1>',handle_left_click)
                                            breakCheck = True
                                        ent.configure(state='disabled')
                                    else:
                                        root.attributes('-topmost', True)
                                        messagebox.showerror(title="Wrong Value",message="Please enter value in E_OD2 and I_ED2",parent=root)
                                        root.attributes('-topmost', False)
                                        boxList['E_OD1'][0][index][1].set('')
                                        boxList['E_ID1'][0][index][1].set('')
                            
                            else:
                                if key != "E_UOM":
                                    newList = filterList(boxList,key,index,df)

                            if str(boxList["C_Quote Yes/No"][0][index][0].get().upper())=="OTHER" and (key == "E_UOM" or key == "E_Location"):
                                check = True


                    elif len(salesDict):
                        key, index = keyFinder(salesDict,(ent,var))
                        df["s_person_name_id"] = df['sales_person'] +" | "+ df["code_no"].astype(str) +" | "+ df["location"]
                        newList = df["s_person_name_id"].values.tolist()

                    elif len(cxDict):
                        key, index = keyFinder(cxDict,(ent,var))
                        df["cx_name_id"] = df['cus_long_name'].astype(str) +" | "+ df["cus_id"]
                        newList = df["cx_name_id"].values.tolist()
                        if key == "cus_long_name":
                            if is_on is not None:
                                if is_on.get()=='Off':
                                    newList = []
                    
                    # lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
                    
                    lbframe.list.delete(0, tk.END)
                    if len(newList)==1:
                        if (key == 'E_ID1' or key == 'E_OD1') and (boxList['E_Type'][0][index][0].get()=="TUI" or boxList['E_Type'][0][index][0].get()=="HR" or boxList['E_Type'][0][index][0].get()=="HM"):
                            pass
                        else:
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
    