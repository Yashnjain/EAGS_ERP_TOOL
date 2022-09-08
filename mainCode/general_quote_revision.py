import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
# from customComboboxV2 import myCombobox
from tkcalendar import DateEntry
from datetime import date
import sys,time
import pandas as pd
from pandastable import Table
from RTools import dfMaker, resource_path,specialCase
from RsfTool import get_connection,get_cx_df, get_inv_df
from final_pdf_creator import pdf_generator
from RsfTool import eagsQuotationuploader
import os, shutil
from tkPDFViewer import tkPDFViewer as pdf
import ctypes
from mail import send_mail
from Tools import starSearch, rangeSearch
from shpUploader import shpUploader
ctypes.windll.shcore.SetProcessDpiAwareness(1)


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


def keyFinder(dict, tupleValue):
            try:
                for key, value in dict.items():
                    for index in range(len(value[0])):
                        
                        if value[0][index] == tupleValue:
                            return key,index
            except Exception as e:
                raise e





def general_quote_revision(mainRoot,user,conn,quotedf,quote_number, df):
    try:

        # def list_up(specialList,cx_type_list_var,tupVar):
        #     key, index = keyFinder(specialList,tupVar)
        #     if len(key) and str(specialList[key][0][index][0].get()).upper() not in cx_type_list_var and specialList[key][0][index][0].get()!="":
        #         specialList[key][0][index][1].set("")
        #         messagebox.showerror(title="Wrong Value",message="Please enter value from list only!")
        #     return

        def cost_error(specialList, tupVar):
            try:
                if len(specialList):
                    key, index = keyFinder(specialList,tupVar) 
                    if key=='E_Qty' and (specialList["C_Quote Yes/No"][0][index][0].get()).strip() != 'No':
                            if specialList["E_COST"][0][index][0].get()=='':
                                messagebox.showerror(title="COST not Selected",message="Please select cost price(on_dollars_per_pound) row from table present below")
                                specialList["E_Qty"][0][index][0].focus()

                    elif key=='E_Selling Cost/LBS' and (specialList["C_Quote Yes/No"][0][index][0].get()).strip() != 'No':
                            if specialList["E_COST"][0][index][0].get() == '':
                                messagebox.showerror(title="COST not Selected",message="Please select cost price(on_dollars_per_pound) row from table present below")
                                specialList["E_COST"][0][index][0].focus()
                    else:
                        specialList["E_Selling Cost/LBS"][0][index][0].focus()
                                                  
                # else:
                #     messagebox.showerror(title="Wrong Value",message="Selling Cost/LBS or COST is blank, please fill their respective boxes")
                #     return
            except Exception as ex:
                raise ex

        def margin_cal(specialList, tupVar):
            try:
                if len(specialList):
                    key, index = keyFinder(specialList,tupVar) 
                    if specialList["E_Selling Cost/LBS"][0][index][0].get() != '' and specialList["E_COST"][0][index][0].get() != '' and (specialList["C_Quote Yes/No"][0][index][0].get()).strip() != 'No':
                        salePrice = float(specialList["E_Selling Cost/LBS"][0][index][0].get())
                        costPrice = float(specialList["E_COST"][0][index][0].get())
                        margin_per_lbs = round(((salePrice - costPrice)/salePrice) * 100, 2)
                        specialList["E_MarginLBS"][0][index][1].set(margin_per_lbs)
                        specialList["E_UOM"][0][index][0].focus()
                        # breakCheck = True
                    else:
                        pass
                else:
                    messagebox.showerror(title="Wrong Value",message="Selling Cost/LBS or COST is blank, please fill their respective boxes")
                    return
            except Exception as ex:
                raise ex

        def freight_cal(specialList, tupVar):
            try:
                if len(specialList):
                    key, index = keyFinder(specialList,tupVar) 
                    if specialList["E_freightIncured"][0][index][0].get() != '' and specialList["E_freightCharged"][0][index][0].get() != '' and (specialList["C_Quote Yes/No"][0][index][0].get()).strip() != 'No':
                        #Calculating Freight Margin
                        fcostPrice = float(specialList["E_freightIncured"][0][index][0].get())
                        fsalePrice = float(specialList["E_freightCharged"][0][index][0].get())
                        if fcostPrice != 0 and fsalePrice != 0:
                            margin_freight = round(((fsalePrice - fcostPrice)/fsalePrice) * 100, 2)
                        specialList["E_Margin_Freight"][0][index][1].set(margin_freight)
                        
                        # breakCheck = True
                    else:
                        pass
                else:
                    messagebox.showerror(title="Wrong Value",message="FreightIncured or FreightCharged is blank, please fill their respective boxes")
                    return 
            except Exception as ex:
                raise ex          

        def addCostCalc(specialList,index):
            try:
                addCost = float(specialList["E_Additional_Cost"][0][index][1].get())
                sellCost = float(specialList["E_Selling Cost/UOM"][0][index][1].get())
                finalPrice = addCost+sellCost
                specialList["E_Final Price"][0][index][1].set(finalPrice)
            except Exception as ex:
                raise ex

        def formulaCalc(specialList, tupVar):
            try:
                if len(specialList):
                    key, index = keyFinder(specialList,tupVar) 
                if specialList['C_Quote Yes/No'][0][index][0].get()=="Yes":            
                    if specialList['E_Type'][0][index][0].get()=="TUI" or specialList['E_Type'][0][index][0].get()=="HR" or specialList['E_Type'][0][index][0].get()=="HM":
                        e_od = float(specialList["E_OD2"][0][index][1])
                        e_id = float(specialList["E_ID2"][0][index][1])
                    else:
                        e_od = float(specialList["E_OD1"][0][index][1].get())
                        e_id = float(specialList["E_ID1"][0][index][1].get())
                    uom = specialList["E_UOM"][0][index][1].get()
                    e_length = float(specialList["E_Length"][0][index][1].get())#int(boxList["E_Length"][0][index][1].get())
                    sellCostLBS = float(specialList["E_Selling Cost/LBS"][0][index][1].get())
                    wt = (e_od - e_id)/2
                    # THF
                    if specialList["E_Type"][0][index][1].get() == "THF" or specialList["E_Type"][0][index][1].get() == "TUI" or specialList["E_Type"][0][index][1].get() == "HT":
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
                        #For Foot
                        else:
                            sellCostUOM = round((sellCostLBS * mid_formula),2)
                            sellCostUOM = sellCostUOM *12
                    elif specialList["E_Type"][0][index][1].get() == "BR" or specialList["E_Type"][0][index][1].get() == "HR" or specialList["E_Type"][0][index][1].get() == "HB":
                        # BR
                        # mid_formula = (od*od*2.71)/12
                        mid_formula = ((e_od-wt)*wt*10.68)/12
                        if uom == "Each":
                            # For Each: 
                            # Selling cost/UOM ="SellingCost/LBS" * mid_formula * Length 
                            sellCostUOM = round((sellCostLBS * mid_formula * e_length),2)
                        elif uom == "Inch":
                            # For Inch: 
                            # Selling cost/UOM ="SellingCost/LBS" * mid_formula	
                            sellCostUOM = round((sellCostLBS * mid_formula),2)
                        #For Foot
                        else:
                            sellCostUOM = round((sellCostLBS * mid_formula),2)
                            sellCostUOM = sellCostUOM *12
                    else:
                        sellCostUOM = 0
                    specialList["E_Selling Cost/UOM"][0][index][1].set(sellCostUOM)
                    addCostCalc(specialList,index)    
            except Exception as ex:
                raise ex
     

        def display(specialList,tupVar,df):
            key, index = keyFinder(specialList,tupVar)
            if key == 'E_ID1' and specialList['E_ID1'][0][index][0].get()!="" and specialList['E_ID1'][0][index][0].get()!="NA":
                newDf = df[(df["site"] == specialList['E_Location'][0][index][0].get())
                            & (df["global_grade"]==specialList['E_Grade'][0][index][0].get())& (df["heat_condition"]==specialList['E_Yield'][0][index][0].get())
                            & (df["od_in"]==float(specialList['E_OD1'][0][index][0].get())) & (df["od_in_2"]==float(specialList['E_ID1'][0][index][0].get()))]
                newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in','date_last_receipt','age', 'heat_number', 'lot_serial_number']]
                newDf['date_last_receipt'] = pd.to_datetime(newDf['date_last_receipt'])
                newDf['date_last_receipt'] = newDf['date_last_receipt'].dt.date
                newDf= newDf[newDf['available_pieces']>0]
                newDf = newDf.sort_values('age', ascending=False).sort_values('date_last_receipt', ascending=True)
                specialList['E_Length'][0][index][0].focus()
                #Resetting Index
                newDf.reset_index(inplace=True, drop=True)
                pt.model.df = newDf
                pt.redraw()


        def tui_hr_cehcker(specialList,quotedf,row_num,tupVar,df):
            print('sdgvdf')
            key, index = keyFinder(specialList,tupVar)
            if key == 'E_OD1' and specialList['E_OD1'][0][index][0].get()!="" and specialList['E_OD1'][0][index][0].get()!="NA":
                if specialList['E_Type'][0][index][0].get()=='TUI' or specialList['E_Type'][0][index][0].get()=='HR':
                    specialCase(root, specialList,index,pt,df,row_num,quotedf)
                

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
                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in Integer or Decimal format only",parent=entryFrame)
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
                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in Integer format only",parent=entryFrame)
                    return False
                return True
            except Exception as e:
                raise e
        def on_configure(event):
            try:
                # update scrollregion after starting 'mainloop'
                # when all widgets are in canvas
                entryCanvas.configure(scrollregion=entryCanvas.bbox('all'))#,width=1890,height=380)#(0,0,300,200)
                entryCanvas.yview_moveto('1.0')
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

        def addRow(quotedf,check=None):
            try:
                # for i in range(0,len(quotedf)):                    
                    row_num = len(quoteYesNo)

                    cx_spec_var =tk.StringVar()
                    cx_spec.append((ttk.Entry(entryFrame,width=15,textvariable=cx_spec_var),cx_spec_var))
                    cx_spec[-1][0].grid(row=2+row_num,column=0,padx=(15,0))
                    if check:
                        cx_spec[i][1].set(quotedf['C_SPECIFICATION'][i])    

                    cx_type.append((None, None))
                    # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                    cx_grade_var =tk.StringVar()
                    cx_grade.append((ttk.Entry(entryFrame,width=15,textvariable=cx_grade_var),cx_grade_var))
                    cx_grade[-1][0].grid(row=2+row_num,column=1)
                    if check:
                        cx_grade[i][1].set(quotedf['C_GRADE'][i]) 
                    # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                    cx_yield_var =tk.StringVar()
                    cx_yield.append((ttk.Entry(entryFrame,width=15,textvariable=cx_yield_var),cx_yield_var))
                    cx_yield[-1][0].grid(row=2+row_num,column=2)
                    if check:
                        cx_yield[i][1].set(quotedf['C_YIELD'][i])
                    # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=3,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                    
                    cx_od_var =tk.StringVar()
                    vcmd = tab1.register(intFloat)
                    cx_od.append((ttk.Entry(entryFrame, width=5,validate = "key",
                            validatecommand=(vcmd, '%P','%d'),textvariable=cx_od_var),cx_od_var))
                    cx_od[-1][0].grid(row=2+row_num,column=3)
                    if check:
                        cx_od[i][1].set(quotedf['C_OD'][i])
                    # cx_od['validatecommand'] = (cx_od.register(intFloat),'%P','%d')



                    # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=4,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                    cx_id_var =tk.StringVar()
                    cx_id.append((ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_id_var),cx_id_var))
                    cx_id[-1][0].grid(row=2+row_num,column=4)
                    cx_id[-1][0]['validatecommand'] = (cx_id[-1][0].register(intFloat),'%P','%d')
                    if check:
                        cx_id[i][1].set(quotedf['C_ID'][i])
                    # myCombobox(df,tab1,cx_list,frame=entryFrame,row=1,column=5,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")

                    cx_len_var = tk.StringVar()
                    cx_len.append((ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_len_var),cx_len_var))
                    cx_len[-1][0].grid(row=2+row_num,column=5)
                    cx_len[-1][0]['validatecommand'] = (cx_len[-1][0].register(intFloat),'%P','%d')
                    if check:
                        cx_len[i][1].set(quotedf['C_LENGTH'][i])
                    
                    
                    # myCombobox(df,tab1,cx_list,frame=entryFrame,row=1,column=6,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                    cx_qty_var = tk.StringVar()
                    cx_qty.append((ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_qty_var), cx_qty_var))
                    cx_qty[-1][0].grid(row=2+row_num,column=6)
                    cx_qty[-1][0]['validatecommand'] = (cx_qty[-1][0].register(intChecker),'%P','%d')
                    if check:
                        cx_qty[i][1].set(quotedf['C_QTY'][i])

                    # var =tk.StringVar()
           
                    # quoteYesNo.append((ttk.Entry(entryFrame, width=5, validate = "key", textvariable=var), var))
                    # var.set(C_QUOTE_YES_NO)
                    cx_qyn_var = tk.StringVar()
                    cx_qyn_var_list=["Yes","No","Other"]
                    cx_qyn_entry_var=ttk.Combobox(entryFrame, background='white', font=('Segoe UI', 10), justify='center',textvariable=cx_qyn_var,values=cx_qyn_var_list, width=5)
                    quoteYesNo.append((cx_qyn_entry_var, cx_qyn_var))
                    quoteYesNo[-1][0].grid(row=2+row_num,column=7)
                    if check:
                        cx_qyn_entry_var.insert(tk.END,str(quotedf['C_QUOTE_YES/NO'][i]))
                    # cx_qyn_entry_var.bind("<Leave>",lambda a:list_up(specialList,cx_qyn_var_list,tupVar = (cx_qyn_entry_var, cx_qyn_var)))    

            #--------- old ---------
                    # C_QUOTE_YES_NO=quotedf['C_QUOTE_YES/NO'][i]
                    # quoteYesNo.append(myCombobox(df,tab1,item_list=["Yes","No","Other"],frame=entryFrame,row=2+row_num,column=7,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # quoteYesNo[0][1].set(C_QUOTE_YES_NO)
            #--------- old ---------
                    # quoteYesNo[-1]['validate']='focusout'
                    # quoteYesNo[-1]['validatecommand'] = (quoteYesNo[-1].register(yesNo),'%P','%W')
                    # cx_lctn_var = tk.StringVar()
                    # e_location.append((ttk.Entry(entryFrame, width=10, validate = "key",textvariable=cx_lctn_var), cx_lctn_var))
                    # e_location[-1][0].grid(row=2+row_num,column=8)
                    # if check:
                    #     e_location[i][1].set(quotedf['E_LOCATION'][i])
                    e_location_var = tk.StringVar()
                    e_location_list_var=["Dubai", "USA", "UK", "Singapore"]
                    e_location_entry_var=ttk.Combobox(entryFrame, background='white', font=('Segoe UI', 10), justify='center',textvariable=e_location_var,values=e_location_list_var, width=5)
                    e_location.append((e_location_entry_var, e_location_var))
                    e_location[-1][0].grid(row=2+row_num,column=8)
                    if check:
                        e_location_entry_var.insert(tk.END,str(quotedf['E_LOCATION'][i]))

            #--------- old ---------
                    # E_LOCATION=quotedf['E_LOCATION'][i]
                    # e_location.append(myCombobox(df,tab1,item_list=["Dubai","Singapore","USA","UK"],frame=entryFrame,row=2+row_num,column=8,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # e_location[0][1].set(E_LOCATION)
            #--------- old ---------
                    # e_location[-1].config(textvariable="NA", state='disabled')
                    cx_type_var = tk.StringVar()
                    cx_type_list_var=["THF","BR","TUI","HR"]
                    cx_type_entry_var=ttk.Combobox(entryFrame, background='white', font=('Segoe UI', 10), justify='center',textvariable=cx_type_var,values=cx_type_list_var, width=5)
                    e_type.append((cx_type_entry_var, cx_type_var))
                    e_type[-1][0].grid(row=2+row_num,column=9)
                    if check:
                      cx_type_entry_var.insert(tk.END,str(quotedf['E_TYPE'][i]))
                    # cx_type_entry_var.bind("<Leave>",lambda a:list_up(specialList,cx_type_list_var,tupVar = (cx_type_entry_var, cx_type_var)))

            #--------- old ---------
                    # E_TYPE=quotedf['E_TYPE'][i]
                    # e_type.append(myCombobox(df,tab1,item_list=["THF","BR"],frame=entryFrame,row=2+row_num,column=9,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                    # e_type[0][1].set(E_TYPE)
                    e_spec.append((None, None))
            #--------- old ---------

                    # e_type[-1].config(textvariable="NA", state='disabled')
                    cx_grd_var = tk.StringVar()
                    e_grade.append((ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_grd_var), cx_grd_var))
                    e_grade[-1][0].grid(row=2+row_num,column=10)
                    if check:
                        e_grade[i][1].set(quotedf['E_GRADE'][i])

                    # E_GRADE=quotedf['E_GRADE'][i]
                    # e_grade.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=10,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                    # e_grade[0][1].set(E_GRADE)




                    # e_grade[-1].config(textvariable="NA", state='disabled')
                    cx_yld_var = tk.StringVar()
                    e_yield.append((ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_yld_var), cx_yld_var))
                    e_yield[-1][0].grid(row=2+row_num,column=11)
                    if check:
                        e_yield[i][1].set(quotedf['E_YIELD'][i])



                    # E_YIELD=quotedf['E_YIELD'][i]
                    # e_yield.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=11,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                    # e_yield[0][1].set(E_YIELD)




                    # e_yield[-1].config(textvariable="NA", state='disabled')

                    cx_od1_var = tk.StringVar()
                    cx_ent_od1 = ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_od1_var)
                    e_od1.append((cx_ent_od1, cx_od1_var))
                    e_od1[-1][0].grid(row=2+row_num,column=12)
                    e_od1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')
                    if check:
                        e_od1[i][1].set(quotedf['E_OD1'][i])
                    cx_ent_od1.bind('<1>',lambda a:tui_hr_cehcker(specialList,quotedf,row_num,tupVar = (cx_ent_od1, cx_od1_var),df=df))
                    cx_ent_od1.bind('<Tab>',lambda a:tui_hr_cehcker(specialList,quotedf,row_num,tupVar = (cx_ent_od1, cx_od1_var),df=df))
                    cx_ent_od1.bind('<FocusIn>',remember_focus)
                 

                    # E_OD1=quotedf['E_OD1'][i]
                    # e_od1.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=12,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                    # # e_od[-1].config(textvariable="NA", state='disabled')
                    # e_od1[-1][0]['validate']='key'
                    # e_od1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')
                    # e_od1[0][1].set(E_OD1)
                    
                    cx_id1_var = tk.StringVar()
                    cx_ent_id1 = ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_id1_var)
                    e_id1.append((cx_ent_id1, cx_id1_var))
                    e_id1[-1][0].grid(row=2+row_num,column=13)
                    e_id1[-1][0]['validatecommand'] = (e_id1[-1][0].register(intFloat),'%P','%d')
                    if check:
                        e_id1[i][1].set(quotedf['E_ID1'][i])

                    # new_tags = e_id1[row_num][0].bindtags() + ("mytag",)
                    # e_id1[row_num][0].bindtags(new_tags)

                    # e_id1[row_num][0].bind_class("mytag", '<1>',lambda a:display(specialList,tupVar = (e_id1[-1][0], e_id1[-1][1]),df=df))

                    cx_ent_id1.bind('<1>',lambda a:display(specialList,tupVar = (cx_ent_id1, cx_id1_var),df=df))
                    cx_ent_id1.bind('<Tab>',lambda a:display(specialList,tupVar = (cx_ent_id1, cx_id1_var),df=df))

                    cx_ent_id1.bind('<FocusIn>',remember_focus)


                    # E_ID1=quotedf['E_ID1'][i]
                    # e_id1.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=13,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                    # e_id1[-1][0]['validate']='key'
                    # e_id1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')
                    # e_od1[0][1].set(E_ID1)
                    e_od2.append((None, None))
                    e_id2.append((None, None))

                    # e_id[-1].config(textvariable="NA", state='disabled')

                    cx_lnth_var = tk.StringVar()
                    e_len_ent = ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_lnth_var)
                    e_len.append((e_len_ent, cx_lnth_var))
                    e_len[-1][0].grid(row=2+row_num,column=14)
                    e_len[-1][0]['validatecommand'] = (e_len[-1][0].register(intFloat),'%P','%d')
                    if check:
                        e_len[i][1].set(quotedf['E_LENGTH'][i])
                    e_len_ent.bind('<FocusIn>',remember_focus)
                    # E_LENGTH=quotedf['E_LENGTH'][i]
                    # e_len.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=14,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # e_len[-1][0]['validate']='key'
                    # e_len[-1][0]['validatecommand'] = (e_len[-1][0].register(intFloat),'%P','%d')
                    # e_len[0][1].set(E_LENGTH)




                    # e_len[-1].config(textvariable="NA", state='disabled')
                    cx_qty_var = tk.StringVar()
                    e_qty_ent = ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_qty_var)
                    e_qty.append((e_qty_ent, cx_qty_var))
                    e_qty[-1][0].grid(row=2+row_num,column=15)
                    e_qty[-1][0]['validatecommand'] = (e_qty[-1][0].register(intFloat),'%P','%d')
                    if check:
                        e_qty[i][1].set(quotedf['E_QTY'][i])

                    e_qty_ent.bind("<Tab>", lambda a:cost_error(specialList,tupVar = (e_qty_ent, cx_qty_var)))
                    # e_qty_ent.bind("<Leave>", lambda a:cost_error(specialList,tupVar = (e_qty_ent, cx_qty_var)))
                    e_qty_ent.bind('<FocusIn>',remember_focus)

                    # e_cost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=16,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    e_cost_var = tk.StringVar()
                    e_cost_entry_var = ttk.Entry(entryFrame, width=5, validate = "key",textvariable=e_cost_var)
                    e_cost.append((e_cost_entry_var, e_cost_var))
                    e_cost[-1][0].grid(row=2+row_num,column=16)
                    e_cost[-1][0]['validate']='key'
                    e_cost[-1][0]['validatecommand'] = (e_cost[-1][0].register(intFloat),'%P','%d')
                    if check:
                        e_cost[i][1].set(quotedf['E_COST'][i])

                    # cx_ent_finalcost.bind('<1>',lambda a:formulaCalc(specialList,tupVar = (cx_ent_finalcost, cx_fc_var)))
                    e_cost_entry_var.bind("<Leave>", lambda a:margin_cal(specialList,tupVar = (e_cost_entry_var, e_cost_var)))
                    e_cost_entry_var.bind("<Tab>", lambda a:margin_cal(specialList,tupVar = (e_cost_entry_var, e_cost_var)))
               
                    # E_QTY=quotedf['E_QTY'][i]
                    # e_qty.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=15,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # e_qty[-1][0]['validate']='key'
                    # e_qty[-1][0]['validatecommand'] = (e_qty[-1][0].register(intChecker),'%P','%d')
                    # e_qty[0][1].set(E_QTY)

                    # e_qty[-1].config(textvariable="NA", state='disabled')
                    cx_scl_var = tk.StringVar()
                    cx_scl_entry_var = ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_scl_var)
                    sellCostLBS.append((cx_scl_entry_var, cx_scl_var))
                    sellCostLBS[-1][0].grid(row=2+row_num,column=17)
                    sellCostLBS[-1][0]['validatecommand'] = (sellCostLBS[-1][0].register(intFloat),'%P','%d')
                    if check:
                        sellCostLBS[i][1].set(quotedf['E_SELLING_COST/LBS'][i])


                    # cx_scl_entry_var.bind("<Tab>", lambda a:cost_error(specialList,tupVar = (cx_scl_entry_var, cx_scl_var)))
                    # cx_scl_entry_var.bind("<Leave>", lambda a:cost_error(specialList,tupVar = (cx_scl_entry_var, cx_scl_var)))
                    cx_scl_entry_var.bind("<Leave>", lambda a:margin_cal(specialList,tupVar = (cx_scl_entry_var, cx_scl_var)))
                    cx_scl_entry_var.bind("<Tab>", lambda a:margin_cal(specialList,tupVar = (cx_scl_entry_var, cx_scl_var)))

                    # marginlbs.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=18,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    marginlbs_var = tk.StringVar()
                    marginlbs.append((ttk.Entry(entryFrame, width=5, validate = "key",textvariable=marginlbs_var), marginlbs_var))
                    marginlbs[-1][0].grid(row=2+row_num,column=18)
                    marginlbs[-1][0]['validate']='key'
                    marginlbs[-1][0]['validatecommand'] = (marginlbs[-1][0].register(intFloat),'%P','%d')
                    if check:
                        marginlbs[i][1].set(quotedf['E_MARGIN_LBS'][i])
                    # E_SELLING_COST_LBS=quotedf['E_SELLING_COST/LBS'][i]
                    # sellCostLBS.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=16,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # sellCostLBS[-1][0]['validate']='key'
                    # sellCostLBS[-1][0]['validatecommand'] = (sellCostLBS[-1][0].register(intFloat),'%P','%d')
                    # sellCostLBS[0][1].set(E_SELLING_COST_LBS)




                    # sellCostLBS[-1].config(textvariable="NA", state='disabled')
                    cx_uom_var = tk.StringVar()
                    cx_uom_var_list = ["Inch","Each","Foot"]
                    cx_uom_entry_var = ttk.Combobox(entryFrame, background='white', font=('Segoe UI', 10), justify='center',textvariable=cx_uom_var,values=cx_uom_var_list, width=5)
                    e_uom.append((cx_uom_entry_var, cx_uom_var))
                    e_uom[-1][0].grid(row=2+row_num,column=19)
                    if check:
                        cx_uom_entry_var.insert(tk.END,str(quotedf['E_UOM'][i]))
                    # cx_uom_entry_var.bind("<Leave>",lambda a:list_up(specialList,cx_uom_var_list,tupVar = (cx_uom_entry_var, cx_uom_var)))

                    # E_UOM=quotedf['E_UOM'][i]
                    # e_uom.append(myCombobox(df,tab1,item_list=["Inch","Each"],frame=entryFrame,row=2+row_num,column=17,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # sellCostLBS[0][1].set(E_UOM)



                    # e_uom[-1].config(textvariable="NA", state='disabled')


                    cx_cuom_var = tk.StringVar()
                    sellCostUOM.append((ttk.Entry(entryFrame, width=8, validate = "key",textvariable=cx_cuom_var), cx_cuom_var))
                    sellCostUOM[-1][0].grid(row=2+row_num,column=20)
                    sellCostUOM[-1][0]['validatecommand'] = (sellCostUOM[-1][0].register(intFloat),'%P','%d')
                    if check:
                        sellCostUOM[i][1].set(quotedf['E_SELLING_COST/UOM'][i])

                    # E_SELLING_COST_UOM=quotedf['E_SELLING_COST/UOM'][i]
                    # sellCostUOM.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=18,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # sellCostUOM[-1][0]['validate']='key'
                   
                    # sellCostUOM[0][1].set(E_SELLING_COST_UOM)


                    # sellCostUOM[-1].config(textvariable="NA", state='disabled')
                    cx_ac_var = tk.StringVar()
                    addCost.append((ttk.Entry(entryFrame, width=5, validate = "key",textvariable=cx_ac_var), cx_ac_var))
                    addCost[-1][0].grid(row=2+row_num,column=21)
                    addCost[-1][0]['validatecommand'] = (addCost[-1][0].register(intFloat),'%P','%d')
                    if check:
                        addCost[i][1].set(quotedf['E_ADDITIONAL_COST'][i])

                    # E_ADDITIONAL_COST=quotedf['E_ADDITIONAL_COST'][i]
                    # addCost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=19,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # addCost[-1][0]['validate']='key'
                    # addCost[-1][0]['validatecommand'] = (addCost[-1][0].register(intFloat),'%P','%d')
                    # sellCostUOM[0][1].set(E_ADDITIONAL_COST)


                    # addCost[-1].config(textvariable="NA", state='disabled')
                    cx_lt_var = tk.StringVar()
                    leadTime.append((ttk.Entry(entryFrame, width=10, validate = "key",textvariable=cx_lt_var), cx_lt_var))
                    leadTime[-1][0].grid(row=2+row_num,column=22)
                    if check:
                        leadTime[i][1].set(quotedf['LEAD_TIME'][i])
                    # LEAD_TIME=quotedf['LEAD_TIME'][i]
                    # leadTime.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=20,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # leadTime[0][1].set(LEAD_TIME)


                    # leadTime[-1].config(textvariable="NA", state='disabled')
                    cx_fc_var = tk.StringVar()
                    cx_ent_finalcost = ttk.Entry(entryFrame, width=8, validate = "key",textvariable=cx_fc_var)
                    finalCost.append((cx_ent_finalcost, cx_fc_var))
                    finalCost[-1][0].grid(row=2+row_num,column=23)
                    finalCost[-1][0]['validatecommand'] = (finalCost[-1][0].register(intFloat),'%P','%d')
                    if check:
                        finalCost[i][1].set(quotedf['E_FINAL_PRICE'][i])
                    
                        
                    cx_ent_finalcost.bind('<1>',lambda a:formulaCalc(specialList,tupVar = (cx_ent_finalcost, cx_fc_var)))


                    # freightIncured.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=24,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    freightIncured_var = tk.StringVar()
                    freightIncured_entry_var = ttk.Entry(entryFrame, width=5, validate = "key",textvariable=freightIncured_var)
                    freightIncured.append((freightIncured_entry_var, freightIncured_var))
                    freightIncured[-1][0].grid(row=2+row_num,column=24)
                    if check:
                        freightIncured[i][1].set(quotedf['E_FREIGHT_INCURED'][i])

                    freightIncured_entry_var.bind("<Leave>", lambda a:freight_cal(specialList,tupVar = (freightIncured_entry_var, freightIncured_var)))    
                    freightIncured_entry_var.bind("<Tab>", lambda a:freight_cal(specialList,tupVar = (freightIncured_entry_var, freightIncured_var)))    

                    # freightCharged.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=25,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))

                    freightCharged_var = tk.StringVar()
                    freightCharged_entry_var = ttk.Entry(entryFrame, width=5, validate = "key",textvariable=freightCharged_var)
                    freightCharged.append((freightCharged_entry_var, freightCharged_var))
                    freightCharged[-1][0].grid(row=2+row_num,column=25)
                    if check:
                        freightCharged[i][1].set(quotedf['E_FREIGHT_CHARGED'][i])

                    freightIncured_entry_var.bind("<Leave>", lambda a:freight_cal(specialList,tupVar = (freightCharged_entry_var, freightCharged_var)))
                    freightIncured_entry_var.bind("<Tab>", lambda a:freight_cal(specialList,tupVar = (freightCharged_entry_var, freightCharged_var)))

                    # marginFreight.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=26,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))

                    marginFreight_var = tk.StringVar()
                    marginFreight.append((ttk.Entry(entryFrame, width=5, validate = "key",textvariable=marginFreight_var), marginFreight_var))
                    marginFreight[-1][0].grid(row=2+row_num,column=26)
                    if check:
                        marginFreight[i][1].set(quotedf['E_MARGIN_FREIGHT'][i])
                    
                    lot_serial_number.append((None, None))

                    # E_FINAL_PRICE=quotedf['E_FINAL_PRICE'][i]
                    # finalCost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=21,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                    # finalCost[-1][0]['validate']='key'
                    
                    # finalCost[0][1].set(E_FINAL_PRICE)
            except Exception as e:
                raise e

        def cxListCalc():
            try:
                # finalCost[-1].config(textvariable="NA", state='disabled')
                cxList = [cxDatadict["Prepared_By"],cxDatadict["Date"],cxDatadict["cus_long_name"][0][0][0].get(), cxDatadict["payment_term"][0][0].get(),currency.get(), cxDatadict["cus_address"][0][0].get(),
                    cxDatadict["cus_phone"][0][0].get(),cxDatadict["cus_email"][0][0].get(),cxDatadict["cus_city_zip"]]
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
                    if (len(specialList[key][0])==1):
                        if key!='E_OD2' and key != 'E_ID2' and key != 'C_Type' and key != 'E_Spec':
                            specialList[key][0][0][0].configure(state='normal')
                            specialList[key][0][0][0].delete(0, tk.END)
                        
                        # time.sleep(1)
                        # addRow()
                    else:
                        if key!='E_OD2' and key != 'E_ID2' and key != 'C_Type' and key != 'E_Spec':
                            specialList[key][0][-1][0].destroy()
                        specialList[key][0].pop()
                # show bottom of canvas
                entryCanvas.yview("moveto", 0)
                entryCanvas.yview_moveto('1.0')
            except Exception as e:
                raise e
        
        def create_pdf():
            try:
                global quoteDf,latest_revised_quote
                quoteDf,latest_revised_quote = dfMaker(specialList,cxListCalc(),otherListCalc(),pt,conn,quote_number)
                if len(quoteDf):
                    pt.model.df = quoteDf
                    pt.redraw()
                    
                    global pdf_path
                    pdf_path = pdf_generator(quoteDf)
                    pdfRoot = tk.Toplevel()
                    pdfRoot.title(quoteDf["QUOTENO"][0])
                    pdfviewer = pdf.ShowPdf()
                    # zoom_scale = tk.Scale(pdfRoot, orient='vertical', from_=1, to=500)
                    # zoom_scale.config(command=zoom)
                    # Adding pdf path and width and height.
                    # zoom_scale.pack(fill='y', side='right')
                    # zoom_scale.set(10)
                    pdfframe = pdfviewer.pdf_view(pdfRoot, pdf_location=pdf_path, width=120 ,zoomDPI=100)
                    pdfframe.pack(expand=True, fill='both')
                    pdfRoot.state('zoomed')
                    submitButton.configure(state='normal')
                else:
                    messagebox.showerror("Error", "Empty dataframe was given in input")
                return latest_revised_quote    
            except Exception as e:
                raise e

        def uploadDf(conn, quoteDf ,latest_revised_quote):
            try:
                # pt.model.df = quoteDf
                # pt.redraw()
                if messagebox.askyesno("Upload to Database", "Are sure that you want to generate quote and upload Data?"):
                    eagsQuotationuploader(conn, quoteDf,latest_revised_quote)
                    
                    messagebox.showinfo("Info", "Data uploaded Successfully!")

                    current_work_dir = os.getcwd()
                    # current_work_dir = r'I:\EAGS\Quotes'
                    cx_init_name = str(quoteDf['QUOTENO'][0]).split("_")[0]
                    filename = str(quoteDf['QUOTENO'][0])+".pdf"
                    # save_dir = current_work_dir+"\\"+cx_init_name
                    # if not os.path.exists(save_dir):
                    #     os.mkdir(save_dir)
                    # os.rename(pdf_path,save_dir+"\\"+filename)
                    desktopDir = os.path.join(os.environ["HOMEPATH"], "Desktop\\EAGS_Quotes")
                    desktopDir = os.path.join('C:', desktopDir)
                    if not os.path.exists(desktopDir):
                        os.mkdir(desktopDir)
                    # shutil.copy(save_dir+"\\"+filename, desktopDir)
                    shpUploader(pdf_path,filename)
                    shutil.move(pdf_path,desktopDir+"\\"+filename)
                    send_mail(receiver_email = user[-1], mail_subject=f"ALERT Revision generated by {user[0]} for {quoteDf['QUOTENO'][0]}", 
                    mail_body= f"{user[0]} has generated revision for quote number {quoteDf['QUOTENO'][0]}, initial quote was {quote_number}  on {str(date.today())}",
                    attachment_locations=[desktopDir+"\\"+filename])
                    
                else:
                    os.remove(pdf_path)
                submitButton.configure(state='disable')
            except Exception as e:
                raise e

        def remember_focus(event):
            global focused_entry
            focused_entry = event.widget
        mainRoot.withdraw()
        global row_num
        row_num=0

        #Getting invoentory dataframe
        # df = get_inv_df(conn,table = INV_TABLE)
        
        
        # df = pd.read_excel("sampleInventory.xlsx")
        # Getting Cx Dataframe
        # cx_df = get_cx_df(conn,table = CX_TABLE)
        # cx_df = pd.read_excel("cxDatabase.xlsx")

        count = 0
        root = tk.Toplevel(mainRoot, bg = "#9BC2E6")
        root.state('zoomed')
        root.title('EAGS Quote Generator Revision')
        tabControl = ttk.Notebook(root)
        s = ttk.Style(tabControl)
        s.configure("TFrame", background=root["bg"])
        tab1 = ttk.Frame(tabControl)
        tab2 = ttk.Frame(tabControl)
        tab3 = ttk.Frame(tabControl)
        

        tabControl.add(tab1, text='Quote Generator Revision')
        tabControl.add(tab2, text='Machining Revision')
        tabControl.add(tab3, text='Quote Generator Revision + Machining')

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

        cxFrame.grid(row=0, column=0,pady=(24,0), padx=(30,0),sticky="nsew")
        cxFrame2.grid(row=0, column=1,pady=(24,0), padx=(30,120),sticky="nsew")
        # headerFrame.grid(row=1, column=0,sticky="sew",columnspan=2)
        m_entryFrame.grid(row=1, column=0,sticky="nsew", columnspan=2)
        xscrollbar.grid(row=1,column=0,sticky=tk.NSEW)
        yscrollbar.grid(row=0,column=1,sticky=tk.NSEW)
        entryCanvas.grid(row=0,column=0, sticky=tk.NSEW)
        databaseFrame.grid(row=2,column=0, sticky="nsew")
        controlFrame.grid(row=2,column=1, sticky="nsew")
        
        entryCanvas.create_window((0,0),window=entryFrame,tags='expand')
        
        # entryFrame.grid(row=0,column=0)
        # tab1.grid_rowconfigure(0, weight=1) # For row 0
        # tab1.grid_rowconfigure(1, weight=1) # For row 1
        # tab1.grid_rowconfigure(2, weight=1) # For row 1

        # tab1.grid_columnconfigure(0, weight=1) # For column 0
        # tab1.grid_columnconfigure(1, weight=1) # For column 1


        # cxFrame.grid_rowconfigure(0, weight=1) # For row 0
        # cxFrame.grid_rowconfigure(1, weight=1) # For row 1
        # cxFrame.grid_rowconfigure(2, weight=1) # For row 2
        # cxFrame.grid_rowconfigure(3, weight=1) # For row 3
        # cxFrame.grid_rowconfigure(4, weight=1) # For row 4

        # cxFrame.grid_columnconfigure(0, weight=1) # For column 0
        # cxFrame.grid_columnconfigure(1, weight=1) # For column 1
        # cxFrame.grid_columnconfigure(2, weight=1) # For column 2
        # cxFrame.grid_columnconfigure(3, weight=1) # For column 3
        
        # cxFrame2.grid_rowconfigure(0, weight=1) # For row 0
        # cxFrame2.grid_rowconfigure(1, weight=1) # For row 1
        # cxFrame2.grid_rowconfigure(2, weight=1) # For row 1

        # cxFrame2.grid_columnconfigure(0, weight=1) # For column 0
        # cxFrame2.grid_columnconfigure(1, weight=1) # For column 1
        # cxFrame2.grid_columnconfigure(2, weight=1) # For column 1


        # m_entryFrame.grid_rowconfigure(0, weight=1) # For row 0
        # # m_entryFrame.grid_rowconfigure(1, weight=1) # For row 1

        # m_entryFrame.grid_columnconfigure(0, weight=1) # For column 0
        # # m_entryFrame.grid_columnconfigure(1, weight=1) # For column 1

        # entryFrame.grid_rowconfigure(0, weight=1) # For row 0
        # entryFrame.grid_rowconfigure(1, weight=1) # For row 1
        # entryFrame.grid_rowconfigure(2, weight=1) # For row 1

        # entryFrame.grid_columnconfigure(0, weight=1) # For column 0
        # entryFrame.grid_columnconfigure(1, weight=1) # For column 1
        # entryFrame.grid_columnconfigure(2, weight=1) # For column 2
        # entryFrame.grid_columnconfigure(3, weight=1) # For column 3
        # entryFrame.grid_columnconfigure(4, weight=1) # For column 4
        # entryFrame.grid_columnconfigure(5, weight=1) # For column 5
        # entryFrame.grid_columnconfigure(6, weight=1) # For column 6
        # entryFrame.grid_columnconfigure(7, weight=1) # For column 7
        # entryFrame.grid_columnconfigure(8, weight=1) # For column 8
        # entryFrame.grid_columnconfigure(9, weight=1) # For column 9
        # entryFrame.grid_columnconfigure(10, weight=1) # For column 10
        # entryFrame.grid_columnconfigure(11, weight=1) # For column 11
        # entryFrame.grid_columnconfigure(12, weight=1) # For column 12
        # entryFrame.grid_columnconfigure(13, weight=1) # For column 13
        # entryFrame.grid_columnconfigure(14, weight=1) # For column 14
        # entryFrame.grid_columnconfigure(15, weight=1) # For column 15
        # entryFrame.grid_columnconfigure(16, weight=1) # For column 16
        # entryFrame.grid_columnconfigure(17, weight=1) # For column 17
        # entryFrame.grid_columnconfigure(18, weight=1) # For column 18
        # entryFrame.grid_columnconfigure(19, weight=1) # For column 19
        # entryFrame.grid_columnconfigure(20, weight=1) # For column 20
        # entryFrame.grid_columnconfigure(21, weight=1) # For column 21
        
        # controlFrame.grid_rowconfigure(0, weight=1) # For column 21
        # controlFrame.grid_rowconfigure(1, weight=1) # For column 21

        
        # controlFrame.grid_rowconfigure(2, weight=1) # For column 21
    
        
        # controlFrame.grid_rowconfigure(3, weight=1) # For column 21
        # controlFrame.grid_columnconfigure(1, weight=1) # For column 21

        # databaseFrame.grid_rowconfigure(1, weight=1) # For column 21
        # databaseFrame.grid_columnconfigure(1, weight=1) # For column 21
      ###############Importing Images fir buttons #####################################
        button_dict = {}#defining dict storing images for hover effect
        home_path = resource_path("home(2).png")
        home_path1 = resource_path("home(4).png")
        add_img_path = resource_path("addRowS.png")
        add_img2_path = resource_path("addRow2S.png")
        delete_img_path = resource_path("deleteRowS.png")
        delete_img2_path = resource_path("deleteRow2S.png")
        preview_img_path = resource_path("previewButtonS.png")
        preview_img2_path = resource_path("previewButton2S.png")
        submit_img_path = resource_path("submitButtonS.png")
        submit_img2_path = resource_path("submitButton2S.png")
        home_img = tk.PhotoImage(master=tab1, file=home_path)
        home_img1 = tk.PhotoImage(master=tab1, file=home_path1)
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
        nonList = [[None,None,None,None,None,None,None,None,None]]
        pandasDf = pd.DataFrame(nonList,columns=['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in','date_last_receipt','age', 'heat_number', 'lot_serial_number'])
        # pandasDf = pd.DataFrame(cx_df)
        pt = Table(databaseFrame, editable=False,dataframe=pandasDf,showtoolbar=False, showstatusbar=True, maxcellwidth=1500)
        pt.cellwidth=145
        pt.thefont = ('Segoe UI', 12)
        pt.rowheight = 30
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

        cxDatadict["CURRENCY"] = []


        # item_list = ('A4140', 'A4140M', 'A4330V', 'A4715', 'BS708M40', 'A4145M', '4542','4462')

        cxLabel = tk.Label(cxFrame, text="Customer Details", bg = "#9BC2E6")
        lb1 = tk.Label(cxFrame,text="Prepared By", bg = "#9BC2E6")
        lb2 = tk.Label(cxFrame,text="Date", bg = "#9BC2E6")
        lb3 = tk.Label(cxFrame,text="Customer Name", bg = "#9BC2E6")
        lb4 = tk.Label(cxFrame,text="Location/Address", bg = "#9BC2E6")
        lb5 = tk.Label(cxFrame,text="Email", bg = "#9BC2E6")
        lb6 = tk.Label(cxFrame,text="Payment Terms", bg = "#9BC2E6")

        #Adding Search button in cxFrame 2
        # starButton = tk.Button(cxFrame2, text="Star Search", font = ("Segoe UI", 10, 'bold'), bg="#20bebe", fg="white", height=1, width=14, command=lambda: starSearch(root, df), activebackground="#20bebb", highlightbackground="#20bebd")
        rangeButton = tk.Button(cxFrame2, text="Range Search", font = ("Segoe UI", 10, 'bold'), bg="#20bebe", fg="white", height=1, width=14, command=lambda: rangeSearch(root, df, specialList, 0), activebackground="#20bebb", highlightbackground="#20bebd")

        lb_ex= tk.Label(cxFrame,text="Mobile No.", bg = "#9BC2E6")
        blanckLabel = tk.Label(cxFrame2,text="", bg = "#9BC2E6")
        lb8 = tk.Label(cxFrame2,text="Validity", bg = "#9BC2E6")
        lb9 = tk.Label(cxFrame2,text="Additional Comments", bg = "#9BC2E6")
        lb7 = tk.Label(cxFrame2,text="Currency", bg = "#9BC2E6")
        cxLabel.grid(row=0,column=0)
        lb1.grid(row=1,column=0)
        lb2.grid(row=1,column=1)
        lb3.grid(row=3,column=0)
        lb4.grid(row=3,column=1)
        lb5.grid(row=3,column=2)
        lb6.grid(row=3,column=3)

        #Adding Search button in cxFrame 2
        # starButton.grid(row=0, column=0, pady=(20,0))
        rangeButton.grid(row=0,column=1, pady=(20,0))

        lb_ex.grid(row=3,column=4)
        blanckLabel.grid(row=0,column=0)
        lb7.grid(row=1,column=0)
        lb8.grid(row=1,column=1,padx=(100,5))
        lb9.grid(row=1,column=2)

        
        prep_by = ttk.Entry(cxFrame)
        new_usr=quotedf['PREPAREDBY'][0]
        prep_by.insert(tk.END, new_usr)
        prep_by.grid(row=2,column=0)
        prep_by.config(state= "disabled")
        cxDatadict["Prepared_By"] = prep_by.get()

        
        inpDate = MyDateEntry(master=cxFrame, width=17, selectmode='day')
        inpDate.grid(row=2, column=1)
        cxDatadict["Date"] = inpDate.get()
        

        
        #Validity
        validityVar = tk.StringVar()
        validity = ttk.Entry(cxFrame2, textvariable=validityVar, foreground='blue', background = 'white',width = 15, font=('Segoe UI', 10))
        validity.grid(row=2,column=1,padx=(100,5),pady=5)
        validityVar.set(quotedf['VALIDITY'][0])
        
        #Additional Comments
        addCommVar = tk.StringVar()
        addComm = ttk.Entry(cxFrame2, textvariable=addCommVar, foreground='blue', background = 'white',width = 15, font=('Segoe UI', 10))
        addComm.grid(row=2,column=2,sticky=tk.EW,padx=5,pady=5)
        addCommVar.set(quotedf['ADD_COMMENTS'][0])
        addComm.bind("<Tab>",tabFunc)
       
        #CURRENCY
        currencyVar = tk.StringVar()
        currency = ttk.Combobox(cxFrame2, background='white', font=('Segoe UI', 10), justify='center',textvariable=currencyVar,values=["$","€"], width=5)
        currency.grid(row=2,column=0,sticky=tk.EW,padx=5,pady=5)
        currency.insert(tk.END,str(quotedf['CURRENCY'][0]))
        cxDatadict["CURRENCY"].append((currency, currencyVar))

        #Mobile Number
        mobileVar = tk.StringVar()
        mobile = ttk.Entry(cxFrame, textvariable=mobileVar, foreground='blue', background = 'white',width = 20, font=('Segoe UI', 10))
        mobile.grid(row=4,column=4,sticky=tk.EW,padx=5,pady=5)
        mobileVar.set(quotedf['CUS_PHONE'][0])
        # cxDatadict["mobile_number"].append((mobile, mobileVar))
        cxDatadict["cus_phone"].append((mobile, mobileVar))
        
        #Customer Name Entry Box
        cxName1Var = tk.StringVar()
        cxNamer = ttk.Entry(cxFrame, textvariable=cxName1Var, foreground='blue', background = 'white',width = 5, font=('Segoe UI', 10))
        cxNamer.grid(row=4,column=0,sticky=tk.EW,padx=5,pady=5)
        cxName1Var.set(quotedf['CUS_NAME'][0])
        cxNameVar.append((cxNamer, cxName1Var))
        




        # cxNameVar.append(myCombobox(cx_df,tab1,item_list=item_list,frame=cxFrame,row=4,column=0,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",cxDict= cxDatadict,val=validity))
        
        #location Address entry box
        locAddVar = tk.StringVar()
        locAdd = ttk.Entry(cxFrame, textvariable=locAddVar, foreground='blue', background = 'white',width = 20, font=('Segoe UI', 10))
        locAdd.grid(row=4,column=1,sticky=tk.EW,padx=5,pady=5)
        locAddVar.set(quotedf['CUS_ADDRESS'][0])
        
        # cxLocVar = []
        cxDatadict["cus_address"].append((locAdd, locAddVar))
        # myCombobox(df,root,item_list,frame=cxFrame,row=4,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")

        #Email
        emailAddVar = tk.StringVar()
        emailAdd = ttk.Entry(cxFrame, textvariable=emailAddVar, foreground='blue', background = 'white',width = 20, font=('Segoe UI', 10))
        emailAdd.grid(row=4,column=2,sticky=tk.EW,padx=5,pady=5)
        emailAddVar.set(quotedf['CUS_EMAIL'][0])
        # cxemailAddVar = []
        cxDatadict["cus_email"].append((emailAdd, emailAddVar))
        # myCombobox(df,root,item_list,frame=cxFrame,row=4,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")

        #Payment Terms Entry
        payTermVar = tk.StringVar()
        payTerm = ttk.Entry(cxFrame, textvariable=payTermVar, foreground='blue', background = 'white',width = 20, font=('Segoe UI', 10))
        payTerm.grid(row=4,column=3,sticky=tk.EW,padx=5,pady=5)
        payTermVar.set(quotedf['PAYMENT_TERM'][0])
        cxDatadict["payment_term"].append((payTerm, payTermVar))
        

        home_button = tk.Button(cxFrame2, image=home_img, borderwidth=0,bg=root["bg"],activebackground=root["bg"],command=returnTohome)
        home_button.image = home_img #Preventing image to go into garbage
        home_button.grid(row=0,column=3,sticky="ne")
        button_dict[home_button] = [home_img, home_img1]
        home_button.bind("<Enter>", on_enter)
        home_button.bind("<Leave>", on_leave)
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

        e_costLabel = tk.Label(entryFrame, text="Cost", bg= "#DDEBF7")

        sellcostLbsLabel1 = tk.Label(entryFrame, text="Selling", bg= "#DDEBF7")
        sellcostLbsLabel2 = tk.Label(entryFrame, text="Cost/LBS", bg= "#DDEBF7")

        marginLBSLabel1 = tk.Label(entryFrame, text="Margin/LBS", bg= "#DDEBF7")
        marginLBSLabel2 = tk.Label(entryFrame, text="%", bg= "#DDEBF7")

        uom = tk.Label(entryFrame, text="UOM", bg= "#DDEBF7")
        sellcostUOMLabel1 = tk.Label(entryFrame, text="Selling", bg= "#DDEBF7")
        sellcostUOMLabel2 = tk.Label(entryFrame, text="Cost/UOM", bg= "#DDEBF7")
        addCostLabel1 = tk.Label(entryFrame, text="Additional", bg= "#DDEBF7")
        addCostLabel2 = tk.Label(entryFrame, text="Cost", bg= "#DDEBF7")
        leadTimeLAbel = tk.Label(entryFrame, text="Lead Time", bg= "#DDEBF7")
        finalPriceLabel = tk.Label(entryFrame, text="Final Price", bg= "#DDEBF7")

        freightCostLabel1 = tk.Label(entryFrame, text="Freight", bg= "#DDEBF7")
        freightCostLabel2 = tk.Label(entryFrame, text="Incured", bg= "#DDEBF7")
        freightSaleLabel1 = tk.Label(entryFrame, text="Freight to", bg= "#DDEBF7")
        freightSaleLabel2 = tk.Label(entryFrame, text="be Charged", bg= "#DDEBF7")
        
        

        marginFreightLabel1 = tk.Label(entryFrame, text="Freight Margin", bg= "#DDEBF7")
        marginFreightLabel2 = tk.Label(entryFrame, text="%", bg= "#DDEBF7")


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

        e_costLabel.grid(row=0,column=16, sticky="ew")

        sellcostLbsLabel1.grid(row=0,column=17, sticky="ew")
        sellcostLbsLabel2.grid(row=1,column=17, sticky="ew")

        marginLBSLabel1.grid(row=0,column=18, sticky="ew")
        marginLBSLabel2.grid(row=1,column=18, sticky="ew")

        uom.grid(row=0,column=19, sticky="ew")
        sellcostUOMLabel1.grid(row=0,column=20, sticky="ew")
        sellcostUOMLabel2.grid(row=1,column=20, sticky="ew")
        addCostLabel1.grid(row=0,column=21, sticky="ew")
        addCostLabel2.grid(row=1,column=21, sticky="ew")
        leadTimeLAbel.grid(row=0,column=22, sticky="ew")
        finalPriceLabel.grid(row=0,column=23,padx=(0,10), sticky="ew")
        freightCostLabel1.grid(row=0,column=24, sticky="ew")
        freightCostLabel2.grid(row=1,column=24, sticky="ew")
        freightSaleLabel1.grid(row=0,column=25, sticky="ew")
        freightSaleLabel2.grid(row=1,column=25, sticky="ew")

        

        marginFreightLabel1.grid(row=0,column=26,padx=(0,10), sticky="ew")
        marginFreightLabel2.grid(row=1,column=26,padx=(0,10), sticky="ew")
        ###################################################################
        ######################Defining List variables for various entry boxes######################
        global specialList
        
        specialList = {}
        


        #General Quote Form Variables
        cx_spec = []
        specialList["C_Specification"] = []
        specialList["C_Specification"].append(cx_spec)

        cx_type = []
        specialList["C_Type"] = []
        specialList["C_Type"].append(cx_type)

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

        e_spec = []
        specialList["E_Spec"] = []
        specialList["E_Spec"].append(e_spec)

        e_grade = []
        specialList["E_Grade"] = []
        specialList["E_Grade"].append(e_grade)

        e_yield = []
        specialList["E_Yield"] = []
        specialList["E_Yield"].append(e_yield)

        e_od1 = []
        specialList["E_OD1"] = []
        specialList["E_OD1"].append(e_od1)

        e_id1 = []
        specialList["E_ID1"] = []
        specialList["E_ID1"].append(e_id1)

        e_od2 = []
        specialList["E_OD2"] = []
        specialList["E_OD2"].append(e_od2)

        e_id2 = []
        specialList["E_ID2"] = []
        specialList["E_ID2"].append(e_id2)

        e_len = []
        specialList["E_Length"] = []
        specialList["E_Length"].append(e_len)

        e_qty = []
        specialList["E_Qty"] = []
        specialList["E_Qty"].append(e_qty)

        e_cost = []
        specialList["E_COST"] = []
        specialList["E_COST"].append(e_cost)


        sellCostLBS = []
        specialList["E_Selling Cost/LBS"] = []
        specialList["E_Selling Cost/LBS"].append(sellCostLBS)

        marginlbs = []
        specialList["E_MarginLBS"] = []
        specialList["E_MarginLBS"].append(marginlbs)

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

        freightIncured = []
        specialList["E_freightIncured"] = []
        specialList["E_freightIncured"].append(freightIncured)

        freightCharged = []
        specialList["E_freightCharged"] = []
        specialList["E_freightCharged"].append(freightCharged)

        

        marginFreight = []
        specialList["E_Margin_Freight"] = []
        specialList["E_Margin_Freight"].append(marginFreight)

        lot_serial_number = []
        specialList["Lot_Serial_Number"] = []
        specialList["Lot_Serial_Number"].append(lot_serial_number)
        #For range search
        searchLocation = []
        specialList["searchLocation"] = []
        specialList["searchLocation"].append(searchLocation)

        # specialList = [[quoteYesNo],[e_location], [e_type], [e_grade], [e_yield], [e_od], [e_id], [e_len], [e_qty], [sellCostLBS], [sellCostUOM],
        # [e_uom], [addCost], [leadTime], [finalCost]]
        ###########################################################################################
        
        # var = tk.StringVar()
        # spec = ttk.Entry(entryFrame,textvariable=var, foreground='blue',background='white',width=5)
        # spec.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=0,width=2,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        i=0
        while i<len(quotedf):
            addRow(quotedf,check=True)
            i+=1
        
        # button_dict = {}

        addRowbut = tk.Button(controlFrame, image=add_img, command=lambda:addRow(quotedf),borderwidth=0, background=controlFrame["bg"])
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

        submitButton = tk.Button(controlFrame, image=submit_img, text="Submit",command=lambda: uploadDf(conn, quoteDf ,latest_revised_quote),borderwidth=0, background=controlFrame["bg"])
        submitButton.image = submit_img
        submitButton.configure(state='disable')
        submitButton.grid(row=3,column=1)
        button_dict[submitButton] = [submit_img, submit_img2]
        submitButton.bind("<Enter>", on_enter)
        submitButton.bind("<Leave>", on_leave)
        
        
        def keyFinder2(dict, tupleValue):
            try:
                for key, value in dict.items():
                    for index in range(len(value[0])):
                        
                        if value[0][index][0] == tupleValue[0] and value[0][index][1]._name == tupleValue[1]:
                            return key,index
            except Exception as e:
                raise e 

        def handle_left_click(e):
            try:
                rowclicked_single = pt.get_row_clicked(e)
                print(f"Row clicked is {rowclicked_single+1}")
                if len(specialList):
                    if list(pt.model.df.columns) == ['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in','date_last_receipt','age', 'heat_number', 'lot_serial_number']:
                            varname = focused_entry.cget("textvariable")
                            focused_var = focused_entry.getvar(varname)
                            key, index = keyFinder2(specialList,(focused_entry,varname))
                            print(key, index)
                            specialList['Lot_Serial_Number'][0][index] = (pt.model.df['lot_serial_number'][rowclicked_single], None)
                            specialList['E_COST'][0][index][1].set(float(pt.model.df['onhand_dollars_per_pounds'][rowclicked_single]))

                            specialList['E_freightIncured'][0][index][1].set(0)
                            specialList['E_freightCharged'][0][index][1].set(0)
                            specialList['E_Margin_Freight'][0][index][1].set(0)
                            specialList['E_Additional_Cost'][0][index][1].set(0)
                pt.setSelectedRow(rowclicked_single)
                pt.redraw()
            except Exception as e:
                raise e

        

        # ent.bind('<FocusIn>',remember_focus)
        if pt is not None:
            pt.bind('<Button-1>',handle_left_click)
        # cxNamer.insert(tk.END, quotedf['CUS_NAME'][0])
        ##############Adding weight to mainFrames##############
        mainRowNum = 2
        mainColNum = 1
        # for i in range(mainRowNum+1):
        #     tab1.grid_rowconfigure(index=i,weight=1)

        tab1.grid_rowconfigure(0, weight=1) # For row 0
        tab1.grid_rowconfigure(1, weight=6) # For row 1
        tab1.grid_rowconfigure(2, weight=1) # For row 1
        for i in range(mainColNum+1):
            tab1.grid_columnconfigure(index=i,weight=1)
        # tab1.grid_columnconfigure(0, weight=1) # For column 0
        # tab1.grid_columnconfigure(1, weight=1) # For column 1

        #Configuring CxFrame grids as well as Controlgrid
        for i in range(5):
            cxFrame.grid_rowconfigure(index=i, weight=1)
            if i !=4:
                cxFrame.grid_columnconfigure(index=i, weight=1)
            #Configuring CxFrame2
            if i<3:
                cxFrame2.grid_rowconfigure(index=i, weight=1)
                cxFrame2.grid_columnconfigure(index=i, weight=1)
            #Configuring control Frame
            controlFrame.grid_rowconfigure(index=i, weight=1)
            if i <2:
                controlFrame.grid_columnconfigure(index=i, weight=1)
        # cxFrame.grid_rowconfigure(0, weight=1) # For row 0
        # cxFrame.grid_rowconfigure(1, weight=1) # For row 1
        # cxFrame.grid_rowconfigure(2, weight=1) # For row 2
        # cxFrame.grid_rowconfigure(3, weight=1) # For row 3
        # cxFrame.grid_rowconfigure(4, weight=1) # For row 4

        # cxFrame.grid_columnconfigure(0, weight=1) # For column 0
        # cxFrame.grid_columnconfigure(1, weight=1) # For column 1
        # cxFrame.grid_columnconfigure(2, weight=1) # For column 2
        # cxFrame.grid_columnconfigure(3, weight=1) # For column 3
        
        # cxFrame2.grid_rowconfigure(0, weight=1) # For row 0
        # cxFrame2.grid_rowconfigure(1, weight=1) # For row 1
        # cxFrame2.grid_rowconfigure(2, weight=1) # For row 1

        # cxFrame2.grid_columnconfigure(0, weight=1) # For column 0
        # cxFrame2.grid_columnconfigure(1, weight=1) # For column 1
        # cxFrame2.grid_columnconfigure(2, weight=1) # For column 1


        m_entryFrame.grid_rowconfigure(0, weight=1) # For row 0

        # # m_entryFrame.grid_rowconfigure(1, weight=1) # For row 1

        m_entryFrame.grid_columnconfigure(0, weight=1) # For column 0
        # # m_entryFrame.grid_columnconfigure(1, weight=1) # For column 1
        # databaseFrame.grid_rowconfigure(index=0,weight=1)
        # databaseFrame.grid_columnconfigure(index=1,weight=1)

        # bakerTableFrame.grid_rowconfigure(index=0, weight=1)
        # bakerTableFrame.grid_columnconfigure(index=1, weight=1)
        # entryFrame.grid_rowconfigure(0, weight=1) # For row 0
        # entryFrame.grid_rowconfigure(1, weight=1) # For row 1
        # entryFrame.grid_rowconfigure(2, weight=1) # For row 1

        # entryFrame.grid_columnconfigure(0, weight=1) # For column 0
        # entryFrame.grid_columnconfigure(1, weight=1) # For column 1
        # entryFrame.grid_columnconfigure(2, weight=1) # For column 2
        # entryFrame.grid_columnconfigure(3, weight=1) # For column 3
        # entryFrame.grid_columnconfigure(4, weight=1) # For column 4
        # entryFrame.grid_columnconfigure(5, weight=1) # For column 5
        # entryFrame.grid_columnconfigure(6, weight=1) # For column 6
        # entryFrame.grid_columnconfigure(7, weight=1) # For column 7
        # entryFrame.grid_columnconfigure(8, weight=1) # For column 8
        # entryFrame.grid_columnconfigure(9, weight=1) # For column 9
        # entryFrame.grid_columnconfigure(10, weight=1) # For column 10
        # entryFrame.grid_columnconfigure(11, weight=1) # For column 11
        # entryFrame.grid_columnconfigure(12, weight=1) # For column 12
        # entryFrame.grid_columnconfigure(13, weight=1) # For column 13
        # entryFrame.grid_columnconfigure(14, weight=1) # For column 14
        # entryFrame.grid_columnconfigure(15, weight=1) # For column 15
        # entryFrame.grid_columnconfigure(16, weight=1) # For column 16
        # entryFrame.grid_columnconfigure(17, weight=1) # For column 17
        # entryFrame.grid_columnconfigure(18, weight=1) # For column 18
        # entryFrame.grid_columnconfigure(19, weight=1) # For column 19
        # entryFrame.grid_columnconfigure(20, weight=1) # For column 20
        # entryFrame.grid_columnconfigure(21, weight=1) # For column 21
        
        # controlFrame.grid_rowconfigure(0, weight=1) # For column 21
        # controlFrame.grid_rowconfigure(1, weight=1) # For column 21       
        # controlFrame.grid_rowconfigure(2, weight=1) # For column 21       
        # controlFrame.grid_rowconfigure(3, weight=1) # For column 21
        # controlFrame.grid_columnconfigure(1, weight=1) # For column 21

        databaseFrame.grid_rowconfigure(1, weight=1) # For column 21
        databaseFrame.grid_columnconfigure(1, weight=1) # For column 21

        #Moving horizontal scroll bar to initial position
        entryCanvas.xview("moveto", 0)
        # df = get_inv_df(conn,table = INV_TABLE)
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
    
    root.mainloop()

# conn = get_connection()
# mainRoot = tk.Tk()
# user = "Imam"
# quotedf=None
# df = get_inv_df(conn,table = INV_TABLE)
# general_quote_revision(mainRoot, user, conn,quotedf,df)
# mainRoot.mainloop()
