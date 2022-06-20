from hmac import new
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.tix import ButtonBox

def myCombobox(df,root,item_list,frame,row,column,width=10,list_bd = 0,foreground='blue', background='white',sticky = tk.EW,boxList={}):
    # def __init__(self,item_list,frame,row,column,width=10,list_bd = 0,foreground='blue', background='white',sticky = tk.EW):
    global checker
    checker = True
    def OnMouseWheel(event):
        lbframe.list.yview_scroll(int(-1*(event.delta/220)), "units")
    max_len = len(max(item_list, key=len))
    if max_len > width:
        width = max_len
    var = tk.StringVar()
    ent = ttk.Entry(frame,textvariable=var, foreground=foreground,background=background,width=width)
    lst = tk.Listbox(frame, bd=list_bd, background=background,width=width,height=20)

    lbframe = tk.Frame(root)
    lbframe.list_values = tk.StringVar() 
    lbframe.list_values.set(item_list) 
    lbframe.list= tk.Listbox(lbframe, height=5, width= 10,
                                    listvariable=lbframe.list_values)
    lbframe.list.pack(side='left', fill="both", expand=True)
    s = ttk.Scrollbar(lbframe, orient=tk.VERTICAL, command=lbframe.list.yview)
    s.pack(side='right', fill = "y")
    lbframe.list['yscrollcommand'] = s.set
    lbframe.list.bind_all("<MouseWheel>", OnMouseWheel)



    def nextKey(key,i):
        keyList=boxList.keys()
        for k,v in enumerate(keyList):
            if v==key:
                newKey = list(keyList)[k+1]
        return newKey
        
    
    def entryUpdater(check,i,value,key, index):
        initValue = int(ent._name.split("y")[-1])
        ent._name = f'!entry{initValue+int(i//2)}'
        ent._w = f'.!toplevel.!frame3.!canvas.!frame.!entry{initValue+int(i//2)}'
        if check:
            ent.configure(state='disabled')
        else:
            ent.configure(state='normal')
        if value != "No" and value != "Yes":
            # if i ==2:
            #     key, index = keyFinder(boxList,(ent,var))
            
            newKey = nextKey(key,int(i//2))
            newList = filterList(boxList,newKey,index,df)
            # if value in newList:
            #     var.set(value)
            if len(newList)==1:
                var.set(newList[0])
            else:
                var.set("")
        ent._name = f'!entry{initValue}'
        ent._w = f'.!toplevel.!frame3.!canvas.!frame.!entry{initValue}'
            # ent.delete(0, tk.END)


    ent.grid(row=row,column=column,sticky=tk.EW,padx=5,pady=5)
    def list_hide(e=None):
        try:
            # global get_input_checker
            global checker
            checker = False
            if lbframe.list.size()==1:
                value = lbframe.list.get(0)
                var.set(value)
                #finding current key and index
                key, index = keyFinder(boxList,(ent,var))
                if value != "Yes" and value != "No" and key!='e_location':
                    current_key = key
                    while True:
                        next_key = nextKey(current_key, index)
                        new_list = filterList(boxList,next_key,index,df)
                        cx_eags = {'e_grade':'cx_grade','e_yield':'cx_yield','e_od':'cx_od','e_id':'cx_id'}
                        #if cx value in current list, set next key value = cx value
                        if str(boxList[cx_eags[next_key]][0][index].get()).upper() in list(map(lambda x: str(x).upper(),new_list)):
                            boxList[next_key][0][index][1].set(boxList[cx_eags[next_key]][0][index].get())
                        elif len(new_list)==1:
                            boxList[next_key][0][index][1].set(new_list[0])
                        
                        elif next_key == list(boxList.keys())[0]:
                            break
                        else:
                            break
                        current_key = next_key
                        
                elif value == "Yes" or value == "No":
                    keyIndex= list(boxList.keys()).index('e_location')
                    for i in range(keyIndex,len(list(boxList.keys()))):
                        newKey = list(boxList.keys())[i]
                        if value == "Yes":
                            boxList[newKey][0][index][1].set("")
                            boxList[newKey][0][index][0].configure(state='normal')
                        else:
                            boxList[newKey][0][index][1].set("NA")
                            boxList[newKey][0][index][0].configure(state='disabled')
                    
            
                                    


                        
                        

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
                
        except:
            pass
        lbframe.list.delete(0, tk.END)
        lbframe.place_forget()

    def list_input(_):
        if ent.get()=='':
            key, index = keyFinder(boxList,(ent,var))
            newList = filterList(boxList,key,index,df)
            lbframe.list.delete(0, tk.END)
            for item in newList:
                lbframe.list.insert(tk.END, item)
                lbframe.list.itemconfigure(tk.END, foreground="black")
            lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
        lbframe.list.focus()
        lbframe.list.select_set(0)

    def list_up(_):
        if not lbframe.list.curselection()[0]:
            ent.focus()
            list_hide()


    def get_selection(_):
        value = lbframe.list.get(lbframe.list.curselection())
        var.set(value)
        list_hide()
        ent.focus()
        ent.icursor(tk.END)
    
    
    ent.bind('<Down>', list_input)
    ent.bind('<Return>', list_hide)
    ent.bind('<Tab>',list_hide)

    lbframe.list.bind('<Up>', list_up)
    lbframe.list.bind('<Return>', get_selection)
    lbframe.list.bind('<Double-Button-1>', get_selection)
    lbframe.list.bind('<Tab>',list_hide)
    lbframe.list.bind('<Escape>',list_hide)
    # ent.bind('<Down>', list_input)
    # # ent.bind('<Return>', list_hide)

    # lbframe.list.bind('<Up>', list_up)
    # lbframe.list.bind('<Return>', get_selection)

    # return ent,lbframe.list,var

    def keyFinder(dict, tupleValue):
        
        for key, value in dict.items():
            for index in range(len(value[0])):
                
                if value[0][index] == tupleValue:
                    return key,index
    def filterList(boxList,key,index,df):
        newList = []
        if key == "e_type":
            #filter df based on e_location and make unique column of e_type
            new_df = df[(df["Location"] == boxList["e_location"][0][index][0].get())]
            newList = list(new_df["Type"].unique())

        elif key == "e_grade":
            #filter df based on e_location and make unique column of e_type
            new_df = df[(df["Location"] == boxList["e_location"][0][index][0].get())& (df["Type"]==boxList["e_type"][0][index][0].get())]
            newList = list(new_df["Grade"].unique())
        elif key == "e_yield":
            #filter df based on e_location and make unique column of e_type
            new_df = df[(df["Location"] == boxList["e_location"][0][index][0].get())& (df["Type"]==boxList["e_type"][0][index][0].get())
             & (df["Grade"]==boxList["e_grade"][0][index][0].get())]
            newList = list(new_df["Yield"].unique())
        elif key == "e_od":
            #filter df based on e_location and make unique column of e_type
            new_df = df[(df["Location"] == boxList["e_location"][0][index][0].get())& (df["Type"]==boxList["e_type"][0][index][0].get())
             & (df["Grade"]==boxList["e_grade"][0][index][0].get())& (df["Yield"]==boxList["e_yield"][0][index][0].get())]
            newList = list(new_df["OD"].unique())
        elif key == "e_id":
            #filter df based on e_location and make unique column of e_type
            new_df = df[(df["Location"] == boxList["e_location"][0][index][0].get())& (df["Type"]==boxList["e_type"][0][index][0].get())
             & (df["Grade"]==boxList["e_grade"][0][index][0].get())& (df["Yield"]==boxList["e_yield"][0][index][0].get())
             & (df["OD"]==float(boxList["e_od"][0][index][0].get()))]
            newList = list(new_df["ID"].unique())
        try:
            newList.sort()
        except:
            pass
        return newList




    def get_input(*args):
        if checker:
            newList = item_list
            check = True
            # print(boxList)
            if len(boxList):# and len(lbframe.list.get(0,tk.END))!=1
                # boxList["quoteYesNo"]
                key, index = keyFinder(boxList,(ent,var))
                # print(key, index)
                if key != "quoteYesNo" and key != "e_location" and boxList["e_location"][0][index][0].get()!='':
                    if not str(boxList["quoteYesNo"][0][index][0].get().upper()).startswith("Y"):
                        check = False
                    else:
                        newList = filterList(boxList,key,index,df)
            
            # lbframe.place(in_=ent, x=0, rely=1, relwidth=1.0, anchor="nw")
            
            lbframe.list.delete(0, tk.END)
            if len(newList)==1:
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
                        if not str(boxList["quoteYesNo"][0][index][0].get()).upper().startswith("N"):
                            messagebox.showerror(title="Wrong Value",message="Please enter value from list only!")
                            ent.delete(0, tk.END)
                            ent.focus()

                else:
                    lbframe.place_forget()
                    if string!='' and check:
                        messagebox.showerror(title="Wrong Value",message="Please enter value from list only!222")
                        ent.focus()
                    
                # lst.grid_remove()

    var.trace('w', get_input)

    return (ent,var)
    # ent.bind("<Button-1>", get_input)
    # ent.bind('<Down>', list_input)
    # ent.bind('<Return>', list_hide)

    # lst.bind('<Up>', list_up)
    # lst.bind('<Return>', get_selection)
    
# def mylistBox(ent, item_list, lst, var):
#     var.trace('w', get_input(ent, item_list, lst, var))
    