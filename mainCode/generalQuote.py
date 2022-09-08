from re import search
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from customComboboxV2 import myCombobox
from tkcalendar import DateEntry
from datetime import date
import sys
import pandas as pd
from pandastable import Table
from Tools import dfMaker, resource_path, rangeSearch
from sfTool import get_connection,get_cx_df, get_inv_df
from final_pdf_creator import pdf_generator
from sfTool import eagsQuotationuploader
import os, shutil
from tkPDFViewer import tkPDFViewer as pdf
import ctypes

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





def quoteGenerator(mainRoot,user,conn, df):
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
                # entryCanvas.yview_moveto('1.0')
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
                cx_spec[-1][0].grid(row=2+row_num,column=0,padx=(10,0))

                cx_type.append((None, None))
                # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                cx_grade.append((ttk.Entry(entryFrame,width=10),None))
                cx_grade[-1][0].grid(row=2+row_num,column=1)
                # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
                cx_yield.append((ttk.Entry(entryFrame,width=10),None))
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
                quoteYesNo.append(myCombobox(df,tab1,item_list=["Yes","No","Other"],frame=entryFrame,row=2+row_num,column=7,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # quoteYesNo[-1]['validate']='focusout'
                # quoteYesNo[-1]['validatecommand'] = (quoteYesNo[-1].register(yesNo),'%P','%W')
                e_location.append(myCombobox(df,tab1,item_list=["Dubai","Singapore","USA","UK"],frame=entryFrame,row=2+row_num,column=8,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # e_location[-1].config(textvariable="NA", state='disabled')
                e_type.append(myCombobox(df,tab1,item_list=["THF","BR", "TUI", "HR"],frame=entryFrame,row=2+row_num,column=9,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))

                e_spec.append((None, None))
                # e_type[-1].config(textvariable="NA", state='disabled')
                e_grade.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=10,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # e_grade[-1].config(textvariable="NA", state='disabled')
                e_yield.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=11,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # e_yield[-1].config(textvariable="NA", state='disabled')
                e_od1.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=12,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                # e_od[-1].config(textvariable="NA", state='disabled')
                e_od1[-1][0]['validate']='key'
                e_od1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')
                
                e_id1.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=13,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt))
                e_id1[-1][0]['validate']='key'
                e_id1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')

                e_od2.append((None, None))
                e_id2.append((None, None))
                # searchGrade.append((None, None))
                # searchYield.append((None, None))
                searchLocation.append((None, None))

                # e_id[-1].config(textvariable="NA", state='disabled')
                e_len.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=14,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                e_len[-1][0]['validate']='key'
                e_len[-1][0]['validatecommand'] = (e_len[-1][0].register(intFloat),'%P','%d')
                # e_len[-1].config(textvariable="NA", state='disabled')
                e_qty.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=15,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                e_qty[-1][0]['validate']='key'
                e_qty[-1][0]['validatecommand'] = (e_qty[-1][0].register(intChecker),'%P','%d')
                # e_qty[-1].config(textvariable="NA", state='disabled')

                e_cost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=16,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                e_cost[-1][0]['validate']='key'
                e_cost[-1][0]['validatecommand'] = (e_cost[-1][0].register(intFloat),'%P','%d')

                sellCostLBS.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=17,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                sellCostLBS[-1][0]['validate']='key'
                sellCostLBS[-1][0]['validatecommand'] = (sellCostLBS[-1][0].register(intFloat),'%P','%d')

                marginlbs.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=18,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                # sellCostLBS[-1].config(textvariable="NA", state='disabled')
                e_uom.append(myCombobox(df,tab1,item_list=["Inch","Each","Foot"],frame=entryFrame,row=2+row_num,column=19,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                # e_uom[-1].config(textvariable="NA", state='disabled')
                sellCostUOM.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=20,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                sellCostUOM[-1][0]['validate']='key'
                sellCostUOM[-1][0]['validatecommand'] = (sellCostUOM[-1][0].register(intFloat),'%P','%d')
                # sellCostUOM[-1].config(textvariable="NA", state='disabled')
                addCost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=21,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                addCost[-1][0]['validate']='key'
                addCost[-1][0]['validatecommand'] = (addCost[-1][0].register(intFloat),'%P','%d')
                # addCost[-1].config(textvariable="NA", state='disabled')
                leadTime.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=22,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                # leadTime[-1].config(textvariable="NA", state='disabled')

                finalCost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=23,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                finalCost[-1][0]['validate']='key'
                finalCost[-1][0]['validatecommand'] = (finalCost[-1][0].register(intFloat),'%P','%d')
                

                freightIncured.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=24,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                
                freightCharged.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=25,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))



                

                marginFreight.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=26,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))

                lot_serial_number.append((None, None))


            except Exception as e:
                raise e

        def cxListCalc():
            try:
                cxList = [cxDatadict["Prepared_By"],cxDatadict["Date"],cxDatadict["cus_long_name"][0][0][0].get(), cxDatadict["payment_term"][0][0].get(), currency.get(),  cxDatadict["cus_address"][0][0].get(),
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
                root.update()
                entryCanvas.yview("moveto", 0)

                
                # entryCanvas.yview_moveto('1.0')
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
                    # try:
                    #     pdfRoot.destroy()
                    # except:
                    #     pass
                    pdfRoot = tk.Toplevel()
                    pdfRoot.title(quoteDf["QUOTENO"][0])
                    pdfviewer = pdf.ShowPdf()
                    # zoom_scale = tk.Scale(pdfRoot, orient='vertical', from_=1, to=500)
                    # zoom_scale.config(command=zoom)
                    # Adding pdf path and width and height.
                    # zoom_scale.pack(fill='y', side='right')
                    # zoom_scale.set(10)
                    screen_width = (pdfRoot.winfo_screenwidth())//6
                    screen_height = (pdfRoot.winfo_screenheight())//6
                    pdfframe = pdfviewer.pdf_view(pdfRoot, pdf_location=pdf_path, width=120 ,zoomDPI=100)
                    pdfframe.pack(expand=True, fill='both')
                    pdfRoot.state('zoomed')
                    submitButton.configure(state='normal')
                else:
                    messagebox.showerror("Error", "Empty dataframe was given in input")
            except Exception as e:
                raise e

        def uploadDf(conn, quoteDf):
            try:
                # pt.model.df = quoteDf
                # pt.redraw()
                submitButton.configure(state='disable')
                if messagebox.askyesno("Upload to Database", "Are sure that you want to generate quote and upload Data?"):
                    eagsQuotationuploader(conn, quoteDf, latest_revised_quote=None)
                    
                    messagebox.showinfo("Info", "Data uploaded Successfully!")

                    # current_work_dir = os.getcwd() #Should be Shared Drive
                    current_work_dir = r'I:\EAGS\Quotes'
                    cx_init_name = str(quoteDf['QUOTENO'][0]).split("_")[0]
                    filename = str(quoteDf['QUOTENO'][0])+".pdf"
                    # save_dir = current_work_dir+"\\"+cx_init_name
                    # if not os.path.exists(save_dir):
                    #     os.mkdir(save_dir)
                    # # os.rename(pdf_path,save_dir+"\\"+filename)
                    # shutil.move(pdf_path,save_dir+"\\"+filename)
                    desktopDir = os.path.join(os.environ["HOMEPATH"], "Desktop\\EAGS_Quotes")
                    desktopDir = os.path.join('C:', desktopDir)
                    if not os.path.exists(desktopDir):
                        os.mkdir(desktopDir)
                    # directories_created = [desktopDir]
                    # for directory in directories_created:
                    #     path3 = os.path.join(os.getcwd(),directory)  
                    #     try:
                    #         os.makedirs(path3, exist_ok = True)
                    #         print("Directory '%s' created successfully" % directory)
                    #     except OSError as error:
                    #         print("Directory '%s' can not be created" % directory) 

                    # if not os.path.exists(desktopDir):
                    #     os.mkdir(desktopDir)
                    shpUploader(pdf_path,filename)
                    shutil.move(pdf_path,desktopDir+"\\"+filename)
                    
                    # shutil.copy(save_dir+"\\"+filename, desktopDir)
                    # shutil.copy(save_dir+"\\"+filename, desktopDir)
                    
                else:
                    os.remove(pdf_path)
                
            except Exception as e:
                raise e

        
        
        
        mainRoot.withdraw()
        global row_num
        row_num=0

        #Getting invoentory dataframe
        # df = get_inv_df(conn,table = INV_TABLE)
        
        
        # df = pd.read_excel("sampleInventory.xlsx")
        # Getting Cx Dataframe
        cx_df = get_cx_df(conn,table = CX_TABLE)
        # cx_df = pd.read_excel("cxDatabase.xlsx")

        count = 0
        root = tk.Toplevel(mainRoot, bg = "#9BC2E6")
        root.state('zoomed')
        root.title('EAGS Quote Generator')
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

        #Frames Under Tab1
        cxFrame.grid(row=0, column=0,pady=(24,0), padx=(30,0),sticky="nsew")
        cxFrame2.grid(row=0, column=1,pady=(24,0), padx=(30,40),sticky="nsew")
        # headerFrame.grid(row=1, column=0,sticky="sew",columnspan=2)
        m_entryFrame.grid(row=1, column=0,sticky="nsew", columnspan=2)
        xscrollbar.grid(row=1,column=0,sticky=tk.NSEW)
        yscrollbar.grid(row=0,column=1,sticky=tk.NSEW)
        entryCanvas.grid(row=0,column=0, sticky=tk.NSEW)
        databaseFrame.grid(row=2,column=0, sticky="nsew")
        controlFrame.grid(row=2,column=1, sticky="nsew")
        
        entryCanvas.create_window((0,0),window=entryFrame,tags='expand')
        
        
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
        nonList = [[None,None,None,None,None, None, None, None, None]]
        # pandasDf = pd.DataFrame(nonList,columns=['onhand_pieces', 'onhand_length_in', 'reserved_pieces', 'reserved_length_in', 'available_pieces', 'available_length_in'])
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
        cxDatadict["currency"] = []


        item_list = () #('A4140', 'A4140M', 'A4330V', 'A4715', 'BS708M40', 'A4145M', '4542','4462')

        cxLabel = tk.Label(cxFrame, text="Customer Details", bg = "#9BC2E6", font=("Segoe UI", 12))
        prepByLb = tk.Label(cxFrame,text="Prepared By", bg = "#9BC2E6", font=("Segoe UI", 10))
        prep_by = ttk.Entry(cxFrame)
        prep_by.insert(tk.END, user)
        inpDateLb = tk.Label(cxFrame,text="Date", bg = "#9BC2E6", font=("Segoe UI", 10))
        inpDate = MyDateEntry(master=cxFrame, width=17, selectmode='day', font=("Segoe UI", 10))
        cxNameLb = tk.Label(cxFrame,text="Customer Name", bg = "#9BC2E6", font=("Segoe UI", 10))
        locAddLb = tk.Label(cxFrame,text="Location/Address", bg = "#9BC2E6", font=("Segoe UI", 10))
        emailLb = tk.Label(cxFrame,text="Email", bg = "#9BC2E6", font=("Segoe UI", 10))
        payTermLb = tk.Label(cxFrame,text="Payment Terms", bg = "#9BC2E6", font=("Segoe UI", 10))

        #Adding Search button in cxFrame 2
        # starButton = tk.Button(cxFrame2, text="Star Search", font = ("Segoe UI", 10, 'bold'), bg="#20bebe", fg="white", height=1, width=14, command=lambda: starSearch(root, df), activebackground="#20bebb", highlightbackground="#20bebd")
        rangeButton = tk.Button(cxFrame2, text="Range Search", font = ("Segoe UI", 10, 'bold'), bg="#20bebe", fg="white", height=1, width=14, command=lambda: rangeSearch(root, df, specialList, 0), activebackground="#20bebb", highlightbackground="#20bebd")

        mobileLb = tk.Label(cxFrame2, text="Mobile", bg = "#9BC2E6", font=("Segoe UI", 10))
        currencyLabel = tk.Label(cxFrame2,text="Currency", bg = "#9BC2E6", font=("Segoe UI", 10))
        valLb = tk.Label(cxFrame2,text="Validity", bg = "#9BC2E6", font=("Segoe UI", 10))
        addCommLb = tk.Label(cxFrame2,text="Additional Comments", bg = "#9BC2E6", font=("Segoe UI", 10))
        # remarksLabel = tk.Label(cxFrame2,text="Remarks", bg = "#9BC2E6", font=("Segoe UI", 10))
        cxLabel.grid(row=0,column=0)
        prepByLb.grid(row=1,column=0)
        prep_by.grid(row=2,column=0)
        inpDateLb.grid(row=1,column=1)
        inpDate.grid(row=2, column=1)
        cxNameLb.grid(row=3,column=0)
        locAddLb.grid(row=3,column=1)
        emailLb.grid(row=3,column=2)
        payTermLb.grid(row=3,column=3)

        #label grid using cxFrame2

        #Adding Search button in cxFrame 2
        rangeButton.grid(row=0, column=0, pady=(20,0))
        mobileLb.grid(row=1, column=0, pady=(20,0))
        # rangeButton.grid(row=0,column=1, pady=(20,0))
        currencyLabel.grid(row=1,column=1, pady=(20,0))
        # remarksLabel.grid(row=1,column=2, pady=(20,0))#padx=(50,5)
        valLb.grid(row=1,column=2, pady=(20,0))#5x=(50,5)
        addCommLb.grid(row=1,column=3, pady=(20,0))#padx=(50,5)

        
        
        
        prep_by.config(state= "disabled")
        cxDatadict["Prepared_By"] = prep_by.get()

        
        
        
        cxDatadict["Date"] = inpDate.get()
        
        #Currency
        currencyVar = tk.StringVar()
        currency = ttk.Combobox(cxFrame2, background='white', font=('Segoe UI', 10), justify='center',textvariable=currencyVar,values=["$","£"], width=5, text='$')
    
        # currencyVar.set("$")
        currency.delete(0, tk.END)
        currency.insert(tk.END,'$')
        # currency = ttk.Entry(cxFrame2, textvariable=currencyVar, foreground='blue', background = 'white',width = 10, font=('Segoe UI', 10))
        currency.grid(row=2,column=1,pady=5)
        # #Remarks
        # remarksVar = tk.StringVar()
        # remarks = ttk.Entry(cxFrame2, textvariable=remarksVar, foreground='blue', background = 'white',width = 40, font=('Segoe UI', 10))
        # remarks.grid(row=2,column=2,pady=5)
        
        #Validity
        validityVar = tk.StringVar()
        validity = ttk.Entry(cxFrame2, textvariable=validityVar, foreground='blue', background = 'white',width = 20, font=('Segoe UI', 10))
        validity.grid(row=2,column=2,pady=5)
        
        #Additional Comments
        addCommVar = tk.StringVar()
        addComm = ttk.Entry(cxFrame2, textvariable=addCommVar, foreground='blue', background = 'white',width = 40, font=('Segoe UI', 10))
        addComm.grid(row=2,column=3,sticky=tk.EW,pady=5)

        addComm.bind("<Tab>",tabFunc)
        
        


        #Customer Name Entry Box
        cxNameVar.append(myCombobox(cx_df,tab1,item_list=list(cx_df['cus_long_name']),frame=cxFrame,row=4,column=0,width=25,list_bd = 0,foreground='blue', background='white',sticky = "nsew",cxDict= cxDatadict,val=currency))
        #location Address entry box
        locAddVar = tk.StringVar()
        locAdd = ttk.Entry(cxFrame, textvariable=locAddVar, foreground='blue', background = 'white',width = 20, font=('Segoe UI', 10))
        locAdd.grid(row=4,column=1,sticky=tk.EW,padx=5,pady=5)
        # cxLocVar = []
        cxDatadict["cus_address"].append((locAdd, locAddVar))
        # myCombobox(df,root,item_list,frame=cxFrame,row=4,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")

        #Email
        emailAddVar = tk.StringVar()
        emailAdd = ttk.Entry(cxFrame, textvariable=emailAddVar, foreground='blue', background = 'white',width = 20, font=('Segoe UI', 10))
        emailAdd.grid(row=4,column=2,sticky=tk.EW,padx=5,pady=5)
        # cxemailAddVar = []
        cxDatadict["cus_email"].append((emailAdd, emailAddVar))
        # myCombobox(df,root,item_list,frame=cxFrame,row=4,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew")

        #Payment Terms Entry
        payTermVar = tk.StringVar()
        payTerm = ttk.Entry(cxFrame, textvariable=payTermVar, foreground='blue', background = 'white',width = 5, font=('Segoe UI', 10))
        payTerm.grid(row=4,column=3,sticky=tk.EW,padx=5,pady=5)
        cxDatadict["payment_term"].append((payTerm, payTermVar))
        
        #Mobile No. Entry
        mobileVar = tk.StringVar()
        mobile = ttk.Entry(cxFrame2, textvariable=mobileVar, foreground='blue', background = 'white',width = 20, font=('Segoe UI', 10))
        mobile.grid(row=2,column=0,sticky=tk.EW,padx=5,pady=5)
        cxDatadict["cus_phone"].append((mobile, mobileVar))
        

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
        addCostLabel2 = tk.Label(entryFrame, text="Cost/UOM", bg= "#DDEBF7")
        leadTimeLabel = tk.Label(entryFrame, text="Lead Time", bg= "#DDEBF7")

        finalPriceLabel = tk.Label(entryFrame, text="Final Price/UOM", bg= "#DDEBF7")

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
        leadTimeLabel.grid(row=0,column=22, sticky="ew")

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

        # searchGrade = []
        # specialList["searchGrade"] = []
        # specialList["searchGrade"].append(searchGrade)

        # searchYield = []
        # specialList["searchYield"] = []
        # specialList["searchYield"].append(searchYield)

        

        # specialList = [[quoteYesNo],[e_location], [e_type], [e_grade], [e_yield], [e_od], [e_id], [e_len], [e_qty], [sellCostLBS], [sellCostUOM],
        # [e_uom], [addCost], [leadTime], [finalCost]]
        ###########################################################################################
        
        # var = tk.StringVar()
        # spec = ttk.Entry(entryFrame,textvariable=var, foreground='blue',background='white',width=5)
        # spec.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        # myCombobox(df,root,cx_list,frame=entryFrame,row=1,column=0,width=2,list_bd = 0,foreground='blue', background='white',sticky = "nsew")
        
        while len(quoteYesNo)<1:
            addRow()
        
        # button_dict = {}

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
            if i==3:
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
        root.update()
        entryCanvas.xview("moveto", 0)

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
# # df = pd.read_excel("sampleInventory.xlsx")
# df = get_inv_df(conn,table = INV_TABLE)
# quoteGenerator(mainRoot, user, conn,df)
# mainRoot.mainloop()
