import tkinter as tk
from tkinter import messagebox
from tkinter import ttk, Canvas
from customComboboxV2 import myCombobox
from tkcalendar import DateEntry
from datetime import date
import sys
import pandas as pd
from pandastable import Table
from Tools import dfMaker, resource_path, bakerMaker, starSearch, rangeSearch
from sfTool import get_connection,get_cx_df, get_inv_df, get_salesperson_df
from final_pdf_creator import pdf_generator
from sfTool import eagsQuotationuploader
import os
from pandasPaste import pandasPaste
from pandastable.headers import ColumnHeader, RowHeader, IndexHeader
from pandastable.dialogs import *
from pandastable.core import *
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(2)
import shutil
from shpUploader import shpUploader

UNITS = "units"
INV_TABLE = "EAGS_INVENTORY"
CX_TABLE = "EAGS_CUSTOMER"
S_PERSON_TABLE = "EAGS_SALESPERSON_V2"

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


class ResizingCanvas(Canvas):
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
class myIndexHeader(IndexHeader):
    """Class that displays the row index headings."""

    def __init__(self, parent=None, table=None, width=40, height=25, bg='gray50'):
        Canvas.__init__(self, parent, bg=bg, width=width, height=height)
        if table != None:
            subtractor=0
            if parent.winfo_pixels('1i')==96:
                subtractor=10
            self.table = table
            self.width = width
            self.height = 18+30-subtractor
            self.config(height=self.height)
            self.textcolor = 'white'
            self.bgcolor = bg
            self.startrow = self.endrow = None
            self.model = self.table.model
            self.bind('<Button-1>',self.handle_left_click)
            return

class myToolBar(ToolBar):
    """Uses the parent instance to provide the functions"""
    def __init__(self, parent=None, parentapp=None):

        Frame.__init__(self, parent, width=600, height=40)
        self.parentframe = parent
        self.parentapp = parentapp
        img = images.open_proj()
        # addButton(self, 'Load table', self.parentapp.load, img, 'load table')
        # img = images.save_proj()
        # addButton(self, 'Save', self.parentapp.save, img, 'save')
        # img = images.importcsv()
        # func = lambda: self.parentapp.importCSV(dialog=1)
        # addButton(self, 'Import', func, img, 'import csv')
        # img = images.excel()
        # addButton(self, 'Load excel', self.parentapp.loadExcel, img, 'load excel file')
        img = images.copy()
        addButton(self, 'Copy', self.parentapp.copyTable, img, 'copy table to clipboard')
        img = images.paste()
        addButton(self, 'Paste', self.parentapp.pasteTable, img, 'paste table')
        # img = images.plot()
        # addButton(self, 'Plot', self.parentapp.plotSelected, img, 'plot selected')
        # img = images.transpose()
        # addButton(self, 'Transpose', self.parentapp.transpose, img, 'transpose')
        # img = images.aggregate()
        # addButton(self, 'Aggregate', self.parentapp.aggregate, img, 'aggregate')
        # img = images.pivot()
        # addButton(self, 'Pivot', self.parentapp.pivot, img, 'pivot')
        # img = images.melt()
        # addButton(self, 'Melt', self.parentapp.melt, img, 'melt')
        # img = images.merge()
        # addButton(self, 'Merge', self.parentapp.doCombine, img, 'merge, concat or join')
        # img = images.table_multiple()
        # addButton(self, 'Table from selection', self.parentapp.tableFromSelection,
        #             img, 'sub-table from selection')
        img = images.filtering()
        addButton(self, 'Query', self.parentapp.queryBar, img, 'filter table')
        # img = images.calculate()
        # addButton(self, 'Evaluate function', self.parentapp.evalBar, img, 'calculate')
        # img = images.fit()
        # addButton(self, 'Stats models', self.parentapp.statsViewer, img, 'model fitting')

        # img = images.table_delete()
        # addButton(self, 'Clear', self.parentapp.clearTable, img, 'clear table')
        #img = images.prefs()
        #addButton(self, 'Prefs', self.parentapp.showPrefs, img, 'table preferences')
        return



class myFilterBar(FilterBar):
    """Class providing filter widgets"""

    operators = ['contains','excludes','equals','not equals','>','<','is empty','not empty',
                 'starts with','ends with','has length','is number','is lowercase','is uppercase']
    booleanops = ['AND','OR','NOT']

    def __init__(self, parent, parentframe, cols,df):

        Frame.__init__(self, parentframe)
        self.parent = parent
        self.filtercol = StringVar()
        initial = cols[0]
        self.filtercolmenu = Combobox(self,
                textvariable = self.filtercol,
                values = cols,
                #initialitem = initial,
                width = 14)
        self.filtercolmenu.grid(row=0,column=1,sticky='news',padx=2,pady=2)
        self.operator = StringVar()
        #self.operator.set('equals')
        # operatormenu = Combobox(self,
        #         textvariable = self.operator,
        #         values = self.operators,
        #         width = 10)
        # operatormenu.grid(row=0,column=2,sticky='news',padx=2,pady=2)
        self.filtercolvalue=StringVar()
        # valsbox = Entry(self,textvariable=self.filtercolvalue,width=26)
        
        valsbox = Combobox(self,
                textvariable = self.filtercolvalue,
                values = df[self.filtercol.get()] if self.filtercol.get()!='' else [''],
                #initialitem = initial,
                width = 14)
        valsbox.grid(row=0,column=3,sticky='news',padx=2,pady=2)

        #handling value update in 1st entry box
        def updateValue(*args):
            valsbox['values'] = list(df[self.filtercol.get()]) if self.filtercol.get()!='' else ['']
        self.filtercol.trace('w', updateValue)

        
        
        #valsbox.bind("<Return>", self.parent.callback)
        # self.booleanop = StringVar()
        # self.booleanop.set('AND')
        # booleanopmenu = Combobox(self,
        #         textvariable = self.booleanop,
        #         values = self.booleanops,
        #         width = 6)
        # booleanopmenu.grid(row=0,column=0,sticky='news',padx=2,pady=2)
        #disable the boolean operator if it's the first filter
        #if self.index == 0:
        #    booleanopmenu.component('menubutton').configure(state=DISABLED)
        img = images.cross()
        cb = Button(self,text='-', image=img, command=self.close)
        cb.image = img
        cb.grid(row=0,column=5,sticky='news',padx=2,pady=2)
        return


    def getFilter(self):
        """Get filter values for this instance"""

        col = self.filtercol.get()
        val = self.filtercolvalue.get()
        try:
            int(val)
            op = 'equals'
        except:
            try:
                float(val)
                op = 'equals'
            except:
                op = 'contains'#self.operator.get()
        booleanop = 'AND'#self.booleanop.get()
        return col, val, op, booleanop


def bakerQuoteGenerator(mainRoot,user,conn, df):
    try:
        class myQueryDialog(QueryDialog):
            """Use string query to filter. Will not work with spaces in column
                names, so these would need to be converted first."""

            def __init__(self, table):
                parent = table.parentframe
                Frame.__init__(self, parent)
                self.parent = parent
                self.table = table
                self.setup()
                self.filters = []
                return

            def setup(self):

                qf = self
                sfont = "Helvetica 10 bold"
                self.fbar = Frame(qf)
                self.fbar.pack(side=TOP,fill=BOTH,expand=1,padx=2,pady=2)
                f = Frame(qf)
                f.pack(side=TOP, fill=BOTH, padx=2, pady=2)
                self.addFilter
                addButton(f, 'find', self.query, images.filtering(), 'apply filters', side=LEFT)
                addButton(f, 'add manual filter', self.addFilter, images.add(),
                        'add manual filter', side=LEFT)
                addButton(f, 'close', self.close, images.cross(), 'close', side=LEFT)
                self.applyqueryvar = BooleanVar()
                self.applyqueryvar.set(True)
                c = Checkbutton(f, text='show filtered only', variable=self.applyqueryvar,
                            command=self.query)
                c.pack(side=LEFT,padx=2)
                # addButton(f, 'color rows', self.colorResult, images.color_swatch(), 'color filtered rows', side=LEFT)

                self.queryresultvar = StringVar()
                l = Label(f,textvariable=self.queryresultvar, font=sfont)
                l.pack(side=RIGHT)
                return

            def addFilter(self):
                """Add a filter using widgets"""

                df = self.table.model.df
                fb = myFilterBar(self, self.fbar, list(df.columns),df)
                fb.pack(side=TOP, fill=BOTH, expand=1, padx=2, pady=2)
                self.filters.append(fb)
                return
            
            def close(self):
                self.destroy()
                self.table.qframe = None
                self.table.showAll()


            def query(self, evt=None):
                """Do query"""
                global temp_bakerDf
                table = self.table
                #saving current values in temporary df and update main baker dataframe
                # print(temp_bakerDf)
                tempbakerDfUpdate()
                s = ''#self.queryvar.get()
                if table.filtered == True:
                    table.model.df = table.dataframe
                # df = table.model.df
                #instead of taking df from pandas table we will pick it from current temp baker df
                df = temp_bakerDf.copy()
                #Declaring prev df length variable
                prev_len = len(df)
                mask = None

                #string query first
                if s!='':
                    try:
                        mask = df.eval(s)
                    except:
                        mask = df.eval(s, engine='python')
                #add any filters from widgets
                if len(self.filters)>0:
                    mask = self.applyFilter(df, mask)
                if mask is None:
                    table.showAll()
                    self.queryresultvar.set('')
                    return
                #apply the final mask
                self.filtdf = filtdf = df[mask]
                self.queryresultvar.set('%s rows found' %len(filtdf))

                if self.applyqueryvar.get() == 1:
                    #replace current dataframe but keep a copy!
                    table.dataframe = table.model.df.copy()
                    table.delete('rowrect')
                    table.multiplerowlist = []
                    table.model.df = filtdf[table.model.df.columns.to_list()]#using columns that were earlier print in pandas table
                    table.filtered = True
                else:
                    idx = filtdf.index
                    rows = table.multiplerowlist = table.getRowsFromIndex(idx)
                    if len(rows)>0:
                        table.currentrow = rows[0]

                table.redraw()
                #updating temp baker dataframe
                
                temp_bakerDf = filtdf.copy()
                # temp_bakerDf["C_Quote Yes/No"], temp_bakerDf["E_Location"], temp_bakerDf["E_Type"],temp_bakerDf["E_Spec"], temp_bakerDf["E_Grade"], temp_bakerDf["E_Yield"], temp_bakerDf["E_OD1"], temp_bakerDf["E_ID1"] = [None, None, None, None, None, None, None, None]
                # temp_bakerDf["E_OD2"], temp_bakerDf["E_ID2"] = [None, None]
                # temp_bakerDf["E_Length"], temp_bakerDf["E_Qty"], temp_bakerDf["E_Selling Cost/LBS"], temp_bakerDf["E_UOM"], temp_bakerDf["E_Selling Cost/UOM"], temp_bakerDf["E_Additional_Cost"] = [None, None, None, None, None, None]
                # temp_bakerDf["E_LeadTime"], temp_bakerDf["E_Final Price"] = [None, None]
                #deleting exxtra entry boxs
                curr_len = len(filtdf)
                # while curr_len < prev_len:
                #     deleteRow()
                #     prev_len -= 1
                ###############check this
                #Removing grid of those rows which are not present in temporary dataframe
                for row_number in range(len(specialList['C_Quote Yes/No'])):
                    if row_number not in list(temp_bakerDf.index):
                        keyIndex= list(specialList.keys()).index('C_Quote Yes/No')
                        for i in range(keyIndex,len(list(specialList.keys()))):
                            newKey = list(specialList.keys())[i]
                            print(newKey)
                            if newKey!='E_OD2' and newKey != 'E_ID2':
                                specialList[newKey][0][row_number][0].grid_remove()
                # row_number=0
                # for index in temp_bakerDf.index:
                #     # print(index)
                #     entryBoxUpdater(row_number, index_num= index)
                #     row_number+=1
                print('While loop completed')
                return
        #Custom ColHeader
        class MyTable(Table):
            # based on original drawCellEntry() with required changes
            def show(self, callback=None):
                """Adds column header and scrollbars and combines them with
                the current table adding all to the master frame provided in constructor.
                Table is then redrawn."""

            
                #Add the table and header to the frame
                self.rowheader = RowHeader(self.parentframe, self)
                self.colheader = ColumnHeader(self.parentframe, self, bg='gray25')
                self.rowindexheader = myIndexHeader(self.parentframe, self, bg='gray75')
                self.Yscrollbar = AutoScrollbar(self.parentframe,orient=VERTICAL,command=self.set_yviews)
                self.Yscrollbar.grid(row=1,column=2,rowspan=1,sticky='news',pady=0,ipady=0)
                self.Xscrollbar = AutoScrollbar(self.parentframe,orient=HORIZONTAL,command=self.set_xviews)
                self.Xscrollbar.grid(row=2,column=1,columnspan=1,sticky='news')
                self['xscrollcommand'] = self.Xscrollbar.set
                self['yscrollcommand'] = self.Yscrollbar.set
                self.colheader['xscrollcommand'] = self.Xscrollbar.set
                self.rowheader['yscrollcommand'] = self.Yscrollbar.set
                self.parentframe.rowconfigure(1,weight=1)
                self.parentframe.columnconfigure(1,weight=1)

                self.rowindexheader.grid(row=0,column=0,rowspan=1,sticky='news')
                self.colheader.grid(row=0,column=1,rowspan=1,sticky='news')
                self.rowheader.grid(row=1,column=0,rowspan=1,sticky='news')
                self.grid(row=1,column=1,rowspan=1,sticky='news',pady=0,ipady=0)

                self.adjustColumnWidths()
                #bind redraw to resize, may trigger redraws when widgets added
                self.parentframe.bind("<Configure>", self.resized) #self.redrawVisible)
                self.colheader.xview("moveto", 0)
                self.xview("moveto", 0)
                if self.showtoolbar == True:
                    self.toolbar = myToolBar(self.parentframe, self)
                    self.toolbar.grid(row=0,column=3,rowspan=2,sticky='news')
                if self.showstatusbar == True:
                    self.statusbar = statusBar(self.parentframe, self)
                    self.statusbar.grid(row=3,column=0,columnspan=2,sticky='ew')
                #self.redraw(callback=callback)
                self.currwidth = self.parentframe.winfo_width()
                self.currheight = self.parentframe.winfo_height()
                if hasattr(self, 'pf'):
                    self.pf.updateData()

            def pasteTable(self, event=None):
                """Paste a new table from the clipboard"""

                self.storeCurrent()
                #Adding code for reading email data as well
                df = pd.read_clipboard()
                if len(df.columns)==1:
                    pandasPaste()
                try:
                    df = pd.read_clipboard(sep=',',on_bad_lines='skip')
                except Exception as e:
                    root.attributes('-topmost', True)
                    messagebox.showwarning("Could not read data", e, parent=root)
                    root.attributes('-topmost', False)
                    return
                if len(df) == 0:
                    return
                
                df = pd.read_clipboard(sep=',', on_bad_lines='skip')
                if len(df.columns)!=8:
                    root.attributes('-topmost', True)
                    messagebox.showwarning("WARNING","All 8 columns not copied please check",parent=root)
                    root.attributes('-topmost', False)
                    return
                try:
                    df['RM']
                except:
                    root.attributes('-topmost', True)
                    messagebox.showwarning("WARNING","RM Column not present in pasted table",parent=root)
                    root.attributes('-topmost', False)
                    return
                try:
                    df['Qty']
                except:
                    root.attributes('-topmost', True)
                    messagebox.showwarning("WARNING","Qty Column not present in pasted table",parent=root)
                    root.attributes('-topmost', False)
                    return
                try:
                    df['Saw Cut']
                except:
                    root.attributes('-topmost', True)
                    messagebox.showwarning("WARNING","Saw Cut Column not present in pasted table",parent=root)
                    root.attributes('-topmost', False)
                    return
                
                # self.autoResizeColumns()

                #EntryBox row adder
                
                global bakerDf
                previousDf = bakerDf.copy()
                
                # bakerDf['Quote_Yes_No'], bakerDf['Location'], bakerDf['Type'],bakerDf['Spec'], bakerDf['Grade'], bakerDf['Yield'], bakerDf['OD'], bakerDf['ID'] = [None, None, None, None, None, None, None, None]
                # bakerDf['Length'], bakerDf['Qty'], bakerDf['E_SELLING_COST/LBS'], bakerDf['UOM'], bakerDf["E_SELLING_COST/UOM"], bakerDf['ADDITIONAL_COST'] = [None, None, None, None, None, None]
                # bakerDf['LEAD_TIME'], bakerDf['FINAL_PRICE'] = [None, None]

                
                rowNum=0
                while len(previousDf)!=rowNum:
                    deleteRow()
                    rowNum+=1


                
                bakerDf = df.copy()
                bakerDf["C_Quote Yes/No"], bakerDf["E_Location"], bakerDf["E_Type"],bakerDf["E_Spec"], bakerDf["E_Grade"], bakerDf["E_Yield"], bakerDf["E_OD1"], bakerDf["E_ID1"] = [None, None, None, None, None, None, None, None]
                bakerDf["E_OD2"], bakerDf["E_ID2"] = [None, None]
                bakerDf["E_Length"], bakerDf["E_Qty"], bakerDf["E_COST"] , bakerDf["E_MarginLBS"],bakerDf["E_Selling Cost/LBS"], bakerDf["E_UOM"] = [None, None, None, None, None, None]
                bakerDf["E_Selling Cost/UOM"], bakerDf["E_Additional_Cost"], bakerDf["E_LeadTime"], bakerDf["E_Final Price"] = [None, None, None, None]
                bakerDf["E_freightIncured"], bakerDf["E_freightCharged"], bakerDf["E_Margin_Freight"], bakerDf["Lot_Serial_Number"] = [None, None, None, None]
                #declaring temporary baker df
                global temp_bakerDf
                temp_bakerDf = bakerDf.copy()
                ##Adding entry box columns in df
                i = 1
                while i<len(df):
                    addRow()
                    i+=1

                #####################for single row insertion##########################
                if len(df)==1:
                    addRow()
                    deleteRow()
                    

                # df["Quote_Yes_No"] = None
                # df['Location'] = None
                model = TableModel(df)
                self.updateModel(model)
                self.redraw()
                ptBaker.autoResizeColumns()
                return
        
            def queryBar(self, evt=None):
                """Query/filtering dialog"""

                if hasattr(self, 'qframe') and self.qframe != None:
                    return
                self.qframe = myQueryDialog(self)
                self.qframe.grid(row=self.queryrow,column=0,columnspan=3,sticky='news')
                return

            def showAll(self):
                global bakerDf
                global temp_bakerDf
                """Re-show unfiltered"""

                if hasattr(self, 'dataframe'):
                    self.model.df = self.dataframe
                self.filtered = False
                self.redraw()
                #Redrawing entry boxes
                print(len(quoteYesNo))
                #Just like removing grid as per temp dataframe now adding grid back as per temp dataframe
                for row_number in range(len(specialList['C_Quote Yes/No'])):
                    if row_number not in list(temp_bakerDf.index):
                        keyIndex= list(specialList.keys()).index('C_Quote Yes/No')
                        for i in range(keyIndex,len(list(specialList.keys()))):
                            newKey = list(specialList.keys())[i]
                            print(newKey)
                            if newKey!='E_OD2' and newKey != 'E_ID2':
                                specialList[newKey][0][row_number][0].grid()
                #Updating main baker df frame columns as per temp dataframe as well
                bakerDf.update(temp_bakerDf)
                temp_bakerDf = bakerDf.copy()
                #changeing tem_bakerDf to baker Df
                # while len(quoteYesNo)<len(self.dataframe):
                #     addRow()
                # return
        
        
        
        def tempbakerDfUpdate():
            
            xlList = ["C_Specification","C_Type","C_Grade","C_Yield", "C_OD", "C_ID", "C_QRD", "C_Length", "C_Qty", 'E_freightIncured', 'E_freightCharged','E_Margin_Freight', 'Lot_Serial_Number', "searchLocation"]
            for index in range(len(temp_bakerDf)):
                for key in specialList.keys():
                    if key!='E_OD2' and key != 'E_ID2' and key not in xlList:
                        temp_bakerDf[key][index] = specialList[key][0][index][1].get()
                        # print(temp_bakerDf[key][index])
            #update baker main df
            bakerDf.update(temp_bakerDf)
        # def jump_to_widget_top(self, widget):
        #     if self.scrollable.winfo_height() > self.canvas.winfo_height():
        #         pos = widget.winfo_rooty() - self.scrollable.winfo_rooty()
        #         height = self.scrollable.winfo_height()
        #         self.canvas.yview_moveto(pos / height)
        # def jump_to_widget_top(self, widget):
        #     if m_entryFrame.winfo_width() > entryCanvas.winfo_width():
        #         pos = widget.winfo_rooty() - self.scrollable.winfo_rooty()
        #         width = m_entryFrame.winfo_width()
        #         entryCanvas.yview_moveto(pos / width)
        
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
                #scroll Canvas using mouse wheel
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
                    root.attributes('-topmost', True)
                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in Integer or Decimal format only",parent=root)
                    root.attributes('-topmost', False)
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
                    root.attributes('-topmost', True)
                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in Integer format only",parent=root)
                    root.attributes('-topmost', False)
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

        # def tabFunc(e):
        #     try:
        #         cx_spec[0][0].focus_set()
        #         return "break"
            
        #     except Exception as e:
        #         raise e
        def entryBoxUpdater(row_number, index_num = -1):
            try:
                if index_num == -1:
                    index_num = row_number
                global temp_bakerDf
                if len(temp_bakerDf):
                    #updating entry box values from baker dataframe for handling condition after clearing filter
                    quoteYesNo[row_number][0].configure(state='normal')
                    quoteYesNo[row_number][1].set(temp_bakerDf["C_Quote Yes/No"][index_num] if temp_bakerDf["C_Quote Yes/No"][index_num]!= None else "")
                    root.update()
                    e_location[row_number][0].configure(state='normal')
                    e_location[row_number][1].set(temp_bakerDf["E_Location"][index_num] if temp_bakerDf["E_Location"][index_num]!=None else "")
                    root.update()
                    e_type[row_number][0].configure(state='normal')
                    e_type[row_number][1].set(temp_bakerDf["E_Type"][index_num] if temp_bakerDf["E_Type"][index_num]!=None else "")
                    root.update()
                    e_spec[row_number][0].configure(state='normal')
                    e_spec[row_number][1].set(temp_bakerDf["E_Spec"][index_num] if temp_bakerDf["E_Spec"][index_num] !=None else "")
                    root.update()
                    e_grade[row_number][0].configure(state='normal')
                    e_grade[row_number][1].set(temp_bakerDf["E_Grade"][index_num] if temp_bakerDf["E_Grade"][index_num] !=None else "")
                    root.update()
                    e_yield[row_number][0].configure(state='normal')
                    e_yield[row_number][1].set(temp_bakerDf["E_Yield"][index_num] if temp_bakerDf["E_Yield"][index_num] !=None else "")
                    root.update()
                    e_od1[row_number][0].configure(state='normal')
                    e_od1[row_number][1].set(temp_bakerDf["E_OD1"][index_num] if temp_bakerDf["E_OD1"][index_num]!=None else "")
                    root.update()
                    e_id1[row_number][0].configure(state='normal')
                    e_id1[row_number][1].set(temp_bakerDf["E_ID1"][index_num] if temp_bakerDf["E_ID1"][index_num]!=None else "")
                    root.update()
                    # e_od2[row_num][1].set(temp_bakerDf["E_OD2"][row_num] if temp_bakerDf["E_OD2"][row_num]!=None else "")
                    # e_id2[row_num][1].set(temp_bakerDf["E_ID2"][row_num] if temp_bakerDf["E_ID2"][row_num]!=None else "")
                    e_len[row_number][0].configure(state='normal')
                    e_len[row_number][1].set(temp_bakerDf["E_Length"][index_num] if temp_bakerDf["E_Length"][index_num]!=None else "")
                    root.update()
                    e_qty[row_number][0].configure(state='normal')
                    e_qty[row_number][1].set(temp_bakerDf["E_Qty"][index_num] if temp_bakerDf["E_Qty"][index_num]!=None else "")
                    root.update()
                    e_cost[row_number][0].configure(state='normal')
                    e_cost[row_number][1].set(temp_bakerDf["E_COST"][index_num] if temp_bakerDf["E_COST"][index_num]!=None else "")
                    root.update()
                    sellCostLBS[row_number][0].configure(state='normal')
                    sellCostLBS[row_number][1].set(temp_bakerDf["E_Selling Cost/LBS"][index_num] if temp_bakerDf["E_Selling Cost/LBS"][index_num]!=None else "")
                    root.update()
                    marginlbs[row_number][0].configure(state='normal')
                    marginlbs[row_number][1].set(temp_bakerDf["E_MarginLBS"][index_num] if temp_bakerDf["E_MarginLBS"][index_num]!=None else "")
                    root.update()
                    e_uom[row_number][0].configure(state='normal')
                    e_uom[row_number][1].set(temp_bakerDf["E_UOM"][index_num] if temp_bakerDf["E_UOM"][index_num]!=None else "")
                    root.update()
                    addCost[row_number][0].configure(state='normal')
                    addCost[row_number][1].set(temp_bakerDf["E_Additional_Cost"][index_num] if temp_bakerDf["E_Additional_Cost"][index_num]!=None else "")
                    root.update()
                    leadTime[row_number][0].configure(state='normal')
                    leadTime[row_number][1].set(temp_bakerDf["E_LeadTime"][index_num] if temp_bakerDf["E_LeadTime"][index_num]!=None else "")
                    root.update()
                    finalCost[row_number][0].configure(state='normal')
                    finalCost[row_number][1].set(temp_bakerDf["E_Final Price"][index_num] if temp_bakerDf["E_Final Price"][index_num]!=None else "")
                    root.update()
                    # freightIncured[row_number][0].configure(state='normal')
                    # freightIncured[row_number][1].set(temp_bakerDf["E_freightIncured"][index_num] if temp_bakerDf["E_freightIncured"][index_num]!=None else "")
                    # root.update()
                    # freightCharged[row_number][0].configure(state='normal')
                    # freightCharged[row_number][1].set(temp_bakerDf["E_freightCharged"][index_num] if temp_bakerDf["E_freightCharged"][index_num]!=None else "")
                    # root.update()
                    # marginFreight[row_number][0].configure(state='normal')
                    # marginFreight[row_number][1].set(temp_bakerDf["E_Margin_Freight"][index_num] if temp_bakerDf["E_Margin_Freight"][index_num]!=None else "")
                    # root.update()
                else:
                    pass
            except Exception as e:
                raise e

        def addRow(check=False):
            global bakerDf
            global temp_bakerDf
            try:
                row_num = len(quoteYesNo)
                # print(df)
                
                print(row_num)
                
                        
                        # time.sleep(1)
                        # addRow()
                    

                ####getting data from input data frame
                # if row_num!=0:
                if row_num == 1:
                    if not check:
                        # xlList = ["C_Specification","C_Type","C_Grade","C_Yield", "C_OD", "C_ID", "C_QRD", "C_Length", "C_Qty", "Lot_Serial_Number", "searchLocation"]
                        xlList = ["C_Specification","C_Type","C_Grade","C_Yield", "C_OD", "C_ID", "C_QRD", "C_Length", "C_Qty", 'E_freightIncured', 'E_freightCharged','E_Margin_Freight', 'Lot_Serial_Number', "searchLocation"]
                        for key in specialList.keys():
                            
                            # specialList[key][0][-1][1].destroy()
                        
                            if key!='E_OD2' and key != 'E_ID2' and key not in xlList:
                                specialList[key][0][-1][0].destroy()
                            specialList[key][0].pop()
                        addRow(check=True)
                    if row_num>=len(bakerDf):
                        bakerDf = pd.concat([bakerDf, pd.DataFrame(columns=bakerDf.columns, data=[[None] * len(bakerDf.columns)] * 1)])
                        bakerDf.reset_index(drop=True,inplace=True)
                        temp_bakerDf = pd.concat([temp_bakerDf, pd.DataFrame(columns=temp_bakerDf.columns, data=[[None] * len(temp_bakerDf.columns)] * 1)])
                        temp_bakerDf.reset_index(drop=True,inplace=True)
                        ptBaker.model.df = pd.concat([ptBaker.model.df, pd.DataFrame(columns=ptBaker.model.df.columns, data=[[None] * len(ptBaker.model.df.columns)] * 1)])
                        ptBaker.model.df.reset_index(drop=True,inplace=True)
                        ptBaker.redraw()
                        

                    cx_spec.append((bakerDf['RM'][row_num][2:6] if bakerDf['RM'][row_num] is not None else "", None)) #((None, None))
                    cx_type.append((bakerDf['RM'][row_num][:2] if bakerDf['RM'][row_num] is not None else "", None))
                    cx_od.append((int(bakerDf['RM'][row_num][6:10])/100 if bakerDf['RM'][row_num] is not None else "", None))
                    cx_id.append((int(bakerDf['RM'][row_num][10:14])/100 if bakerDf['RM'][row_num] is not None else "", None))
                    cx_qrd.append((bakerDf['RM'][row_num][14:] if bakerDf['RM'][row_num] is not None else "", None))
                    cx_len.append((bakerDf['Saw Cut'][row_num] if bakerDf['Saw Cut'][row_num] is not None else "", None))
                    cx_qty.append((bakerDf['Qty'][row_num] if bakerDf['Qty'][row_num] is not None else "", None))
                    cx_grade.append((None, None))
                    cx_yield.append((None, None))

                    # cx_spec.append((bakerDf['RM'][row_num][2:6], None)) #((None, None))
                    # cx_type.append((bakerDf['RM'][row_num][:2], None))
                    # cx_od.append((bakerDf['RM'][row_num][6:10], None))
                    # cx_id.append((bakerDf['RM'][row_num][10:14], None))
                    # cx_qrd.append((bakerDf['RM'][row_num][14:], None))
                    # cx_len.append((bakerDf['Saw Cut'][row_num], None))
                    # cx_qty.append((bakerDf['Qty'][row_num], None))
                    # cx_grade.append((None, None))
                    # cx_yield.append((None, None))
                    
                else:
                    if row_num>=len(bakerDf):#For Adding Blank row in both tables
                        bakerDf = pd.concat([bakerDf, pd.DataFrame(columns=bakerDf.columns, data=[[None] * len(bakerDf.columns)] * 1)])
                        bakerDf.reset_index(drop=True,inplace=True)
                        temp_bakerDf = pd.concat([temp_bakerDf, pd.DataFrame(columns=temp_bakerDf.columns, data=[[None] * len(temp_bakerDf.columns)] * 1)])
                        temp_bakerDf.reset_index(drop=True,inplace=True)
                        ptBaker.model.df = pd.concat([ptBaker.model.df, pd.DataFrame(columns=ptBaker.model.df.columns, data=[[None] * len(ptBaker.model.df.columns)] * 1)])
                        ptBaker.model.df.reset_index(drop=True,inplace=True)
                        ptBaker.redraw()
                    cx_spec.append((bakerDf['RM'][row_num][2:6] if bakerDf['RM'][row_num] is not None else "", None)) #((None, None))
                    cx_type.append((bakerDf['RM'][row_num][:2] if bakerDf['RM'][row_num] is not None else "", None))
                    cx_od.append((int(bakerDf['RM'][row_num][6:10])/100 if bakerDf['RM'][row_num] is not None else "", None))
                    cx_id.append((int(bakerDf['RM'][row_num][10:14])/100 if bakerDf['RM'][row_num] is not None else "", None))
                    cx_qrd.append((bakerDf['RM'][row_num][14:] if bakerDf['RM'][row_num] is not None else "", None))
                    cx_len.append((bakerDf['Saw Cut'][row_num] if bakerDf['Saw Cut'][row_num] is not None else "", None))
                    cx_qty.append((bakerDf['Qty'][row_num] if bakerDf['Qty'][row_num] is not None else "", None))
                    cx_grade.append((None, None))
                    cx_yield.append((None, None))

                
                vcmd = tab1.register(intFloat)
                
                # if row_num==0:
                #     entpady = (5,0)
                # else:
                entpady = (0,0)
                
                quoteYesNo.append(myCombobox(df,tab1,item_list=["Yes","No","Other"],frame=entryFrame,row=2+row_num,column=0,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = quoteYesNo[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=["Yes","No","Other"],frame=entryFrame,row=2+row_num,column=0,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                
                # print(quoteYesNo[0][0].grid_info())
                # quoteYesNo[-1]['validate']='focusout'
                # quoteYesNo[-1]['validatecommand'] = (quoteYesNo[-1].register(yesNo),'%P','%W')
                e_location.append(myCombobox(df,tab1,item_list=["Dubai","Singapore","USA","UK"],frame=entryFrame,row=2+row_num,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_location[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=["Dubai","Singapore","USA","UK"],frame=entryFrame,row=2+row_num,column=1,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                

                # e_location[-1].config(textvariable="NA", state='disabled')
                e_type.append(myCombobox(df,tab1,item_list=["HT","HB", "HM"],frame=entryFrame,row=2+row_num,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_type[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=["THF","BR"],frame=entryFrame,row=2+row_num,column=2,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf)[1]
                
                # e_spec_var = StringVar()
                # e_spec.append((ttk.Entry(entryFrame, width=5,font=('Segoe UI', 10),foreground='blue', background='white', textvariable=e_spec_var),e_spec_var))
                # e_spec[-1][0].grid(row=2+row_num,column=3,padx=5,pady=entpady)
                e_spec.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=3,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_spec[0][1]
                # 
                # e_type[-1].config(textvariable="NA", state='disabled')
                e_grade.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=4,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_grade[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=4,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf)[1]
                # e_grade[-1].config(textvariable="NA", state='disabled')
                e_yield.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=5,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_yield[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=5,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf)[1]
                # e_yield[-1].config(textvariable="NA", state='disabled')
                e_od1.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=6,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_od1[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=6,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf)[1]
                # e_od[-1].config(textvariable="NA", state='disabled')
                e_od1[-1][0]['validate']='key'
                e_od1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')
                
                e_id1.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=7,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_id1[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=7,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,pt=pt,entpady=entpady, bakerDf=temp_bakerDf)[1]
                e_id1[-1][0]['validate']='key'
                e_id1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')

                e_od2.append((None, None))
                e_id2.append((None, None))

                searchLocation.append((None, None))

                # e_id[-1].config(textvariable="NA", state='disabled')
                e_len.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=8,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_len[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=8,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                e_len[-1][0]['validate']='key'
                e_len[-1][0]['validatecommand'] = (e_len[-1][0].register(intFloat),'%P','%d')
                # e_len[-1].config(textvariable="NA", state='disabled')
                e_qty.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=9,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_qty[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=9,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                e_qty[-1][0]['validate']='key'
                e_qty[-1][0]['validatecommand'] = (e_qty[-1][0].register(intChecker),'%P','%d')
                # e_qty[-1].config(textvariable="NA", state='disabled')

                e_cost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=10,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                e_cost[-1][0]['validate']='key'
                e_cost[-1][0]['validatecommand'] = (e_cost[-1][0].register(intFloat),'%P','%d')


                sellCostLBS.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=11,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = sellCostLBS[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=10,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                sellCostLBS[-1][0]['validate']='key'
                sellCostLBS[-1][0]['validatecommand'] = (sellCostLBS[-1][0].register(intFloat),'%P','%d')
                # sellCostLBS[-1].config(textvariable="NA", state='disabled')
                marginlbs.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=12,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))

                e_uom.append(myCombobox(df,tab1,item_list=["Inch","Each","Foot"],frame=entryFrame,row=2+row_num,column=13,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = e_uom[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=["Inch","Each"],frame=entryFrame,row=2+row_num,column=11,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                # e_uom[-1].config(textvariable="NA", state='disabled')
                
                sellCostUOM.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=14,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = sellCostUOM[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=12,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                sellCostUOM[-1][0]['validate']='key'
                sellCostUOM[-1][0]['validatecommand'] = (sellCostUOM[-1][0].register(intFloat),'%P','%d')
                # sellCostUOM[-1].config(textvariable="NA", state='disabled')
                addCost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=15,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # temp_bakerDf = addCost[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=13,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                addCost[-1][0]['validate']='key'
                addCost[-1][0]['validatecommand'] = (addCost[-1][0].register(intFloat),'%P','%d')
                # addCost[-1].config(textvariable="NA", state='disabled')
                leadTime.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=16,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                leadTime[-1][0]['validate']='key'
                leadTime[-1][0]['validatecommand'] = (leadTime[-1][0].register(intFloat),'%P','%d')
                # temp_bakerDf = leadTime[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=14,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                # leadTime[-1].config(textvariable="NA", state='disabled')
                finalCost.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=17,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf))
                # finalCost[0].bind("<Enter>", jump_to_widget_top)
                # temp_bakerDf = finalCost[0][1]
                # temp_bakerDf = myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=15,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList,entpady=entpady, bakerDf=temp_bakerDf)[1]
                finalCost[-1][0]['validate']='key'
                finalCost[-1][0]['validatecommand'] = (finalCost[-1][0].register(intFloat),'%P','%d')

                # freightIncured.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=18,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                freightIncured.append((None, None))
                # freightCharged.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=19,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                freightCharged.append((None, None))


                

                # marginFreight.append(myCombobox(df,tab1,item_list=item_list,frame=entryFrame,row=2+row_num,column=20,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = specialList))
                marginFreight.append((None, None))
                lot_serial_number.append((None, None))

                
                if row_num==1:
                    entryBoxUpdater(row_number=0)
                    entryBoxUpdater(row_number=1)
                else:
                    entryBoxUpdater(row_number=row_num)

                #Moving horizontal scroll bar to initial position
                # entryCanvas.xview_moveto('1.0')
                entryCanvas.xview("moveto", 0)


            except Exception as e:
                raise e

        def cxListCalc():
            try:
                cxList = [cxDatadict["Prepared_By"],salespersonDict["sales_person"][0][0][0].get(),cxDatadict["Date"],cxDatadict["cus_long_name"][0][0][0].get(), cxDatadict["payment_term"][0][0].get(), currency.get(),  cxDatadict["cus_address"][0][0].get(),
                    cxDatadict["cus_phone"][0][0].get(),cxDatadict["cus_email"][0][0].get(),cxDatadict["cus_city_zip"]]
                return cxList
            except Exception as e:
                raise e
        def otherListCalc():
            try:
                # otherList = [validityVar.get(), addCommVar.get()]
                otherList = [None, remarksVar.get()]
                return otherList
                # row_num+=1
            except Exception as e:
                raise e
        def deleteRowConfirm():
            try:
                root.attributes('-topmost', True)
                if messagebox.askyesno("Warning", "Are sure that you want to delete last row?",parent=root):
                    root.attributes('-topmost', False)
                    deleteRow()
                else:
                    pass
                root.attributes('-topmost', False)
            except Exception as e:
                raise e
        def deleteRow():
            try:
                
                global quoteDf
                quoteDf = []
                #deleting row from datafrmes as well
                bakerDf.drop(bakerDf.tail(1).index,inplace=True)
                temp_bakerDf.drop(temp_bakerDf.tail(1).index,inplace=True)
                ptBaker.model.df.drop(ptBaker.model.df.tail(1).index,inplace=True)
                ptBaker.redraw()
                submitButton.configure(state='disable')
                xlList = ["C_Specification","C_Type","C_Grade","C_Yield", "C_OD", "C_ID", "C_QRD", "C_Length", "C_Qty", 'E_freightIncured', 'E_freightCharged','E_Margin_Freight', 'Lot_Serial_Number', "searchLocation"]
                for key in specialList.keys():
                    
                    # specialList[key][0][-1][1].destroy()
                    if (len(specialList[key][0])==1):
                        if key!='E_OD2' and key != 'E_ID2' and key not in xlList:
                            specialList[key][0][0][0].configure(state='normal')
                            specialList[key][0][0][0].delete(0, tk.END)
                            entryCanvas.yview_moveto('1.0')
                        
                        # time.sleep(1)
                        # addRow()
                    else:
                        if key!='E_OD2' and key != 'E_ID2' and key not in xlList:
                            specialList[key][0][-1][0].destroy()
                        specialList[key][0].pop()
                # show bottom of canvas
                entryCanvas.yview("moveto", 0)
                root.update()
                entryCanvas.yview("moveto", 0)
                
                #Updating dataframe values in entry boxes after filter or deletion
                
                # entryCanvas.yview_moveto('1.0')
            except Exception as e:
                raise e
        
        def create_xl():
            try:
                MyTable.showAll(ptBaker)
                global quoteDf
                global bakerxlDf
                if len(bakerDf):
                    pass
                #Creating dataframes for uploading into database as well as saving quote xl in current directory
                dfList = bakerMaker(specialList,cxListCalc(),otherListCalc(),ptBaker,conn,root)
                try:
                    quoteDf = dfList[0]
                    bakerxlDf = dfList[1]
                except:
                    quoteDf = dfList
                    bakerxlDf = dfList
                    
                if len(quoteDf) and len(bakerxlDf):
                    pt.model.df = quoteDf
                    pt.redraw()
                    new_quoteNo = quoteDf['QUOTENO'][0]
                    #Displaying xl preview in pandastable
                    xlRoot = tk.Toplevel()
                    xlRoot.title(new_quoteNo)
                    xlRoot.title(quoteDf["QUOTENO"][0])
                    ptBakerxl = MyTable(xlRoot, editable=True,dataframe=bakerxlDf,showtoolbar=False, showstatusbar=False, maxcellwidth=1500, width=600)
                    ptBakerxl.font = 'Segoe UI'
                    ptBakerxl.fontsize = 10
                    ptBakerxl.thefont = ('Segoe UI', 10)
                    ptBakerxl.show()
                    
                    # bakerxlDf.to_excel(new_quoteNo+'.xlsx')
                #     global pdf_path
                #     pdf_path = pdf_generator(quoteDf)
                    submitButton.configure(state='normal')
                else:
                    root.attributes('-topmost', True)
                    messagebox.showerror("Error", "Empty dataframe was given in input",parent=root)
                    root.attributes('-topmost', False)
            except Exception as e:
                raise e

        def uploadDf(conn, quoteDf):
            try:
                # pt.model.df = quoteDf
                # pt.redraw()
                root.attributes('-topmost', True)
                if messagebox.askyesno("Upload to Database", "Are sure that you want to generate quote and upload Data?",parent=root):
                    root.attributes('-topmost', False)
                    submitButton.configure(state='disable')
                    eagsQuotationuploader(conn, quoteDf, latest_revised_quote=None, baker=True)
                    
                    root.attributes('-topmost', True)
                    messagebox.showinfo("Info", "Data uploaded Successfully!",parent=root)
                    root.attributes('-topmost', False)

                    # current_work_dir = os.getcwd()#To be Shared Drive
                    # current_work_dir = r'I:\EAGS\Quotes'
                    cx_init_name = str(quoteDf['QUOTENO'][0]).split("_")[0]
                    filename = str(quoteDf['QUOTENO'][0])+".xlsx"
                    # save_dir = current_work_dir+"\\"+cx_init_name
                    # if not os.path.exists(save_dir):
                    #     os.mkdir(save_dir)

                    desktopDir = os.path.join(os.environ["HOMEPATH"], "Desktop\\EAGS_Quotes")
                    desktopDir = os.path.join('C:', desktopDir)
                    if not os.path.exists(desktopDir):
                        os.mkdir(desktopDir)
                    # shutil.copy(save_dir+"\\"+filename, desktopDir)
                    bakerxlDf.to_excel(desktopDir+"\\"+filename, index=False)
                    localpath = desktopDir+"\\"+filename
                    shpUploader(localpath, filename)
                    
                    # os.rename(pdf_path,save_dir+"\\"+filename)
                    
                else:
                    # os.remove(pdf_path)
                    pass
                    submitButton.configure(state='disable')
                root.attributes('-topmost', False)                
            except Exception as e:
                raise e


        
        mainRoot.withdraw()
        global row_num
        row_num=0

        #Getting invoentory dataframe
        # df = get_inv_df(conn,table = INV_TABLE)
        
        
        # df = pd.read_csv("sampleInventory.csv")
        # Getting Cx Dataframe
        cx_df = get_cx_df(conn,table = CX_TABLE, customer='baker')
        salesperson_df = get_salesperson_df(conn,table = S_PERSON_TABLE)
        # cx_df = pd.read_excel("cxDatabase.xlsx")

        count = 0
        root = tk.Toplevel(mainRoot, bg = "#9BC2E6")
        root.state('zoomed')
        root.title("EAGS Baker Quote Generator")

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        pixelInfo = root.winfo_pixels('1i')
        print(screen_width)
        print(screen_height)
        print(pixelInfo)
        print(root.winfo_screendepth())
        # if pixelInfo == 96.0:
        #     root.tk.call('tk', 'scaling', 1.7)
        tab1=root
        # tabControl = ttk.Notebook(root)
        # s = ttk.Style(tabControl)
        # s.configure("TFrame", background=root["bg"])
        # tab1 = ttk.Frame(tabControl)
        # # tab2 = ttk.Frame(tabControl)
        # # tab3 = ttk.Frame(tabControl)

        # tabControl.add(tab1, text='Baker Quote Generator')
        # # tabControl.add(tab2, text='Machining')
        # # tabControl.add(tab3, text='Quote Generator + Machining')

        # tabControl.pack(expand=1, fill='both')
        cxFrame = tk.Frame(tab1, bg = "#9BC2E6")#,highlightbackground="blue", highlightthickness=2)
        cxFrame2 = tk.Frame(tab1, bg = "#9BC2E6")#,highlightbackground="blue", highlightthickness=2)

        m_entryFrame = tk.Frame(tab1, bg= "#DDEBF7",highlightbackground="black", highlightthickness=2)#width=1700,height=300
        bakerTableFrame = tk.Frame(tab1, bg= "#DDEBF7",highlightbackground="black", highlightthickness=2)#width=1700,height=300
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
        bakerTableFrame.grid(row=1,column=0, sticky=tk.NSEW)
        m_entryFrame.grid(row=1, column=1,sticky="nsew")#, columnspan=2)
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
        salespersonDict = {}
        salespersonDict["sales_person"] = []
        salespersonVar = []
        salespersonDict["sales_person"].append(salespersonVar)

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


        item_list = () #'A4140', 'A4140M', 'A4330V', 'A4715', 'BS708', 'A4145M', '4542','4462'

        cxLabel = tk.Label(cxFrame, text="Customer Details", bg = "#9BC2E6", font=("Segoe UI", 10))
        prepByLb = tk.Label(cxFrame,text="Prepared By", bg = "#9BC2E6", font=("Segoe UI", 10))
        prep_by = ttk.Entry(cxFrame)
        prep_by.insert(tk.END, user)
        salespersonLb = tk.Label(cxFrame,text="Sales Person", bg = "#9BC2E6", font=("Segoe UI", 10))
        inpDateLb = tk.Label(cxFrame,text="Date", bg = "#9BC2E6", font=("Segoe UI", 10))
        inpDate = MyDateEntry(master=cxFrame, width=17, selectmode='day', font=("Segoe UI", 10))
        cxNameLb = tk.Label(cxFrame,text="Customer Name", bg = "#9BC2E6", font=("Segoe UI", 10))
        locAddLb = tk.Label(cxFrame,text="Location/Address", bg = "#9BC2E6", font=("Segoe UI", 10))
        emailLb = tk.Label(cxFrame,text="Email", bg = "#9BC2E6", font=("Segoe UI", 10))
        payTermLb = tk.Label(cxFrame,text="Payment Terms", bg = "#9BC2E6", font=("Segoe UI", 10))

        #Adding Search button in cxFrame 2
        # starButton = tk.Button(cxFrame2, text="Star Search", font = ("Segoe UI", 10, 'bold'), bg="#20bebe", fg="white", height=1, width=14, command=lambda: starSearch(root, df), activebackground="#20bebb", highlightbackground="#20bebd")
        rangeButton = tk.Button(cxFrame2, text="Range Search", font = ("Segoe UI", 10, 'bold'), bg="#20bebe", fg="white", height=1, width=14, command=lambda: rangeSearch(root, df, specialList, 0, pt), activebackground="#20bebb", highlightbackground="#20bebd")
        

        mobileLb = tk.Label(cxFrame2, text="Mobile", bg = "#9BC2E6", font=("Segoe UI", 10))
        currencyLabel = tk.Label(cxFrame2,text="Currency", bg = "#9BC2E6", font=("Segoe UI", 10))
        # valLb = tk.Label(cxFrame2,text="Validity", bg = "#9BC2E6", font=("Segoe UI", 10))
        # addCommLb = tk.Label(cxFrame2,text="Additional Comments", bg = "#9BC2E6", font=("Segoe UI", 10))
        remarksLabel = tk.Label(cxFrame2,text="Remarks", bg = "#9BC2E6", font=("Segoe UI", 10))
        cxLabel.grid(row=0,column=0)
        prepByLb.grid(row=1,column=0)
        prep_by.grid(row=2,column=0)
        
        salespersonLb.grid(row=1,column=1)

        inpDateLb.grid(row=1,column=1)
        inpDate.grid(row=2, column=1)
        cxNameLb.grid(row=3,column=0)
        locAddLb.grid(row=3,column=1)
        emailLb.grid(row=3,column=2)
        payTermLb.grid(row=3,column=3)
        
        #label grid using cxFrame2
        # starButton.grid(row=0, column=0, pady=(20,0))
        mobileLb.grid(row=1, column=0, pady=(20,0))
        rangeButton.grid(row=0,column=0, pady=(20,0))
        currencyLabel.grid(row=1,column=1, pady=(20,0))
        currencyLabel.grid(row=1,column=1, pady=(20,0))
        remarksLabel.grid(row=1,column=2, pady=(20,0))#padx=(50,5)
        # valLb.grid(row=1,column=2, pady=(20,0))#5x=(50,5)
        # addCommLb.grid(row=1,column=3, pady=(20,0))#padx=(50,5)

        
        
        
        prep_by.config(state= "disabled")
        cxDatadict["Prepared_By"] = prep_by.get()

        
        
        
        cxDatadict["Date"] = inpDate.get()
        
        #Currency
        currencyVar = tk.StringVar()
        currency = ttk.Combobox(cxFrame2, background='white', font=('Segoe UI', 10), justify='center',textvariable=currencyVar,values=["$","£"], width=5, text='$')
        currency.delete(0, tk.END)
        currency.insert(tk.END,'$')
        # currency = ttk.Entry(cxFrame2, textvariable=currencyVar, foreground='blue', background = 'white',width = 10, font=('Segoe UI', 10))
        currency.grid(row=2,column=1,pady=5)
        #Remarks
        remarksVar = tk.StringVar()
        remarks = ttk.Entry(cxFrame2, textvariable=remarksVar, foreground='blue', background = 'white',width = 40, font=('Segoe UI', 10))
        remarks.grid(row=2,column=2,pady=5)

        # #Validity
        # validityVar = tk.StringVar()
        # validity = ttk.Entry(cxFrame2, textvariable=validityVar, foreground='blue', background = 'white',width = 15, font=('Segoe UI', 10))
        # validity.grid(row=2,column=2,pady=5)
        
        # #Additional Comments
        # addCommVar = tk.StringVar()
        # addComm = ttk.Entry(cxFrame2, textvariable=addCommVar, foreground='blue', background = 'white',width = 15, font=('Segoe UI', 10))
        # addComm.grid(row=2,column=3,sticky=tk.EW,pady=5)

        # addComm.bind("<Tab>",tabFunc)
        
        


        #Customer Name Entry Box
        cxNameVar.append(myCombobox(cx_df,tab1,item_list=list(cx_df['cus_long_name']),frame=cxFrame,row=4,column=0,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",cxDict= cxDatadict,val=currency))
        #location Address entry box
        salespersonVar.append(myCombobox(salesperson_df,root,item_list=list(salesperson_df['sales_person']),frame=cxFrame,row=2,column=1,width=25,list_bd = 0,foreground='blue',
         background='white',sticky = "nsew",salesDict= salespersonDict,val=cxNameVar[0][0]))
        
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
        home_button.grid(row=0,column=2,sticky="ne")
        button_dict[home_button] = [home_img, home_img1]
        home_button.bind("<Enter>", on_enter)
        home_button.bind("<Leave>", on_leave)
        # home_button.place(x=1600,y=-10,relx=0.1,rely=0.1,anchor="sw")
        #######################################
        
        ########################################
        
        ################defining baker table####################
        nonList = [[None,None,None,None,None,None,None,None]]
        pandasDf = pd.DataFrame(nonList,columns=['Work Order', 'Dlv Date',	'MaterialNumber', 'Materialdescription', 'Qty', 'RM', 'RMQty', 'Saw Cut'])
        global bakerDf
        bakerDf = pd.DataFrame(nonList,columns=['Work Order', 'Dlv Date',	'MaterialNumber', 'Materialdescription', 'Qty', 'RM', 'RMQty', 'Saw Cut'])
        bakerDf["C_Quote Yes/No"], bakerDf["E_Location"], bakerDf["E_Type"],bakerDf["E_Spec"], bakerDf["E_Grade"], bakerDf["E_Yield"], bakerDf["E_OD1"], bakerDf["E_ID1"] = [None, None, None, None, None, None, None, None]
        bakerDf["E_OD2"], bakerDf["E_ID2"] = [None, None]
        bakerDf["E_Length"], bakerDf["E_Qty"], bakerDf["E_COST"] , bakerDf["E_MarginLBS"],bakerDf["E_Selling Cost/LBS"], bakerDf["E_UOM"] = [None, None, None, None, None, None]
        bakerDf["E_Selling Cost/UOM"], bakerDf["E_Additional_Cost"], bakerDf["E_LeadTime"], bakerDf["E_Final Price"] = [None, None, None, None]
        bakerDf["E_freightIncured"], bakerDf["E_freightCharged"], bakerDf["E_Margin_Freight"], bakerDf["Lot_Serial_Number"] = [None, None, None, None]
        # bakerDf['Quote_Yes_No'], bakerDf['Location'], bakerDf['Type'],bakerDf['Spec'], bakerDf['Grade'], bakerDf['Yield'], bakerDf['OD'], bakerDf['ID'] = [None, None, None, None, None, None, None, None]
        # bakerDf['Length'], bakerDf['Qty'], bakerDf['E_SELLING_COST/LBS'], bakerDf['UOM'], bakerDf["E_SELLING_COST/UOM"], bakerDf['ADDITIONAL_COST'] = [None, None, None, None, None, None]
        # bakerDf['LEAD_TIME'], bakerDf['FINAL_PRICE'] = [None, None]
        global temp_bakerDf
        temp_bakerDf = bakerDf.copy()
        # pandasDf = pd.DataFrame(cx_df)
        global ptBaker
        ptBaker = MyTable(bakerTableFrame, editable=True,dataframe=pandasDf,showtoolbar=True, showstatusbar=False, maxcellwidth=1500)
        
        ptBaker.cellwidth=130
        ptBaker.autoResizeColumns()
        subtractor=0
        if root.winfo_pixels('1i')==96:
            subtractor=6.25
        ptBaker.rowheight=29-subtractor
        ptBaker.font = 'Segoe UI'
        ptBaker.fontsize = 10
        ptBaker.thefont = ('Segoe UI', 10, 'bold')
        ptBaker.show()


        

        #################Entry Form Section##############################################
        ######################defining labels############################################
        # specLabel = tk.Label(entryFrame, text="Specification", bg= "#DDEBF7")
        # gradeLabel = tk.Label(entryFrame, text="Grade", bg= "#DDEBF7")
        # yieldLabel = tk.Label(entryFrame, text="Yield", bg= "#DDEBF7")
        # odLabel = tk.Label(entryFrame, text="OD", bg= "#DDEBF7")
        # idLabel = 	tk.Label(entryFrame, text="ID", bg= "#DDEBF7")
        # lengthLabel = tk.Label(entryFrame, text="Length", bg= "#DDEBF7")
        # qtyLabel = tk.Label(entryFrame, text="Qty", bg= "#DDEBF7")
        quoteLabel1 = tk.Label(entryFrame, text="Quote", bg= "#DDEBF7")
        quoteLabel2 = tk.Label(entryFrame, text="Yes/No", bg= "#DDEBF7")
        locationLabel = tk.Label(entryFrame, text="Location", bg= "#DDEBF7")
        typeLabel = tk.Label(entryFrame, text="Type", bg= "#DDEBF7")
        e_specLabel = tk.Label(entryFrame, text="Spec", bg= "#DDEBF7")
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
        leadTimeLabel1 = tk.Label(entryFrame, text="Lead Time", bg= "#DDEBF7")
        leadTimeLabel2 = tk.Label(entryFrame, text="in Days", bg= "#DDEBF7")
        
        finalPriceLabel = tk.Label(entryFrame, text="Final Price/UOM", bg= "#DDEBF7")

        # freightCostLabel1 = tk.Label(entryFrame, text="Freight", bg= "#DDEBF7")
        # freightCostLabel2 = tk.Label(entryFrame, text="Incured", bg= "#DDEBF7")
        # freightSaleLabel1 = tk.Label(entryFrame, text="Freight to", bg= "#DDEBF7")
        # freightSaleLabel2 = tk.Label(entryFrame, text="be Charged", bg= "#DDEBF7")
        
        

        # marginFreightLabel1 = tk.Label(entryFrame, text="Freight Margin", bg= "#DDEBF7")
        # marginFreightLabel2 = tk.Label(entryFrame, text="%", bg= "#DDEBF7")

        # specLabel.grid(row=0,column=0,padx=(15,0), sticky="ew")
        # gradeLabel.grid(row=0,column=1, sticky="ew")
        # yieldLabel.grid(row=0,column=2, sticky="ew")
        # odLabel.grid(row=0,column=3, sticky="ew")
        # idLabel.grid(row=0,column=4, sticky="ew")
        # lengthLabel.grid(row=0,column=5, sticky="ew")
        # qtyLabel.grid(row=0,column=6, sticky="ew")
        quoteLabel1.grid(row=0,column=0, sticky="ew")
        quoteLabel2.grid(row=1,column=0, sticky="ew")
        locationLabel.grid(row=0,column=1, sticky="ew")
        typeLabel.grid(row=0,column=2, sticky="ew")
        e_specLabel.grid(row=0,column=3, sticky="ew")
        e_gradeLabel.grid(row=0,column=4, sticky="ew")
        e_yieldLabel.grid(row=0,column=5, sticky="ew")
        e_odLabel.grid(row=0,column=6, sticky="ew")
        e_idLabel.grid(row=0,column=7, sticky="ew")
        e_Length.grid(row=0,column=8, sticky="ew")
        e_Qty.grid(row=0,column=9, sticky="ew")

        e_costLabel.grid(row=0,column=10, sticky="ew")

        sellcostLbsLabel1.grid(row=0,column=11, sticky="ew")
        sellcostLbsLabel2.grid(row=1,column=11, sticky="ew")

        marginLBSLabel1.grid(row=0,column=12, sticky="ew")
        marginLBSLabel2.grid(row=1,column=12, sticky="ew")

        uom.grid(row=0,column=13, sticky="ew")
        sellcostUOMLabel1.grid(row=0,column=14, sticky="ew")
        sellcostUOMLabel2.grid(row=1,column=14, sticky="ew")
        addCostLabel1.grid(row=0,column=15, sticky="ew")
        addCostLabel2.grid(row=1,column=15, sticky="ew")
        leadTimeLabel1.grid(row=0,column=16, sticky="ew")
        leadTimeLabel2.grid(row=1,column=16, sticky="ew")

        finalPriceLabel.grid(row=0,column=17,padx=(0,10), sticky="ew")

        # freightCostLabel1.grid(row=0,column=18, sticky="ew")
        # freightCostLabel2.grid(row=1,column=18, sticky="ew")
        # freightSaleLabel1.grid(row=0,column=19, sticky="ew")
        # freightSaleLabel2.grid(row=1,column=19, sticky="ew")

        

        # marginFreightLabel1.grid(row=0,column=20,padx=(0,10), sticky="ew")
        # marginFreightLabel2.grid(row=1,column=20,padx=(0,10), sticky="ew")
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

        cx_qrd = []
        specialList["C_QRD"] = []
        specialList["C_QRD"].append(cx_qrd)

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

        #Adding for margin
        e_cost = []
        specialList["E_COST"] = []
        specialList["E_COST"].append(e_cost)

        sellCostLBS = []
        specialList["E_Selling Cost/LBS"] = []
        specialList["E_Selling Cost/LBS"].append(sellCostLBS)

        #Adding for margin
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

        #Adding for margin
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
        
        while len(quoteYesNo)<1:
            addRow()
        
        # button_dict = {}

        addRowbut = tk.Button(controlFrame, image=add_img, command=addRow,borderwidth=0, background=controlFrame["bg"])
        addRowbut.image = add_img
        addRowbut.grid(row=0,column=1)
        button_dict[addRowbut] = [add_img, add_img2]
        addRowbut.bind("<Enter>", on_enter)
        addRowbut.bind("<Leave>", on_leave)

        deleteRowbut = tk.Button(controlFrame, image=delete_img, text="Delete Row",command=deleteRowConfirm,borderwidth=0, background=controlFrame["bg"])
        deleteRowbut.image = delete_img
        deleteRowbut.grid(row=1,column=1)
        button_dict[deleteRowbut] = [delete_img, delete_img2]
        deleteRowbut.bind("<Enter>", on_enter)
        deleteRowbut.bind("<Leave>", on_leave)
        
            
        Previewbut = tk.Button(controlFrame, image=preview_img,text="Preview",command=create_xl,borderwidth=0, background=controlFrame["bg"])
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

        #Focus on Sales Person
        salespersonVar[0][0].focus_set()

        #Moving horizontal scroll bar to initial position
        entryCanvas.xview("moveto", 0)

        def on_closing(window):
            window.attributes('-topmost', True)
            if messagebox.askokcancel("Quit", "Do you want to quit?",parent=window):
                window.attributes('-topmost', False)
                # mainRoot.destroy()
                conn.close()
                root.destroy()
                sys.exit()
            window.attributes('-topmost', False)
        mainRoot.protocol("WM_DELETE_WINDOW", lambda: on_closing(mainRoot))
        root.protocol("WM_DELETE_WINDOW", lambda: on_closing(root)) 
    except Exception as e:
        raise e
    
    # root.mainloop()

# conn = get_connection()
# # conn=None
# mainRoot = tk.Tk()
# user = "Imam"
# # df = pd.read_excel("sampleInventory.xlsx")
# df = get_inv_df(conn,table = INV_TABLE)
# bakerQuoteGenerator(mainRoot, user, conn, df)
# mainRoot.mainloop()
