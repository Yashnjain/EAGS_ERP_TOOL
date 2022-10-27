from shutil import ExecError
import pandas as pd
from datetime import datetime, date
from sfTool import eagsQuotationuploader,getLatestQuote
import os, sys
import tkinter as tk
from tkinter import ttk
import customComboboxV2 #import myCombobox
from tkinter import messagebox
from pandasPaste import pandasPaste
from pandastable.headers import ColumnHeader, RowHeader, IndexHeader
from pandastable.dialogs import *
from pandastable.core import *


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
        # img = images.copy()
        # addButton(self, 'Copy', self.parentapp.copyTable, img, 'copy table to clipboard')
        # img = images.paste()
        # addButton(self, 'Paste', self.parentapp.pasteTable, img, 'paste table')
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
                values = df[self.filtercol.get()].unique() if self.filtercol.get()!='' else [''],
                #initialitem = initial,
                width = 14)
        valsbox.grid(row=0,column=3,sticky='news',padx=2,pady=2)

        #handling value update in 1st entry box
        def updateValue(*args):
            valsbox['values'] = sorted(list(df[self.filtercol.get()].unique()) if self.filtercol.get()!='' else [''], reverse=True)
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


    def applyFilter(self, df, mask=None):
        """Apply the widget based filters, returns a boolean mask"""

        if mask is None:
            mask = df.index==df.index

        for f in self.filters:
            col, val, op, b = f.getFilter()
            try:
                val = float(val)
            except:
                pass
            #print (col, val, op, b)
            if op == 'contains':
                m = df[col].str.contains(str(val))
            elif op == 'equals':
                m = df[col]==val
            elif op == 'not equals':
                m = df[col]!=val
            elif op == '>':
                m = df[col]>val
            elif op == '<':
                m = df[col]<val
            elif op == 'is empty':
                m = df[col].isnull()
            elif op == 'not empty':
                m = ~df[col].isnull()
            elif op == 'excludes':
                m = -df[col].str.contains(val)
            elif op == 'starts with':
                m = df[col].str.startswith(val)
            elif op == 'has length':
                m = df[col].str.len()>val
            elif op == 'is number':
                m = df[col].astype('object').str.isnumeric()
            elif op == 'is lowercase':
                m = df[col].astype('object').str.islower()
            elif op == 'is uppercase':
                m = df[col].astype('object').str.isupper()
            else:
                continue
            if b == 'AND':
                mask = mask & m
            elif b == 'OR':
                mask = mask | m
            elif b == 'NOT':
                mask = mask ^ m
        return mask

    def query(self, evt=None):
        """Do query"""
        # global temp_bakerDf
        table = self.table
        #saving current values in temporary df and update main baker dataframe
        # print(temp_bakerDf)
        # tempbakerDfUpdate()
        s = ''#self.queryvar.get()
        if table.filtered == True:
            table.model.df = table.dataframe
        df = table.model.df
        #instead of taking df from pandas table we will pick it from current temp baker df
        # df = temp_bakerDf.copy()
        #Declaring prev df length variable
        # prev_len = len(df)
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
        
        # temp_bakerDf = filtdf.copy()
        # temp_bakerDf["C_Quote Yes/No"], temp_bakerDf["E_Location"], temp_bakerDf["E_Type"],temp_bakerDf["E_Spec"], temp_bakerDf["E_Grade"], temp_bakerDf["E_Yield"], temp_bakerDf["E_OD1"], temp_bakerDf["E_ID1"] = [None, None, None, None, None, None, None, None]
        # temp_bakerDf["E_OD2"], temp_bakerDf["E_ID2"] = [None, None]
        # temp_bakerDf["E_Length"], temp_bakerDf["E_Qty"], temp_bakerDf["E_Selling Cost/LBS"], temp_bakerDf["E_UOM"], temp_bakerDf["E_Selling Cost/UOM"], temp_bakerDf["E_Additional_Cost"] = [None, None, None, None, None, None]
        # temp_bakerDf["E_LeadTime"], temp_bakerDf["E_Final Price"] = [None, None]
        #deleting exxtra entry boxs
        # curr_len = len(filtdf)
        # while curr_len < prev_len:
        #     deleteRow()
        #     prev_len -= 1
        ###############check this
        #Removing grid of those rows which are not present in temporary dataframe
        # for row_number in range(len(specialList['C_Quote Yes/No'])):
        #     if row_number not in list(temp_bakerDf.index):
        #         keyIndex= list(specialList.keys()).index('C_Quote Yes/No')
        #         for i in range(keyIndex,len(list(specialList.keys()))):
        #             newKey = list(specialList.keys())[i]
        #             print(newKey)
        #             if newKey!='E_OD2' and newKey != 'E_ID2':
        #                 specialList[newKey][0][row_number][0].grid_remove()
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
            self.root.attributes('-topmost', True)
            messagebox.showwarning("Could not read data", e,
                                    parent=self.root)
            self.root.attributes('-topmost', False)
            return
        if len(df) == 0:
            return

        df = pd.read_clipboard(sep=',', on_bad_lines='skip')
        model = TableModel(df)
        self.updateModel(model)
        self.redraw()
        # ptBaker.autoResizeColumns()
        # self.autoResizeColumns()

        #EntryBox row adder
        
        global bakerDf
        bakerDf = df.copy()
        bakerDf["C_Quote Yes/No"], bakerDf["E_Location"], bakerDf["E_Type"],bakerDf["E_Spec"], bakerDf["E_Grade"], bakerDf["E_Yield"], bakerDf["E_OD1"], bakerDf["E_ID1"] = [None, None, None, None, None, None, None, None]
        bakerDf["E_OD2"], bakerDf["E_ID2"] = [None, None]
        bakerDf["E_Length"], bakerDf["E_Qty"], bakerDf["E_Selling Cost/LBS"], bakerDf["E_UOM"], bakerDf["E_Selling Cost/UOM"], bakerDf["E_Additional_Cost"] = [None, None, None, None, None, None]
        bakerDf["E_LeadTime"], bakerDf["E_Final Price"] = [None, None]
        # bakerDf['Quote_Yes_No'], bakerDf['Location'], bakerDf['Type'],bakerDf['Spec'], bakerDf['Grade'], bakerDf['Yield'], bakerDf['OD'], bakerDf['ID'] = [None, None, None, None, None, None, None, None]
        # bakerDf['Length'], bakerDf['Qty'], bakerDf['E_SELLING_COST/LBS'], bakerDf['UOM'], bakerDf["E_SELLING_COST/UOM"], bakerDf['ADDITIONAL_COST'] = [None, None, None, None, None, None]
        # bakerDf['LEAD_TIME'], bakerDf['FINAL_PRICE'] = [None, None]

        #declaring temporary baker df
        global temp_bakerDf
        temp_bakerDf = bakerDf.copy()
        

        # i = 1
        # while i<len(df):
        #     addRow()
        #     i+=1

        ##Adding entry box columns in df
        
        # df["Quote_Yes_No"] = None
        # df['Location'] = None
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
        # print(len(quoteYesNo))
        # #Just like removing grid as per temp dataframe now adding grid back as per temp dataframe
        # for row_number in range(len(specialList['C_Quote Yes/No'])):
        #     if row_number not in list(temp_bakerDf.index):
        #         keyIndex= list(specialList.keys()).index('C_Quote Yes/No')
        #         for i in range(keyIndex,len(list(specialList.keys()))):
        #             newKey = list(specialList.keys())[i]
        #             print(newKey)
        #             if newKey!='E_OD2' and newKey != 'E_ID2':
        #                 specialList[newKey][0][row_number][0].grid()
        # #Updating main baker df frame columns as per temp dataframe as well
        # bakerDf.update(temp_bakerDf)
        # temp_bakerDf = bakerDf.copy()



def resource_path(relative_path):
    try:
        """ Get absolute path to resource, works for dev and for PyInstaller """
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    except Exception as e:
                raise e







def dfMaker(specialList,cxList,otherList,pt,conn,root):
    try:
        # colList = list(inpDict.keys())
        # for col in colList:
        #     for i in range(len(inpDict[col][0])):
        #         inpDict[col][0][i] = inpDict[col][0][i][0].get()
        # EAGS/Location/Year/Number
        # Location: USA/UK/DUB/SGP
        # Year: 2022
        # Number:00xxxx
        for i in range(len(cxList)):
            if i == "" or i == None:
                root.attributes('-topmost', True)
                messagebox.showerror("Error", f"Empty Customer entry found, please fill and then click preview",parent=root)
                root.attributes('-topmost', False)
                return []
        for i in range(len(otherList)):
            if i == "" or i == None:
                root.attributes('-topmost', True)
                messagebox.showerror("Error", f"Currency or Validity or Additional Comment is not filled, please fill and then click preview",parent=root)
                root.attributes('-topmost', False)
                return []
        if specialList['E_Location'][0][-1][0].get() != '' and cxList[2] != '':
            locDict = {"DUBAI":"DUB", "SINGAPORE":"SGP", "USA":"USA","UK":"UK"}
            locVar = specialList['E_Location'][0][0][0].get()
            if locVar != "NA":
                location = locDict[locVar.upper()]

            cxList[1]=datetime.strptime(cxList[1],"%m.%d.%Y").date()
            input_year=datetime.strftime(cxList[1],"%Y")
            
            #try to get latest quote number of same combination if not present then put 1 otherwise increament current contract
            # curr_quoteNo = f"EAGS/{location}/{input_year}/000001"
            cx_init_name = cxList[2].split(" ")[0]
            curr_quoteNo = f"{cx_init_name}_000001"

            new_quoteNo, raw_data = getLatestQuote(conn,curr_quoteNo, previous_quote_number=None, newQuote=True)

            #Appending Previous Quote as None
            otherList.append(None)
            #Appending Rev_Checker
            otherList.append(1)
            otherList.append(date.today())



            # columnList = ['QUOTENO', 'PREPAREDBY', 'DATE', 'CUS_NAME', 'PAYMENT_TERM', 'CURRENCY', 'CUS_ADDRESS', 'CUS_PHONE', 'CUS_EMAIL', 'CUS_CITY_ZIP', 'C_SPECIFICATION', 'C_TYPE',
            # 'C_GRADE', 'C_YIELD', 'C_OD', 'C_ID', 'C_LENGTH', 'C_QTY', 'C_QUOTE_YES/NO', 'E_LOCATION', 'E_TYPE', 'E_SPEC','E_GRADE', 'E_YIELD', 'E_OD1', 'E_ID1', 'E_OD2', 'E_ID2', 'E_LENGTH',
            # 'E_QTY', 'E_SELLING_COST/LBS', 'E_UOM', 'E_SELLING_COST/UOM', 'E_ADDITIONAL_COST', 'LEAD_TIME','E_FINAL_PRICE', 'VALIDITY', 'ADD_COMMENTS','PREVIOUS_QUOTE','REV_CHECKER','INSERT_DATE']

            columnList = ['QUOTENO', 'PREPAREDBY', 'DATE', 'CUS_NAME', 'PAYMENT_TERM', 'CURRENCY', 'CUS_ADDRESS', 'CUS_PHONE', 'CUS_EMAIL', 'CUS_CITY_ZIP', 'C_SPECIFICATION', 'C_TYPE',
            'C_GRADE', 'C_YIELD', 'C_OD', 'C_ID', 'C_LENGTH', 'C_QTY', 'C_QUOTE_YES/NO', 'E_LOCATION', 'E_TYPE', 'E_SPEC','E_GRADE', 'E_YIELD', 'E_OD1', 'E_ID1', 'E_OD2', 'E_ID2', 'E_LENGTH',
            'E_QTY', 'E_COST', 'E_SELLING_COST/LBS', 'E_MARGIN_LBS','E_UOM', 'E_SELLING_COST/UOM', 'E_ADDITIONAL_COST', 'LEAD_TIME','E_FINAL_PRICE', 'E_FREIGHT_INCURED', 'E_FREIGHT_CHARGED', 
            'E_MARGIN_FREIGHT','LOT_SERIAL_NUMBER','VALIDITY', 'ADD_COMMENTS','PREVIOUS_QUOTE','REV_CHECKER','INSERT_DATE']
            row = []
            
            colList = list(specialList.keys())
            for i in range(len(specialList[colList[0]][0])):
                rowList = []
                rowList.append(new_quoteNo)
                rowList.extend(cxList)
                for col in colList:
                    if col == 'searchYield' or col == 'searchGrade' or col == 'searchLocation':
                        pass
                    elif (col == 'E_OD2' or col == 'E_ID2' or col == 'C_Type' or col == 'E_Spec'):
                        rowList.append(specialList[col][0][i][0])
                        if specialList[col][0][i][0] == "":
                            root.attributes('-topmost', True)
                            messagebox.showerror("Error", f"Empty Entry box {col} found in {i} row, please fill and then click preview",parent=root)
                            root.attributes('-topmost', False)
                            return []
                    else:

                        if col == 'Lot_Serial_Number':
                            rowList.append(specialList[col][0][i][0])
                        else:
                            rowList.append(specialList[col][0][i][0].get()) #Insert jth column with ith index in rowList
                            if specialList[col][0][i][0].get() == "" and (col.upper() != "C_YIELD" and col.upper() != "E_YIELD" and col.upper() != "C_SPECIFICATION" and col.upper() != "C_GRADE"):
                                if col in ['E_freightIncured', 'E_freightCharged', 'E_Margin_Freight']:
                                    if not(specialList['E_freightIncured'][0][i][0].get() == specialList['E_freightCharged'][0][i][0].get() == specialList['E_Margin_Freight'][0][i][0].get()):
                                        root.attributes('-topmost', True)
                                        messagebox.showerror("Error", f"Empty Entry box {col} found in {i} row, please fill and then click preview",parent=root)
                                        root.attributes('-topmost', False)
                                        return []
                                else:
                                    root.attributes('-topmost', True)
                                    messagebox.showerror("Error", f"Empty Entry box {col} found in {i} row, please fill and then click preview",parent=root)
                                    root.attributes('-topmost', False)
                                    return []
                                
                rowList.extend(otherList)#insert validity and additional comments
                row.append(rowList)#Append current ith row to row List
                #Empty rowList for fetching next row
            # print(row)
            # print(len(columnList))
            # print(len(row[0]))

            if specialList["E_Final Price"][0][0][0].get() != "":
                sfDf = pd.DataFrame(row, columns=columnList)
                # print(sfDf)
                # pt.model.df = sfDf
                # pt.redraw()
                # eagsQuotationuploader(conn, sfDf)

                
            # print()
            return sfDf
        else:
            return []
    except Exception as e:
        raise e


def bakerMaker(specialList,cxList,otherList,ptBaker,conn,root):
    try:
        # colList = list(inpDict.keys())
        # for col in colList:
        #     for i in range(len(inpDict[col][0])):
        #         inpDict[col][0][i] = inpDict[col][0][i][0].get()
        # EAGS/Location/Year/Number
        # Location: USA/UK/DUB/SGP
        # Year: 2022
        # Number:00xxxx
        for i in range(len(cxList)):
            if i == "" or i == None:
                root.attributes('-topmost', True)
                messagebox.showerror("Error", f"Empty Customer entry found, please fill and then click preview",parent=root)
                root.attributes('-topmost', False)
                return []
        for i in range(len(otherList)):
            if i == "" or i == None:
                root.attributes('-topmost', True)
                messagebox.showerror("Error", f"Currency or Validity or Additional Comment is not filled, please fill and then click preview",parent=root)
                root.attributes('-topmost', False)
                return []
        if specialList['E_Location'][0][0][0].get() != '' and cxList[2] != '':
            locDict = {"DUBAI":"DUB", "SINGAPORE":"SGP", "USA":"USA","UK":"UK"}
            locVar = specialList['E_Location'][0][0][0].get()
            if locVar != "NA":
                location = locDict[locVar.upper()]

            cxList[1]=datetime.strptime(cxList[1],"%m.%d.%Y").date()
            input_year=datetime.strftime(cxList[1],"%Y")
            
            #try to get latest quote number of same combination if not present then put 1 otherwise increament current contract
            # curr_quoteNo = f"EAGS/{location}/{input_year}/000001"
            cx_init_name = cxList[2].split(" ")[0]
            curr_quoteNo = f"{cx_init_name}_000001"

            new_quoteNo, raw_data = getLatestQuote(conn,curr_quoteNo, previous_quote_number=None, baker=True, newQuote=True)

            #Appending Previous Quote as None
            otherList.append(None)
            #Appending Rev_Checker
            otherList.append(1)
            otherList.append(date.today())



            columnList = ['QUOTENO', 'PREPAREDBY', 'DATE', 'CUS_NAME', 'PAYMENT_TERM', 'CURRENCY', 'CUS_ADDRESS', 'CUS_PHONE', 'CUS_EMAIL', 'CUS_CITY_ZIP', 'C_SPECIFICATION', 'C_TYPE',
            'C_GRADE', 'C_YIELD', 'C_OD', 'C_ID', 'C_LENGTH', 'C_QTY', 'C_QUOTE_YES/NO', 'E_LOCATION', 'E_TYPE', 'E_SPEC','E_GRADE', 'E_YIELD', 'E_OD1', 'E_ID1', 'E_OD2', 'E_ID2', 'E_LENGTH',
            'E_QTY', 'E_COST', 'E_SELLING_COST/LBS', 'E_MARGIN_LBS', 'E_UOM', 'E_SELLING_COST/UOM', 'E_ADDITIONAL_COST', 'LEAD_TIME','E_FINAL_PRICE', 'E_FREIGHT_INCURED', 'E_FREIGHT_CHARGED',
            'E_MARGIN_FREIGHT','LOT_SERIAL_NUMBER', 'VALIDITY', 'ADD_COMMENTS','PREVIOUS_QUOTE','REV_CHECKER', 'INSERT_DATE']

            # columnList = ['QUOTENO', 'PREPAREDBY', 'DATE', 'CUS_NAME', 'PAYMENT_TERM', 'CUS_ADDRESS', 'CUS_PHONE', 'CUS_EMAIL',
            #  'CUS_CITY_ZIP','C_QUOTE_YES/NO', 'E_LOCATION', 'E_TYPE', 'E_SPEC','E_GRADE', 'E_YIELD', 'E_OD1', 'E_ID1', 'E_OD2', 'E_ID2',
            #   'E_LENGTH','E_QTY', 'E_SELLING_COST/LBS', 'E_UOM', 'E_SELLING_COST/UOM', 'E_ADDITIONAL_COST', 'LEAD_TIME',
            #   'E_FINAL_PRICE', 'VALIDITY', 'ADD_COMMENTS','INSERT_DATE']
            row = []
            bakerxlDf = ptBaker.model.df.copy()
            bakerxlDf['RM Offer'], bakerxlDf['Price'], bakerxlDf['Location'], bakerxlDf['Lead Time'], bakerxlDf['Remarks'] = [None, None, None, None, None]
            xlList = ["C_Specification","C_Type","C_Grade","C_Yield", "C_OD", "C_ID", "C_Length", "C_Qty", 'E_freightIncured', 'E_freightCharged','E_Margin_Freight', 'Lot_Serial_Number']
            colList = list(specialList.keys())
            
            
            #Inserting quote number in bakerxlDf
            bakerxlDf.insert(0, 'QUOTENO', new_quoteNo)

            for i in range(len(specialList[colList[0]][0])):
                #Making here 2nd dataframe as well
                e_type = str(specialList['E_Type'][0][i][0].get()).upper()
                e_spec = str(specialList['E_Spec'][0][i][0].get()).upper()
                e_od = str(specialList['E_OD1'][0][i][0].get()).replace('.','').zfill(4)
                e_id = str(specialList['E_ID1'][0][i][0].get()).replace('.','').zfill(4)
                e_qrd = str(specialList['C_QRD'][0][i][0]).upper()

                bakerxlDf['RM Offer'][i] = e_type+e_spec+e_od+e_id+e_qrd if e_type!='NA' else "NA"
                bakerxlDf['Price'][i] = specialList['E_Final Price'][0][i][0].get()
                bakerxlDf['Location'][i] = specialList['E_Location'][0][i][0].get()
                bakerxlDf['Lead Time'][i] = specialList['E_LeadTime'][0][i][0].get()
                bakerxlDf['Remarks'][i] = otherList[1]
                bakerxlDf['Dlv Date'] = bakerxlDf['Dlv Date'].astype('datetime64[ns]').astype(str)
                rowList = []
                rowList.append(new_quoteNo)
                rowList.extend(cxList)
                for col in colList:
                    #Adding condition for not including extracted cx details
                    # if col not in xlList:
                    if col == 'C_QRD' or col == 'searchYield' or col == 'searchGrade' or col == 'searchLocation':
                        pass
                    elif col == 'E_OD2' or col == 'E_ID2' or col in xlList:
                        print(specialList[col][0][i][0])
                        rowList.append(specialList[col][0][i][0])
                        if specialList[col][0][i][0] == "":
                            root.attributes('-topmost', True)
                            messagebox.showerror("Error", f"Empty Entry box {col} found in {i} row, please fill and then click preview",parent=root)
                            root.attributes('-topmost', False)
                            return []
                    else:
                        
                        if specialList[col][0][i][0].get() == "" and col!="E_Yield":
                            root.attributes('-topmost', True)
                            messagebox.showerror("Error", f"Empty Entry box {col} found in {i} row, please fill and then click preview",parent=root)
                            root.attributes('-topmost', False)
                            return []
                        
                        if col == 'E_Spec':
                            rowList.append(specialList[col][0][i][0].get().upper())
                        else:
                            rowList.append(specialList[col][0][i][0].get()) #Insert jth column with ith index in rowList
                rowList.extend(otherList)#insert validity and additional comments
                row.append(rowList)#Append current ith row to row List
                #Empty rowList for fetching next row
            # print(row)
            # print(len(columnList))
            # print(len(row[0]))

            if specialList["E_Final Price"][0][-1][0].get() != "":
                initDf = pd.DataFrame(row, columns=columnList)
                #Saving excel in current Directory
                
                # df1 = bakerxlDf.iloc[:,:9]
                # df2 = sfDf.iloc[:,17:]
                # result = pd.concat([df1, df2], axis=1)

                ###separating starting columns from sfDf
                df1 = initDf.iloc[:,:9]
                df2 = bakerxlDf.iloc[:,1:9]#Getting cx input data
                df3 = initDf.iloc[:,9:]#getting reaing main dataframe columns
                sfDf = pd.concat([df1, df2, df3], axis=1)
                print("Df created")
                # bakerxlDf = 
                
                # print(sfDf)
                # pt.model.df = sfDf
                # pt.redraw()
                # eagsQuotationuploader(conn, sfDf)

                
            # print()
            return [sfDf, bakerxlDf]
        else:
            if  cxList[2] == '':
                root.attributes('-topmost', True)
                messagebox.showerror("Error", "Empty Customer dataframe was given in input",parent=root)
                root.attributes('-topmost', False)
            return []
    except Exception as e:
        raise e


    # row = [[trade_date, flow_m1, ch_opis, ny_opis, flow[flow_m1][2], flow[flow_m1][3], flow[flow_m1][1],flow[flow_m1][0], in_date, up_date]]
    #     for i in f_keys[1:len(f_keys)]:
    #         row.append([trade_date, i, np.nan, np.nan, flow[i][2], flow[i][3], flow[i][1], flow[i][0], in_date, up_date])

    #     df = pd.DataFrame(row, columns=['TRADEDATE', 'FLOWMONTH', 'OPIS_CH', 'OPIS_NY', 'NYMEX_CH', 'NYMEX_NY', 'NYMEX_RBOB', 'CORN',
    #                                     'INSERT_DATE', 'UPDATE_DATE'], index=None)




# 'QuoteNo, PreparedBy, Date, Cus_Name, Payment_Term, Cus_Address, Cus_Phone, Cus_Email, Cus_city_zip, C_Specification, C_Grade, C_Yield, C_OD, C_ID, C_Length, C_Qty, C_Quote Yes/No, E_Location, E_Type, E_Grade, E_Yield, E_OD, E_ID, E_Length, E_Qty, E_Selling Cost/LBS, E_UOM, E_Selling Cost / UOM, E_Additional Cost, E_Final Price'
# 'quoteno, preparedby, date, cus_name, payment_term, cus_address, cus_phone, cus_email, cus_city_zip, c_specification, c_grade, c_yield, c_od, c_id, c_length, c_qty, c_quote yes/no, e_location, e_type, e_grade, e_yield, e_od, e_id, e_length, e_qty, e_selling cost/lbs, e_uom, e_selling cost / uom, e_additional cost, e_final price'
# 'QUOTENO, PREPAREDBY, DATE, CUS_NAME, PAYMENT_TERM, CUS_ADDRESS, CUS_PHONE, CUS_EMAIL, CUS_CITY_ZIP, C_SPECIFICATION, C_GRADE, C_YIELD, C_OD, C_ID, C_LENGTH, C_QTY, C_QUOTE YES/NO, E_LOCATION, E_TYPE, E_GRADE, E_YIELD, E_OD, E_ID, E_LENGTH, E_QTY, E_SELLING COST/LBS, E_UOM, E_SELLING COST / UOM, E_ADDITIONAL COST, E_FINAL PRICE, VALIDITY, ADD_COMMENTS'



def specialCase(root, boxList,pt,df,index, item_list, bakerDf=[],cxDict=[]):
    try:
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

        def intChecker(inStr,acttyp):
            try:
                # if acttyp == '1': #insert
                if inStr == '' or inStr == "NA":
                    return True
                if not inStr.isdigit():
                    toproot.attributes('-topmost', True)
                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in Integer format only",parent=toproot)
                    toproot.attributes('-topmost', False)
                    return False
                return True
            except Exception as e:
                raise e
        check=False
        toproot = tk.Toplevel(root, bg = "#9BC2E6")
        toproot.title('EAGS Quote Generator')
        screen_width = toproot.winfo_screenwidth()
        screen_height = toproot.winfo_screenheight()

        width = 315
        height = 280
        # calculate position x and y coordinates
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        toproot.geometry('%dx%d+%d+%d' % (width, height, x, y))
        toproot.attributes('-topmost', True)
        toproot.grab_set()

        labelFrame = tk.Frame(toproot, bg= "#9BC2E6")
        labelFrame.grid(row=0, column=1)
        entryFrame1 = tk.Frame(labelFrame, bg= "#9BC2E6")
        entryFrame1.grid(row=1, column=0)
        entryFrame2 = tk.Frame(labelFrame, bg= "#9BC2E6")
        entryFrame2.grid(row=1, column=1)
        submitFrame = tk.Frame(toproot, bg= "#9BC2E6")
        submitFrame.grid(row=1, column=1)
        


        #Declaring Labels
        tobePrinted = tk.Label(labelFrame, text="To be Printed", bg = "#9BC2E6")
        tobePrinted.grid(row=0, column=0)

        tobeCalculated = tk.Label(labelFrame, text="To be Calculated", bg = "#9BC2E6")
        tobeCalculated.grid(row=0,column=1)

        od1Label = tk.Label(entryFrame1, text="OD", bg = "#9BC2E6")
        od1Label.grid(row=0, column=0)

        id1Label = tk.Label(entryFrame1, text="ID", bg = "#9BC2E6")
        id1Label.grid(row=0, column=1)

        od2Label = tk.Label(entryFrame2, text="OD", bg = "#9BC2E6")
        od2Label.grid(row=0, column=0)

        id2Label = tk.Label(entryFrame2, text="ID", bg = "#9BC2E6")
        id2Label.grid(row=0, column=1)

        #Configuring frame sizes based on screen size
        toproot.grid_rowconfigure(0, weight=1)
        toproot.grid_columnconfigure(0, weight=1)
        toproot.grid_rowconfigure(1, weight=1)
        toproot.grid_columnconfigure(1, weight=1)

        labelFrame.grid_rowconfigure(0, weight=1)
        labelFrame.grid_columnconfigure(0, weight=1)
        labelFrame.grid_rowconfigure(1, weight=1)
        labelFrame.grid_columnconfigure(1, weight=1)

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
        
        
        
        od1Var = tk.StringVar()
        vcmd = toproot.register(intFloat)
        od1 = ttk.Entry(entryFrame1, textvariable=od1Var, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        od1['validatecommand'] = (od1.register(intFloat),'%P','%d')
        od1.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)

        id1Var = tk.StringVar()
        id1 = ttk.Entry(entryFrame1, textvariable=id1Var, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        id1['validatecommand'] = (id1.register(intFloat),'%P','%d')
        id1.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        


        boxList["E_OD2"][0][index] = customComboboxV2.myCombobox(df,toproot,frame=entryFrame2,row=1,column=0,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = boxList,pt=pt,item_list=item_list, bakerDf=bakerDf,cxDict=cxDict)
        boxList["E_OD2"][0][index][0]['validate']='key'
        boxList["E_OD2"][0][index][0]['validatecommand'] = (boxList["E_OD2"][0][index][0].register(intFloat),'%P','%d')
        # e_od[-1].config(textvariable="NA", state='disabled')
        # e_od2[-1][0]['validate']='key'
        # e_od2[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')

        boxList["E_ID2"][0][index] = customComboboxV2.myCombobox(df,toproot,frame=entryFrame2,row=1,column=1,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = boxList,pt=pt,item_list=item_list, bakerDf=bakerDf,cxDict=cxDict)
        boxList["E_ID2"][0][index][0]['validate']='key'
        boxList["E_ID2"][0][index][0]['validatecommand'] = (boxList["E_ID2"][0][index][0].register(intFloat),'%P','%d')
        # e_id1[-1][0]['validate']='key'
        # e_id1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')

        # od2Var = tk.StringVar()
        # od2 = ttk.Entry(entryFrame2, textvariable=od2Var, background = 'white',width = 15)
        # od2.grid(row=1,column=0)

        # id2Var = tk.StringVar()
        # id2 = ttk.Entry(entryFrame2, textvariable=id2Var, background = 'white',width = 15)
        # id2.grid(row=1,column=1)
        def exitTrue(close_check=False):
            try:
                # if (boxList["E_OD2"][0][index] == ('', '')) or (boxList["E_ID2"][0][index] == ('', '')):
                #     messagebox.showerror(title="Value Error",message="Please fill all values first")
                #     return
                # boxList["E_OD2"][0][index] = (boxList["E_OD2"][0][index][0].get(),boxList["E_OD2"][0][index][0].get())
                # boxList["E_ID2"][0][index] = (boxList["E_ID2"][0][index][0].get(),boxList["E_ID2"][0][index][0].get())
                od1=od1Var.get()
                id1 = id1Var.get()
                # if (boxList["E_OD2"][0][index] is not None) and (boxList["E_ID2"][0][index] is not None) and od1 is not None and id1 is not None:
                # if (boxList["E_OD2"][0][index] == ('', '')) or (boxList["E_ID2"][0][index] == ('', '')) or od1 == '' or id1 == '':
                if not close_check and ((boxList["E_OD2"][0][index][1].get() == '') or (boxList["E_ID2"][0][index][1].get() == '') or od1 == '' or id1 == ''):
                    toproot.attributes('-topmost', True)
                    messagebox.showerror(title="Value Error",message="Please fill all values first",parent=toproot)
                    toproot.attributes('-topmost', False)
                    return
                elif close_check and ((boxList["E_OD2"][0][index][1].get() == '') or (boxList["E_ID2"][0][index][1].get() == '') or od1 == '' or id1 == ''):
                    toproot.attributes('-topmost', True)
                    messagebox.showerror(title="Closing with Blank",message="Please select type again as currently no data was provided",parent=toproot)
                    toproot.attributes('-topmost', False)
                    boxList['E_Type'][0][index][1].set("")
                    boxList['E_Grade'][0][index][1].set("")
                    boxList['E_Yield'][0][index][1].set("")
                    toproot.attributes('-topmost', False)
                    toproot.grab_release()
                    toproot.destroy()
                    
                else:
                    
                    boxList['E_OD1'][0][index][1].set(float(od1))
                    boxList['E_ID1'][0][index][1].set(float(id1))
                    
                    boxList["E_OD2"][0][index] = (boxList["E_OD2"][0][index][0].get(),boxList["E_OD2"][0][index][0].get())
                    boxList["E_ID2"][0][index] = (boxList["E_ID2"][0][index][0].get(),boxList["E_ID2"][0][index][0].get())
                    boxList['E_OD1'][0][index][0].configure(state='disabled')
                    boxList['E_ID1'][0][index][0].configure(state='disabled')
                    toproot.attributes('-topmost', False)
                    toproot.grab_release()
                    toproot.destroy()
                    boxList['E_Length'][0][index][0].focus()
                    check=True
            except Exception as e:
                raise e
        
        def on_closing():
            try:
                toproot.attributes('-topmost', True)
                if messagebox.askokcancel("Quit", "Do you want to quit?",parent=toproot):
                    toproot.attributes('-topmost', False)
                    close_check=True
                    exitTrue(close_check)
                toproot.attributes('-topmost', False)
            except Exception as e:
                raise e

        submitButton = tk.Button(submitFrame,text="Submit", command=exitTrue)
        submitButton.bind("<Return>", exitTrue)
        # submitButton.place(relx=.5, rely=.5, anchor="center")
        submitButton.grid(row=0,column=1,pady=40)
        toproot.focus()
        od1.focus()
        toproot.protocol("WM_DELETE_WINDOW", on_closing)
        if check:
            toproot.attributes('-topmost', False)
            toproot.grab_release()
            toproot.destroy()
            return od1.get(), id1.get(), boxList["E_OD2"][index][0].get(), boxList["E_ID2"][index][0].get()
        else:
            return None, None, None, None
    except Exception as e:
        raise e

def starSearch(root, df):
    try:
        toproot = tk.Toplevel(root, bg = "#9BC2E6")
        toproot.title('EAGS Quote Generator Star Search')
        screen_width = toproot.winfo_screenwidth()
        screen_height = toproot.winfo_screenheight()

        width = 315
        height = 280
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
        submitFrame = tk.Frame(toproot, bg= "#9BC2E6")
        submitFrame.grid(row=1, column=1)
        #Declaring Labels
        # tobePrinted = tk.Label(labelFrame, text="To be Printed", bg = "#9BC2E6")
        # tobePrinted.grid(row=0, column=0)

        # tobeCalculated = tk.Label(labelFrame, text="To be Calculated", bg = "#9BC2E6")
        # tobeCalculated.grid(row=0,column=1)

        gradeLabel = tk.Label(entryFrame1, text="Grade", bg = "#9BC2E6")
        gradeLabel.grid(row=0, column=0)

        yieldLabel = tk.Label(entryFrame1, text="Yield", bg = "#9BC2E6")
        yieldLabel.grid(row=0, column=1)

        odLabel = tk.Label(entryFrame2, text="OD", bg = "#9BC2E6")
        odLabel.grid(row=0, column=0)

        idLabel = tk.Label(entryFrame2, text="ID", bg = "#9BC2E6")
        idLabel.grid(row=0, column=1)

        #Configuring frame sizes based on screen size
        toproot.grid_rowconfigure(0, weight=1)
        toproot.grid_columnconfigure(0, weight=1)
        toproot.grid_rowconfigure(1, weight=1)
        toproot.grid_columnconfigure(1, weight=1)

        labelFrame.grid_rowconfigure(0, weight=1)
        labelFrame.grid_columnconfigure(0, weight=1)
        labelFrame.grid_rowconfigure(1, weight=1)
        labelFrame.grid_columnconfigure(1, weight=1)

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
        
        
        
        gradeVar = tk.StringVar()
        
        gradeBox = ttk.Entry(entryFrame1, textvariable=gradeVar, background = 'white',width = 10)
        gradeBox.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        gradeVar.set("*")

        yieldVar = tk.StringVar()
        yieldBox = ttk.Entry(entryFrame1, textvariable=yieldVar, background = 'white',width = 10)
        yieldBox.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        yieldVar.set("*")

        odVar = tk.StringVar()
        
        odBox = ttk.Entry(entryFrame2, textvariable=odVar, background = 'white',width = 10)
        odBox.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        odVar.set("*")

        idVar = tk.StringVar()
        idBox = ttk.Entry(entryFrame2, textvariable=idVar, background = 'white',width = 10)
        idBox.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        idVar.set("*")


        def starSearcher():
            try:
                
                grade = gradeVar.get()
                yieldField = yieldVar.get()
                od = odVar.get()
                idField = idVar.get()

                filtered_df = df.copy()

                #Filtering based on Grade
                if grade == "*":
                    pass

                
                elif "*" not in grade:
                    filtered_df  = filtered_df[ (filtered_df["grade"]==grade)]
                else:# gradeValue != "*":
                    filtered_df = filtered_df.loc[df["grade"].str.startswith(grade.replace('*',''))]

                if len(filtered_df)!=0:
                    #Filtering based on Yield
                    if yieldField == "*":
                            pass
                    elif "*" not in yieldField:
                        filtered_df  = filtered_df[ (filtered_df["heat_condition"]==yieldField)]
                    else: # yieldValue != "*":
                        filtered_df = filtered_df.loc[df["heat_condition"].str.startswith(yieldField.replace('*',''))]
                if len(filtered_df)!=0:
                    #Filtering based on od
                    if od == "*":
                            pass
                    elif "*" not in od:
                        filtered_df  = filtered_df[ (filtered_df["od_in"]==od)]
                    else: # yieldValue != "*":
                        filtered_df = filtered_df.loc[df["od_in"].str.startswith(od.replace('*',''))]
                if len(filtered_df)!=0:
                    #Filtering based on Yield
                    if idField == "*":
                            pass
                    elif "*" not in idField:
                        filtered_df  = filtered_df[ (filtered_df["od_in_2"]==idField)]
                    else: # yieldValue != "*":
                        filtered_df = filtered_df.loc[df["od_in_2"].str.startswith(idField.replace('*',''))]

                
                
                # if grade != "*" and yieldField != "*" and od != "*" and idField != "*":
                #     filtered_df = df.loc[
                #                         df["grade"].str.startswith(grade.replace('*','')) &
                #                         df["heat_condition"].str.startswith(yieldField.replace('*','')) &
                #                         df["od_in"].str.startswith(od.replace('*','')) &
                #                         df["od_in_2"].str.startswith(idField.replace('*',''))
                #                     ]
                    
                # elif grade != "*" and yieldField != "*" and od != "*":
                #     filtered_df  = df.loc[
                #                             df["grade"].str.startswith(grade.replace('*','')) &
                #                             df["heat_condition"].str.startswith(yieldField.replace('*','')) &
                #                             df["od_in"].str.startswith(od.replace('*',''))
                #                         ]
                # elif grade != "*" and yieldField != "*":
                #     filtered_df  = df.loc[
                #                             df["grade"].str.startswith(grade.replace('*','')) &
                #                             df["heat_condition"].str.startswith(yieldField.replace('*',''))
                #                         ]
                # elif grade != "*":
                #     filtered_df  = df.loc[
                #                             df["grade"].str.startswith(grade.replace('*',''))
                #                         ]
                # else:
                #     messagebox.showerror("Error", f"Please check search query and try again")
                #     return
                if len(filtered_df):
                    filtered_df = filtered_df[["grade", "heat_condition", "od_in","od_in_2",'age' ,'date_last_receipt','onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in']]
                    filtered_df = filtered_df.sort_values(["grade", "heat_condition", "od_in","od_in_2", "age"], ascending=[True, True, True, True, False])
                    
                    screen_width = toproot.winfo_screenwidth()
                    screen_height = toproot.winfo_screenheight()

                    width = 515
                    height = 480
                    # calculate position x and y coordinates
                    x = (screen_width/2) - (width/2)
                    y = (screen_height/2) - (height/2)
                    xlRoot = tk.Toplevel()
                    xlRoot.geometry('%dx%d+%d+%d' % (width, height, x, y))
                    xlRoot.state('zoomed')
                    

                    ptBakerxl = MyTable(xlRoot, editable=False,dataframe=filtered_df,showtoolbar=True, showstatusbar=True, maxcellwidth=2100)
                    ptBakerxl.font = 'Segoe UI'
                    ptBakerxl.fontsize = 12
                    ptBakerxl.thefont = ('Segoe UI', 12)
                    ptBakerxl.cellwidth = 130
                    ptBakerxl.show()
            
                else:
                    toproot.attributes('-topmost', True)
                    messagebox.showerror("Error", f"Please check search query and try again",parent=root)
                    toproot.attributes('-topmost', False)
                    return 
            except Exception as e:
                raise e




        submitButton = tk.Button(submitFrame,text="Submit", command=starSearcher)
        # submitButton.place(relx=.5, rely=.5, anchor="center")
        submitButton.grid(row=0,column=1,pady=40)
        toproot.focus()
        gradeBox.focus()

        gradeVar.set('*')
        yieldVar.set('*')
        odVar.set('*')
        idVar.set('*')
    except Exception as e:
        raise e


    

def rangeSearch(root, df, boxList, index, pt):
    try:
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
        
        #Checking if Quote Yes/No selected as Other or Blank or not
        
        
        toproot = tk.Toplevel(root, bg = "#9BC2E6")
        toproot.title('EAGS Quote Generator Range Search')
        screen_width = toproot.winfo_screenwidth()
        screen_height = toproot.winfo_screenheight()

        width = 315
        height = 280
        # calculate position x and y coordinates
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        toproot.geometry('%dx%d+%d+%d' % (width, height, x, y))
        

        ####################Condition for reasearcher to open################################################################
        #Find last entry row and logic for fill type column based on ID value if 0 then BR else THF
        last_row = len(boxList['E_Qty'][0]) - 1
        if boxList['C_Quote Yes/No'][0][last_row][0].get() != "Other" and boxList['C_Quote Yes/No'][0][last_row][0].get() != "Yes":
            # if boxList['C_Quote Yes/No'][0][index][0].get() != "":
            root.attributes('-topmost', True)
            messagebox.showerror("Error", f"Please Select Quote Yes/No as Other and try again",parent=root)
            root.attributes('-topmost', False)
            
            toproot.destroy()

            return
        ######################################################################################################################
        
        toproot.grab_set()
        labelFrame = tk.Frame(toproot, bg= "#9BC2E6")
        labelFrame.grid(row=0, column=1)
        boxFrame = tk.Frame(toproot, bg= "#9BC2E6")
        boxFrame.grid(row=1, column=1)
        entryFrame1 = tk.Frame(boxFrame, bg= "#9BC2E6")
        entryFrame1.grid(row=0, column=0)
        entryFrame2 = tk.Frame(boxFrame, bg= "#9BC2E6")
        entryFrame2.grid(row=0, column=1)
        submitFrame = tk.Frame(toproot, bg= "#9BC2E6")
        submitFrame.grid(row=2, column=1)
        # Declaring Labels
        locationLabel = tk.Label(labelFrame, text="Location", bg = "#9BC2E6")
        locationLabel.grid(row=0, column=0)
        boxList["searchLocation"][0][index] = customComboboxV2.myCombobox(df,toproot,item_list=["Dubai","Singapore","USA","UK"],frame=labelFrame,row=1,column=0,width=5,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = boxList)
        gradeLabel = tk.Label(labelFrame, text="Grade", bg = "#9BC2E6")
        gradeLabel.grid(row=0, column=1)

        # boxList["searchGrade"][0][index] = customComboboxV2.myCombobox(df,toproot,frame=labelFrame,row=1,column=1,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = boxList)
        
        searchGradeVar = tk.StringVar()
        
        searchGrade = ttk.Entry(labelFrame, textvariable=searchGradeVar, background = 'white',width = 10)
        searchGrade.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        searchGradeVar.set("*")

        yieldLabel = tk.Label(labelFrame, text="Yield", bg = "#9BC2E6")
        yieldLabel.grid(row=0,column=2)

        # boxList["searchYield"][0][index] = customComboboxV2.myCombobox(df,toproot,frame=labelFrame,row=1,column=2,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = boxList)
        searchYieldVar = tk.StringVar()
        
        searchYield = ttk.Entry(labelFrame, textvariable=searchYieldVar, background = 'white',width = 10)
        searchYield.grid(row=1,column=2,sticky=tk.EW,padx=5,pady=5)
        searchYieldVar.set("*")

        
        fromOdLabel = tk.Label(entryFrame1, text="From OD", bg = "#9BC2E6")
        fromOdLabel.grid(row=0, column=0)

        toOdLabel = tk.Label(entryFrame1, text="To OD", bg = "#9BC2E6")
        toOdLabel.grid(row=0, column=1)

        fromIdLabel = tk.Label(entryFrame2, text="From ID", bg = "#9BC2E6")
        fromIdLabel.grid(row=0, column=0)

        toIdLabel = tk.Label(entryFrame2, text="To ID", bg = "#9BC2E6")
        toIdLabel.grid(row=0, column=1)

        #Configuring frame sizes based on screen size
        toproot.grid_rowconfigure(0, weight=1)
        toproot.grid_columnconfigure(0, weight=1)
        toproot.grid_rowconfigure(1, weight=1)
        toproot.grid_columnconfigure(1, weight=1)

        labelFrame.grid_rowconfigure(0, weight=1)
        labelFrame.grid_columnconfigure(0, weight=1)
        labelFrame.grid_rowconfigure(1, weight=1)
        labelFrame.grid_columnconfigure(1, weight=1)
        labelFrame.grid_columnconfigure(2, weight=1)

        boxFrame.grid_rowconfigure(0, weight=1)
        boxFrame.grid_columnconfigure(0, weight=1)
        boxFrame.grid_rowconfigure(1, weight=1)
        boxFrame.grid_columnconfigure(1, weight=1)

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
        
        
        vcmd = toproot.register(intFloat)
        fromOdVar = tk.StringVar()

        #######
        
        fromOd = ttk.Entry(entryFrame1, textvariable=fromOdVar, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        fromOd['validatecommand'] = (fromOd.register(intFloat),'%P','%d')
        fromOd.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        # fromOdVar.set("*")

        toOdVar = tk.StringVar()
        toOd = ttk.Entry(entryFrame1, textvariable=toOdVar, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        toOd['validatecommand'] = (toOd.register(intFloat),'%P','%d')
        toOd.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        # yieldVar.set("*")

        fromIdVar = tk.StringVar()
        
        fromId = ttk.Entry(entryFrame2, textvariable=fromIdVar, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        fromId['validatecommand'] = (fromId.register(intFloat),'%P','%d')

        fromId.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        # odVar.set("*")

        toIdVar = tk.StringVar()
        toId = ttk.Entry(entryFrame2, textvariable=toIdVar, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        toId['validatecommand'] = (toId.register(intFloat),'%P','%d')

        toId.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        # idVar.set("*")


        def rangeSearcher():
            try:
                
                from_od_value = fromOdVar.get()
                to_od_value = toOdVar.get()
                from_id_value = fromIdVar.get()
                to_id_value = toIdVar.get()
                locationValue = boxList["searchLocation"][0][index][1].get()
                gradeValue = searchGradeVar.get()
                yieldValue = searchYieldVar.get()
                # (0.5 <= df['two']) & (df['two'] < 0.5)
                #Filtering based on Location
                if locationValue == '':
                    filtered_df = df.copy()
                else:
                    filtered_df = df[(df['site']==locationValue)]

                #Filtering based on Grade
                if gradeValue == "*":
                    pass

                
                elif "*" not in gradeValue:
                    filtered_df  = filtered_df[ (filtered_df["grade"]==gradeValue)]
                else:# gradeValue != "*":
                    filtered_df = filtered_df.loc[df["grade"].str.startswith(gradeValue.replace('*',''))]


                #Filtering based on Yield
                if yieldValue == "*":
                        pass
                elif "*" not in yieldValue:
                    filtered_df  = filtered_df[ (filtered_df["heat_condition"]==yieldValue)]
                else: # yieldValue != "*":
                    filtered_df = filtered_df.loc[df["heat_condition"].str.startswith(yieldValue.replace('*',''))]

                #Filtering OD
                if from_od_value == "" and to_od_value == "":
                    pass
                elif from_od_value != "" and to_od_value != "":
                    filtered_df = filtered_df[
                            (float(from_od_value) <= filtered_df['od_in'])  & (filtered_df['od_in'] <= float(to_od_value))
                            ]
                else:
                    toproot.attributes('-topmost', True)
                    messagebox.showerror("Error", f"Please check OD search query and try again",parent=toproot)
                    toproot.attributes('-topmost', False)
                    return


                #Filtering ID
                if from_id_value == "" and to_id_value == "":
                    pass
                elif from_id_value != "" and to_id_value != "":
                    filtered_df = filtered_df[
                            (float(from_id_value) <= filtered_df['od_in_2'])  & (filtered_df['od_in_2'] <= float(to_id_value))
                            ]
                else:
                    toproot.attributes('-topmost', True)
                    messagebox.showerror("Error", f"Please check ID search query and try again",parent=toproot)
                    toproot.attributes('-topmost', False)
                    return


                # if gradeValue != '' or yieldValue != '':
                #     if from_od_value != "" and to_od_value != "" and from_id_value != "" and to_id_value != "":
                #         filtered_df = df[
                #             (df["grade"]==gradeValue) & (df["heat_condition"]==yieldValue) &
                #             (float(from_od_value) <= df['od_in'])  & (df['od_in'] <= float(to_od_value)) &
                #             (float(from_id_value) <= df['od_in_2']) & (df['od_in_2'] <=  float(to_id_value))
                #             ]

                        
                #     elif from_od_value == "" and to_od_value == "" and from_id_value != "" and to_id_value != "":
                #         filtered_df  = df[ 
                #             (df["grade"]==gradeValue) & (df["heat_condition"]==yieldValue) &
                #             (float(from_id_value) <= df['od_in_2']) & (df['od_in_2'] <=  float(to_id_value))
                #             ]


                #     elif from_od_value != "" and to_od_value != "" and from_id_value == "" and to_id_value == "":
                #         filtered_df  = df[ 
                #             (df["grade"]==gradeValue) & (df["heat_condition"]==yieldValue) &
                #             (float(from_od_value) <= df['od_in']) & (df['od_in'] <=  float(to_od_value))
                #             ]
                    
                                            
                #     else:
                #         messagebox.showerror("Error", f"Please check search query and try again")
                #         return
                # else:
                #     messagebox.showerror("Error", f"Please fill grade and yield")
                    return
                if len(filtered_df):
                    filtered_df = filtered_df[["site", "grade", "heat_condition", "od_in","od_in_2",'age' ,'date_last_receipt','onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in', 'heat_number', 'lot_serial_number']]
                    filtered_df = filtered_df.sort_values(["site","grade", "heat_condition", "od_in","od_in_2", "age"], ascending=[True, True, True, True, True, False], ignore_index=True)
                    
                    screen_width = toproot.winfo_screenwidth()
                    screen_height = toproot.winfo_screenheight()

                    width = 515
                    height = 480
                    # calculate position x and y coordinates
                    x = (screen_width/2) - (width/2)
                    y = (screen_height/2) - (height/2)
                    xlRoot = tk.Toplevel(root, bg = "#9BC2E6")
                    xlRoot.geometry('%dx%d+%d+%d' % (width, height, x, y))
                    xlRoot.state('zoomed')
                    xlRoot.title("Range Search Table")
                    
                    def handle_double_click(e):
                        try:
                            rowclicked_single = ptBakerxl.get_row_clicked(e)
                            data = ptBakerxl.model.df['lot_serial_number'][rowclicked_single]
                            print(boxList)
                            map_dict = {"site":"E_Location", "grade":"E_Grade","heat_condition":"E_Yield", "od_in":"E_OD1","od_in_2":"E_ID1",
                             "onhand_dollars_per_pounds":"E_COST","lot_serial_number":"Lot_Serial_Number"}

                           

                             
                             
                            #Find last entry row and logic for fill type column based on ID value if 0 then BR else THF
                            last_row = len(boxList['E_Qty'][0]) - 1

                            #Fill quoteYes or no as other if n blank
                            # boxList['C_Quote Yes/No'][0][last_row][1].set("Other")

                            xlRoot.grab_release()
                            for key in map_dict.keys():
                                if key == "lot_serial_number":
                                    boxList['Lot_Serial_Number'][0][last_row] = (ptBakerxl.model.df[key][rowclicked_single], None)
                                elif key == "site":
                                    boxList[map_dict[key]][0][last_row][1].set(ptBakerxl.model.df[key][rowclicked_single])
                                    if ptBakerxl.model.df["od_in_2"][rowclicked_single] != 0.0:
                                        boxList["E_Type"][0][last_row][1].set("THF")
                                    else:
                                        boxList["E_Type"][0][last_row][1].set("BR")
                                else:
                                    boxList[map_dict[key]][0][last_row][1].set(ptBakerxl.model.df[key][rowclicked_single])

                            #################################Updating bottom table###################################################
                            newDf = df[(df["site"] == boxList['E_Location'][0][last_row][0].get()) & (df["grade"]==boxList['E_Grade'][0][last_row][0].get())
                                    & (df["heat_condition"]==boxList['E_Yield'][0][last_row][0].get())& (df["od_in"]==float(boxList['E_OD1'][0][last_row][0].get()))
                                    & (df["od_in_2"]==float(boxList['E_ID1'][0][last_row][0].get()))]

                            newDf = newDf[['onhand_pieces', 'onhand_length_in', 'onhand_dollars_per_pounds', 'available_pieces', 'available_length_in','date_last_receipt','age', 'heat_number', 'lot_serial_number']]
                            newDf['date_last_receipt'] = pd.to_datetime(newDf['date_last_receipt'])
                            newDf['date_last_receipt'] = newDf['date_last_receipt'].dt.date
                            newDf = newDf[newDf['available_pieces']>0]
                            newDf = newDf.sort_values('age', ascending=False).sort_values('date_last_receipt', ascending=True)
                            
                            #Resetting Index
                            newDf.reset_index(inplace=True, drop=True)
                            if pt is not None:
                                pt.model.df = newDf
                                pt.redraw()
                                
                            #########################################################################################################    
                                
                            
                            xlRoot.destroy()
                            print(f"Row clicked is {rowclicked_single+1}")
                        except Exception as ex:
                            raise ex
                    toproot.grab_release()
                    toproot.destroy()
                    xlRoot.grab_set()
                    ptBakerxl = MyTable(xlRoot, editable=False,dataframe=filtered_df,showtoolbar=True, showstatusbar=True, maxcellwidth=2100, width=2100)
                    ptBakerxl.font = 'Segoe UI'
                    ptBakerxl.fontsize = 12
                    ptBakerxl.cellwidth = 130
                    ptBakerxl.thefont = ('Segoe UI', 12)
                    ptBakerxl.show()
                    ptBakerxl.bind('<Double-Button-1>',handle_double_click)
                    

                    def on_closing():
                        try:
                            xlRoot.grab_release()
                            xlRoot.destroy()
                        except Exception as e:
                            raise e
                    
                    
                    xlRoot.protocol("WM_DELETE_WINDOW", on_closing)

            
                else:
                    toproot.attributes('-topmost', True)
                    messagebox.showerror("Error", f"Please check search query and try again",parent=toproot)
                    toproot.attributes('-topmost', False)
                    return
            except Exception as e:
                raise e




        submitButton = tk.Button(submitFrame,text="Submit", command=rangeSearcher)
        # submitButton.place(relx=.5, rely=.5, anchor="center")
        submitButton.grid(row=0,column=1,pady=40)
        toproot.focus()
        boxList["searchLocation"][0][index][0].focus()

        def on_closing():
            try:
                toproot.grab_release()
                toproot.destroy()
            except Exception as e:
                raise e
        
        
        toproot.protocol("WM_DELETE_WINDOW", on_closing)

        
    except Exception as e:
        raise e