
import win32com.client as win32
import pythoncom
import sys, time
import xlwings as xw
from tkinter import messagebox
from tkinter.filedialog import asksaveasfilename
import tkinter as tk
import pandas as pd
from datetime import datetime, date
import os
import ctypes

# use absolute pathes
WORKBOOK = "Sales Form Layout With Database.xlsm"
ENQSHEET = "Enquiry Form"
PDFSHEET = "Final Quote Layout"
TMPLTSHEET = "Final Quote Template"
# PDFSHEET = "DemoLayout"

root = tk.Tk()
root.withdraw()
entry = tk.Entry(root)
import traceback

#   store some stuff for win api interaction
set_to_foreground = ctypes.windll.user32.SetForegroundWindow
keybd_event = ctypes.windll.user32.keybd_event

alt_key = 0x12
extended_key = 0x0001
key_up = 0x0002

# focus_var = 0
def steal_focus():
    
    keybd_event(alt_key, 0, extended_key | 0, 0)
    set_to_foreground(root.winfo_id())
    keybd_event(alt_key, 0, extended_key | key_up, 0)

    entry.focus_set()


def main():
    # the button event will open this class

    class ButtonEvents:
        
        def __init__(self):
            self.keepOpen = True        
            self.myvar = 1
            self.buttonClicked = None
            
        # method executed on doubleclick to close while loop
        def OnDblClick(self, *args):
            print("button double clicked")
            self.keepOpen = False
            self.myvar = 3
        def OnClick(self, *args):
            print("Button Clicked")
            self.buttonClicked = True
             
    def get_fresh_template(wb, quote_sht):
        retry=0
        try:
            tmplt_sht = wb.sheets[TMPLTSHEET]
            wb.app.api.CutCopyMode=False
            quote_sht.activate()
            quote_sht.clear()
            try:
                quote_sht.pictures[0].delete()
            except:
                pass
            wb.app.api.CutCopyMode=False
            tmplt_sht.activate()
            tmplt_sht.range("A1:G32").select()
            wb.app.api.CopyObjectsWithCells = True
            wb.app.api.Selection.Copy()
            quote_sht.activate()    
            quote_sht.range("A1").api.Select()
            quote_sht.api.Paste()  #xlPasteAllUsingSourceTheme	
            wb.app.api.CutCopyMode=False

        except Exception as e:
            retry+=1
            if retry==3:
                raise e
    def excel_to_pdf(wb, pdf_loc):
        try:
        # Initialize new excel workbook
            time.sleep(2)
            # wb = xw.Book(file_loc, update_links=False)
            sheet =wb.sheets[PDFSHEET]
            wb.activate()
            sheet.activate()
            

            
            
            pdf_path = pdf_loc.replace("/","\\").replace(".pdf","")+".pdf"

            # Save excel workbook to pdf file
            print(f"Saving workbook as '{pdf_path}' ...")
            retry=0
            while retry<10:
                try:
                    time.sleep(1)
                    sheet.api.ExportAsFixedFormat(0, pdf_path)
                    break
                except Exception as e:
                    time.sleep(1)
                    retry+=1
                    if retry==10:
                        raise e
                    else:
                        continue

            
            print("Done")
        except Exception as e:
            raise e        

    def quote_generator(wb,xlEnq):
        try:
            rows_filled = xlEnq.range("A8").end("down").row - 6
            quote_sht = wb.sheets[PDFSHEET]
            cx_sht = wb.sheets["Customer_Database"]

            get_fresh_template(wb,quote_sht)
            if xlEnq.range("A9").value is not None:
                last_row = xlEnq.range("A8").end("down").row
                data = xlEnq.range(f"A8:U{last_row}").value
            else:
                data = xlEnq.range(f"A8:U8").value
            #Add extra rows in quote for more than 1 row data present
 
            df = cx_sht.range("E1").options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand='table').value ##df["cus_phone"] = int#
            loc_dict = {"Dubai":"$","Singapore":"$", "USA":"$","UK":"€","NA":""}
            city_dict = df.set_index(['cus_address'])["cus_city_zip"].to_dict()
            location = xlEnq.range(f"I8").value
            currency = loc_dict[location]
            phone_dict = df.set_index(['cus_address'])["cus_phone"].to_dict()
            
            #Basic Details
            #Company(CX Name)
            cx_name = xlEnq.range(f"B3").value
            #Address
            cx_add = xlEnq.range(f"B4").value
            #City,Zip
            # city_zip = xlEnq.range(f"N3").value + " ,123{current_row}5"
            city_zip = city_dict[cx_add]
            #Phone
            phone = phone_dict[cx_add]
            #Email
            email_add = xlEnq.range(f"B6").value
            #Filling Basic Details
            quote_sht.range("B2").value = cx_name
            quote_sht.range("B3").value = cx_add
            quote_sht.range("B4").value = city_zip
            quote_sht.range("B5").value = phone
            quote_sht.range("B6").value = email_add


            #Second Basic Details
            form_date = xlEnq.range(f"D2").value
            quote_no = int(datetime.strftime(form_date,"%m%d%Y")+"001") #mmddyyyy01
            vaildity = xlEnq.range(f"I3").value
            prep_by = xlEnq.range(f"B2").value
            pay_term = xlEnq.range(f"B5").value
            add_comments = xlEnq.range(f"I4").value

            #updating Report Sheet
            report_wb = xw.Book("Report.xlsx")
            report_sht = report_wb.sheets["Report"]
            if report_sht.range("A2").value == None:
                for i in range(len(data)):
                    report_sht.range(f"A{2+i}").value = quote_no
                # report_sht.range("B2").options(transpose = True).value = xlEnq.range(f"A{current_row}:AE{current_row}").value
                report_sht.range("B2").options(transpose = True).value = data
            else:
                latest_row = report_sht.range("A1").end('down').row
                prev_quote_no = int(report_sht.range(f"A{latest_row}").value)
                if str(prev_quote_no)[1:3] == str(quote_no)[1:3]:
                    quote_no = prev_quote_no + 1
                if len(data)!=21:
                    for i in range(len(data)):
                        report_sht.range(f"A{latest_row+1+i}").value = quote_no
                        report_sht.range(f"B{latest_row+1+i}").value = xlEnq.range("B2").value #For multiple entries
                        report_sht.range(f"C{latest_row+1+i}").value = xlEnq.range("D2").value #For multiple entries
                        report_sht.range(f"D{latest_row+1+i}").value = xlEnq.range("B3").value #For multiple entries
                        report_sht.range(f"E{latest_row+1+i}").value = xlEnq.range("B4").value #For multiple entries
                        report_sht.range(f"F{latest_row+1+i}").value = xlEnq.range("B5").value #For multiple entries
                        if i==0:
                            report_sht.range(f"G{latest_row+1+i}").value = data #For multiple entries
                        report_sht.range(f"AB{latest_row+1+i}").value=xlEnq.range("I3").value
                        report_sht.range(f"AC{latest_row+1+i}").value=xlEnq.range("I4").value
                else:
                    report_sht.range(f"A{latest_row+1}").value = quote_no
                # report_sht.range(f"B{latest_row+1}").value = xlEnq.range(f"A{current_row}:AE{current_row}").value
                    report_sht.range(f"B{latest_row+1}").value = xlEnq.range("B2").value #For multiple entries
                    report_sht.range(f"C{latest_row+1}").value = xlEnq.range("D2").value #For multiple entries
                    report_sht.range(f"D{latest_row+1}").value = xlEnq.range("B3").value #For multiple entries
                    report_sht.range(f"E{latest_row+1}").value = xlEnq.range("B4").value #For multiple entries
                    report_sht.range(f"F{latest_row+1}").value = xlEnq.range("B5").value #For multiple entries
                    report_sht.range(f"G{latest_row+1}").value = data #For multiple entries
                    report_sht.range(f"AB{latest_row+1}").value=xlEnq.range("I3").value
                    report_sht.range(f"AC{latest_row+1}").value=xlEnq.range("I4").value
                report_sht.autofit()
            report_wb.save()
            report_wb.close()

            #Filling 2nd Basic Details
            quote_sht.range("F2").value = quote_no
            quote_sht.range("F3").value = form_date
            quote_sht.range("F4").value = vaildity
            quote_sht.range("F5").value = prep_by
            quote_sht.range("F6").value = pay_term
            # quote_sht.range("A32").value = add_comments
            #Logic for multiple pages
            # for j in range(len(data)//5):
            # #Adding logic for multiple entries
            if len(data)==21:
                m_range = 1
            else:
                m_range = len(data)
            for i in range(m_range):
                    #Size Variables
                    cx_od = xlEnq.range(f"D{8+i}").value
                    cx_id = xlEnq.range(f"E{8+i}").value
                    #CX Grade Variables
                    cx_grade = xlEnq.range(f"B{8+i}").value
                    cx_yield = xlEnq.range(f"C{8+i}").value
                    #CX Spec Varaibles
                    cx_spec = xlEnq.range(f"A{8+i}").value
                    #CX Qty Variables
                    cx_qty = xlEnq.range(f"G{8+i}").value
                    cx_len = xlEnq.range(f"F{8+i}").value
                    if i != 0:
                        row_range, data_rows = data_adder(quote_sht)
                    if i ==0:
                        #CX Size
                        quote_sht.range("B9").value = f'{cx_od}I - {cx_id}I'
                        #CX Grade
                        quote_sht.range("B10").value = f'{cx_grade}-{cx_yield}'
                        #CX Spec
                        quote_sht.range("B11").value = cx_spec
                        #CX Qty
                        quote_sht.range("B12").value = f'{cx_qty} @ {cx_len}"'
                    else:
                        #CX Size
                        # quote_sht.range(f"B{9+row_range+i+1}").value = f'{cx_od}I - {cx_id}I'
                        # #CX Grade
                        # quote_sht.range(f"B{10+row_range+i+1}").value = f'{cx_grade}-{cx_yield}'
                        # #CX Spec
                        # quote_sht.range(f"B{11+row_range+i+1}").value = cx_spec
                        # #CX Qty
                        # quote_sht.range(f"B{12+row_range+i+1}").value = f'{cx_qty} @ {cx_len}"'
                        #CX Size
                        quote_sht.range(f"B{data_rows}").value = f'{cx_od}I - {cx_id}I'
                        #CX Grade
                        quote_sht.range(f"B{1 + data_rows}").value = f'{cx_grade}-{cx_yield}'
                        #CX Spec
                        quote_sht.range(f"B{2 + data_rows}").value = cx_spec
                        #CX Qty
                        quote_sht.range(f"B{3 + data_rows}").value = f'{cx_qty} @ {cx_len}"'
                        
        
                    #EAGS details
                    if xlEnq.range(f'H{8+i}').value == 'No':
                        #Size Variable
                    
                        eags_od = "NA"
                        eags_id = "NA"
                        
                        #CX Grade Variables
                        eags_grade = "NA"
                        eags_yield = "NA"
                        #CX Spec Varaibles
                        eags_spec = "NA"
                        #CX Qty Variables
                        eags_qty = "NA"
                        eags_len = "NA"
                    else:
                        if i ==0:

                            #Size Variable
                            
                            eags_od = xlEnq.range(f"M8").value
                            eags_id = xlEnq.range(f"N8").value
                            
                            #CX Grade Variables
                            eags_grade = xlEnq.range(f"K8").value
                            eags_yield = xlEnq.range(f"L8").value
                            #CX Spec Varaibles
                            eags_spec = xlEnq.range(f"A8").value
                            #CX Qty Variables
                            eags_qty = xlEnq.range(f"P8").value
                            eags_len = xlEnq.range(f"O8").value
                            
                        else:
                            #Size Variable
                            
                            eags_od = xlEnq.range(f"M{8+i}").value
                            eags_id = xlEnq.range(f"N{8+i}").value
                            
                            #CX Grade Variables
                            eags_grade = xlEnq.range(f"K{8+i}").value
                            eags_yield = xlEnq.range(f"L{8+i}").value
                            #CX Spec Varaibles
                            eags_spec = xlEnq.range(f"A{8+i}").value
                            #CX Qty Variables
                            eags_qty = xlEnq.range(f"P{8+i}").value
                            eags_len = xlEnq.range(f"O{8+i}").value
                            

                    #Filling EAGS Offer Details
                    if xlEnq.range(f'H{8+i}').value == 'No':
                        if i ==0:
                            #EAEAGS Size
                            quote_sht.range(f"B14").value = "NA"
                            #EAGS Grade
                            quote_sht.range(f"B15").value = "NA"
                            #EAGS Spec
                            quote_sht.range(f"B16").value = "NA"
                            #EAGS Qty
                            quote_sht.range(f"B17").value = "NA"
                        else:
                            # #EAEAGS Size
                            # quote_sht.range(f"B{14+row_range+i+1}").value = "NA"
                            # #EAGS Grade
                            # quote_sht.range(f"B{15+row_range+i+1}").value = "NA"
                            # #EAGS Spec
                            # quote_sht.range(f"B{16+row_range+i+1}").value = "NA"
                            # #EAGS Qty
                            # quote_sht.range(f"B{17+row_range+i+1}").value = "NA"
                            #EAEAGS Size
                            quote_sht.range(f"B{5 + data_rows}").value = "NA"
                            #EAGS Grade
                            quote_sht.range(f"B{6 + data_rows}").value = "NA"
                            #EAGS Spec
                            quote_sht.range(f"B{7 + data_rows}").value = "NA"
                            #EAGS Qty
                            quote_sht.range(f"B{8 + data_rows}").value = "NA"
                            
                    else:
                        if i ==0:
                            #EAEAGS Size
                            quote_sht.range(f"B14").value = f'{eags_od}I-{eags_id}I'
                            #EAGS Grade
                            quote_sht.range(f"B15").value = f'{eags_grade}-{eags_yield}'
                            #EAGS Spec
                            quote_sht.range(f"B16").value = eags_spec
                            #EAGS Qty
                            quote_sht.range(f"B17").value = f'{eags_qty} @ {eags_len}"'
                        else:
                            # #EAEAGS Size
                            # quote_sht.range(f"B{14+row_range+i}").value = f'{eags_od}I-{eags_id}I'
                            # #EAGS Grade
                            # quote_sht.range(f"B{15+row_range+i}").value = f'{eags_grade}-{eags_yield}'
                            # #EAGS Spec
                            # quote_sht.range(f"B{16+row_range+i}").value = eags_spec
                            # #EAGS Qty
                            # quote_sht.range(f"B{17+row_range+i}").value = f'{eags_qty} @ {eags_len}"'
                            #EAEAGS Size
                            quote_sht.range(f"B{5 + data_rows}").value = f'{eags_od}I-{eags_id}I'
                            #EAGS Grade
                            quote_sht.range(f"B{6 + data_rows}").value = f'{eags_grade}-{eags_yield}'
                            #EAGS Spec
                            quote_sht.range(f"B{7 + data_rows}").value = eags_spec
                            #EAGS Qty
                            quote_sht.range(f"B{8 + data_rows}").value = f'{eags_qty} @ {eags_len}"'
                    #EAGS Heat
                    # quote_sht.range("B18").value = eags_heat
                    # #EAGS TAG
                    # quote_sht.range("B18").value = eags_tag

                    #Filling other columns
                    if xlEnq.range(f'H{8+i}').value == 'No':
                        if i==0:
                            #Qty
                            quote_sht.range(f"C13").value = "NA"
                            #UOM
                            quote_sht.range(f"D13").value = "NA"
                            #Price
                            
                            price = xlEnq.range(f"U{8+i}").value
                            quote_sht.range(f"E13").value = "NA"
                            #Amount
                            amount = "NA"
                            quote_sht.range(f"F13").value = "NA"
                            #Price Term
                            quote_sht.range(f"G13").value = "NA"
                            quote_sht.range(f"G14").value = None
                        else:
                            #Qty
                            # quote_sht.range(f"C{13+row_range+i+1}").value = "NA"
                            # #UOM
                            # quote_sht.range(f"D{13+row_range+i+1}").value = "NA"
                            # #Price
                            
                            # price = xlEnq.range(f"U{current_row}").value
                            # quote_sht.range(f"E{13+row_range+i+1}").value = "NA"
                            # #Amount
                            # amount = "NA"
                            # quote_sht.range(f"F{13+row_range+i+1}").value = "NA"
                            # #Price Term
                            # quote_sht.range(f"G{13+row_range+i+1}").value = "NA"
                            # quote_sht.range(f"G{14+row_range+i+1}").value = None
                            #Qty
                            quote_sht.range(f"C{4 + data_rows}").value = "NA"
                            #UOM
                            quote_sht.range(f"D{4 + data_rows}").value = "NA"
                            #Price
                            
                            price = xlEnq.range(f"U{current_row}").value
                            quote_sht.range(f"E{4 + data_rows}").value = "NA"
                            #Amount
                            amount = "NA"
                            quote_sht.range(f"F{4 + data_rows}").value = "NA"
                            #Price Term
                            quote_sht.range(f"G{4 + data_rows}").value = "NA"
                            quote_sht.range(f"G{5 + data_rows}").value = None
                    else:
                        if i ==0:
                            #Qty
                            quote_sht.range(f"C13").value = f'{eags_qty}PC @ {eags_len}"'
                            #UOM
                            quote_sht.range(f"D13").value = xlEnq.range(f"R8").value
                            #Price
                            
                            price = xlEnq.range(f"U{8+i}").value
                            quote_sht.range(f"E13").value = f"{price}"
                            quote_sht.range(f"E13").api.NumberFormat = "[$$-en-CA]#,##0.00;[Red]-[$$-en-CA]#,##0.00" 
                            #Amount
                            amount = round((float(price)*int(eags_qty)),2)
                            quote_sht.range(f"F13").value = f"{amount}"
                            quote_sht.range(f"F13").api.NumberFormat = "[$$-en-CA]#,##0.00;[Red]-[$$-en-CA]#,##0.00" 
                            #Price Term
                            quote_sht.range(f"G13").value = "Ex-works"
                            quote_sht.range(f"G14").value = xlEnq.range(f"I8").value
                            
                        else:
                            # #Qty
                            # quote_sht.range(f"C{13+row_range+i+1}").value = f'{eags_qty}PC @ {eags_len}"'
                            # #UOM
                            # quote_sht.range(f"D{13+row_range+i+1}").value = xlEnq.range(f"R{8+i}").value
                            # #Price
                            
                            # price = xlEnq.range(f"U{8+i}").value
                            # quote_sht.range(f"E{13+row_range+i+1}").value = f"{price} {currency}"
                            # #Amount
                            # amount = round((float(price)*int(eags_qty)),2)
                            # quote_sht.range(f"F{13+row_range+i+1}").value = f"{amount} {currency}"
                            # #Price Term
                            # quote_sht.range(f"G{13+row_range+i+1}").value = "Ex-works"
                            # quote_sht.range(f"G{14+row_range+i+1}").value = xlEnq.range(f"I{8+i}").value
                            #Qty
                            quote_sht.range(f"C{4 + data_rows}").value = f'{eags_qty}PC @ {eags_len}"'
                            #UOM
                            quote_sht.range(f"D{4 + data_rows}").value = xlEnq.range(f"R{8+i}").value
                            #Price
                            
                            price = xlEnq.range(f"U{8+i}").value
                            quote_sht.range(f"E{4 + data_rows}").value = f"{price}"
                            #Amount
                            amount = round((float(price)*int(eags_qty)),2)
                            quote_sht.range(f"F{4 + data_rows}").value = f"{amount}"
                            #Price Term
                            quote_sht.range(f"G{4 + data_rows}").value = "Ex-works"
                            quote_sht.range(f"G{5 + data_rows}").value = xlEnq.range(f"I{8+i}").value
                            
            
            #Deleting extra blank rows
            last_data_row = quote_sht.range("A8").end("down").row + 1
            last_blank_row = quote_sht.range(f"A{last_data_row}").end("down").row - 2
            quote_sht.range(f"A{last_data_row}:G{last_blank_row + 1}").delete()
            quote_last_row = quote_sht.range(f'A' + str(quote_sht.cells.last_cell.row)).end('up').row
            quote_sht.range(f"A{quote_last_row + 1}").value = add_comments 

            files = [('pdf', '*.pdf')]
            root.lift()
            root.overrideredirect(True)
            root.attributes("-topmost", True)
            root.attributes("-alpha", 0)
            pdf_loc = asksaveasfilename(filetypes = files, defaultextension = files, initialfile=quote_no)
            return pdf_loc
        except Exception as e:
            raise e

    def formula_calculator():
        retry=0
        try:
            #check columns before applying forumlas
            numeric_col_checker()
            #Calculating Wall thickness(R4-S4)/2 R4=OD, S4=ID
            if xlEnq.range(f'M{current_row}').value != None and xlEnq.range(f'N{current_row}').value != None and xlEnq.range(f'M{current_row}').value != "NA" and xlEnq.range(f'N{current_row}').value != "NA":
                if xlEnq.range(f'Q{current_row}').value != None and xlEnq.range(f'O{current_row}').value != None and xlEnq.range(f'Q{current_row}').value != "NA" and xlEnq.range(f'O{current_row}').value != "NA":
                    od = float(xlEnq.range(f'M{current_row}').value)
                    id = float(xlEnq.range(f'N{current_row}').value)
                    if xlEnq.range(f'J{current_row}').value == "THF":
                        
                        wt = (od-id)/2 #V column earlier
                        #=((R4-V4)*V4*10.68)/12
                        mid_formula = ((od-wt)*wt*10.68)/12 #W column earlier
                        lbs_sell_cost = float(xlEnq.range(f'Q{current_row}').value)
                        len_col = float(xlEnq.range(f'O{current_row}').value)
                        #Selling cost/UOM =X5*W5*T5 SellingCost/LBS * mid_formula * Length
                        if xlEnq.range(f'R{current_row}').value == "Each":
                            uom_sell_cost = round((lbs_sell_cost * mid_formula * len_col),2)
                        else:
                            uom_sell_cost = round((lbs_sell_cost * mid_formula),2)
                            
                        #Filling Calculated values
                        #Selling cost/UOM
                        xlEnq.range(f'S{current_row}').value = uom_sell_cost
                        #Final Price
                        xlEnq.range(f'U{current_row}').formula = f"=T{current_row}+S{current_row}"
                    elif xlEnq.range(f'J{current_row}').value == "BR":
                        mid_formula = (od*od*2.71)/12 #W column earlier
                        lbs_sell_cost = float(xlEnq.range(f'Q{current_row}').value)
                        len_col = float(xlEnq.range(f'O{current_row}').value)
                        #Selling cost/UOM =X5*W5*T5 SellingCost/LBS * mid_formula * Length
                        if xlEnq.range(f'R{current_row}').value == "Each":
                            uom_sell_cost = round((lbs_sell_cost * mid_formula * len_col),2)
                        else:
                            uom_sell_cost = round((lbs_sell_cost * mid_formula),2)
                            
                        #Filling Calculated values
                        #Selling cost/UOM
                        xlEnq.range(f'S{current_row}').value = uom_sell_cost
                        #Final Price
                        xlEnq.range(f'U{current_row}').formula = f"=T{current_row}+S{current_row}"
                else:
                    #Filling Calculated values
                    #Selling cost/UOM
                    if xlEnq.range(f'S{current_row}').value != None:
                        xlEnq.range(f'S{current_row}').value = None
                        #Final Price
                        xlEnq.range(f'U{current_row}').formula = None
        except Exception as e:
            if retry == 2:
                raise e
            else:
                retry+=1
    def data_adder(quote_sht):
        # quote_sht =  wb.sheets["Sheet2"]
        first_row = 8
        data_rows = 9
        row_range = 8
        quote_sht.activate()
        wb.app.api.CutCopyMode=False
        last_row = 17
        data = quote_sht.range(f"A{first_row}:G{last_row}")
        data.api.Select()
        wb.app.selection.api.Copy()
        
        #Before pasting calculating last row again
        last_row = quote_sht.range(f"A{first_row}").end("down").row
        quote_sht.range(f"A{last_row+1}").api.Select()
        wb.app.selection.api.Insert(Shift=-4121)
        wb.app.api.CutCopyMode=False
        
        last_row = quote_sht.range(f"A{first_row}").end("down").row
        #if we are pasting fro second time
        if last_row !=17:
            first_row = last_row - 9
            data_rows = last_row - 8
        row_range = last_row-data_rows
        
        
        return row_range, data_rows

    def list_formula_updater():
        inv_sht = wb.sheets["Inventory_Lists"]
        cx_sht = wb.sheets["Customer_Database"]
        #Pivot last row
        p_last_row = inv_sht.range(f'V' + str(inv_sht.cells.last_cell.row)).end('up').row
        #updating current row in list formula
        inv_sht.api.Range("I2").Formula2 = f'=UNIQUE(FILTER(J2:J3,ISNUMBER(SEARCH(\'Enquiry Form\'!R{current_row},J2:J3)),"Not found"))'#I2
        inv_sht.api.Range("K2").Formula2 = f'=UNIQUE(FILTER(V3:V{p_last_row},ISNUMBER(SEARCH(\'Enquiry Form\'!I{current_row},V3:V{p_last_row})),"Not found"))'#K2
        inv_sht.api.Range("L2").Formula2 = f'=UNIQUE(FILTER(W3:W{p_last_row},V3:V{p_last_row}=\'Enquiry Form\'!I{current_row},""))'#L2
        inv_sht.api.Range("M2").Formula2 = f'=FILTER(L2#,ISNUMBER(SEARCH(\'Enquiry Form\'!J{current_row},L2#)),"Not found")'#M2
        inv_sht.api.Range("N2").Formula2 = f'=UNIQUE(FILTER(X3:X{p_last_row},((W3:W{p_last_row}=\'Enquiry Form\'!J{current_row})*(V3:V{p_last_row}=\'Enquiry Form\'!I{current_row})),""))'#N2
        inv_sht.api.Range("O2").Formula2 = f'=FILTER(N2#,ISNUMBER(SEARCH(\'Enquiry Form\'!K{current_row},N2#)),"Not found")'#O2
        inv_sht.api.Range("P2").Formula2 = f'=UNIQUE(FILTER(Y3:Y{p_last_row},((W3:W{p_last_row}=\'Enquiry Form\'!J{current_row})*(V3:V{p_last_row}=\'Enquiry Form\'!I{current_row})*(X3:X{p_last_row}=\'Enquiry Form\'!K{current_row})),""))'#P2
        inv_sht.api.Range("Q2").Formula2 = f'=FILTER(P2#,ISNUMBER(SEARCH(\'Enquiry Form\'!L{current_row},P2#)),"Not found")'#Q2
        inv_sht.api.Range("R2").Formula2 = f'=UNIQUE(FILTER(Z3:Z{p_last_row},((W3:W{p_last_row}=\'Enquiry Form\'!J{current_row})*(V3:V{p_last_row}=\'Enquiry Form\'!I{current_row})*(X3:X{p_last_row}=\'Enquiry Form\'!K{current_row})*(Y3:Y{p_last_row}=\'Enquiry Form\'!L{current_row})),""))'#R2
        inv_sht.api.Range("S2").Formula2 = f'=FILTER(R2#,ISNUMBER(SEARCH(\'Enquiry Form\'!M{current_row},R2#)),"Not found")'#S2
        inv_sht.api.Range("T2").Formula2 = f'=UNIQUE(FILTER(AA3:AA{p_last_row},((W3:W{p_last_row}=\'Enquiry Form\'!J{current_row})*(V3:V{p_last_row}=\'Enquiry Form\'!I{current_row})*(X3:X{p_last_row}=\'Enquiry Form\'!K{current_row})*(Y3:Y{p_last_row}=\'Enquiry Form\'!L{current_row})*(Z3:Z{p_last_row}=\'Enquiry Form\'!M{current_row})),""))'#T2
        inv_sht.api.Range("U2").Formula2 = f'=FILTER(T2#,ISNUMBER(SEARCH(\'Enquiry Form\'!N{current_row},T2#)),"Not found")'#U2


        cx_sht.api.Range("P2").Formula2 = f'=UNIQUE(FILTER(Q2:Q4,ISNUMBER(SEARCH(\'Enquiry Form\'!H{current_row},Q2:Q4)),"Not found"))'
    
    def numeric_col_checker():
        """check if numeric cols are filled with numeric value otherwise raise error
        """
        # numeric_cols = ["I", "J", "K", "L", "R", "S", "T", "U", "V", "Y"]
        numeric_cols = ["D", "E", "F", "G", "M", "N", "O", "P", "Q", "T"]
        for col in numeric_cols:
            if current_row>7:
                if (xlEnq.range(f"{col}{current_row}").value is not None and xlEnq.range(f"{col}{current_row}").value != "NA") and not(isinstance(xlEnq.range(f"{col}{current_row}").value, (int, float))):
                    root.attributes("-topmost", True)
                    #   after 0.5 seconds focus will be stolen
                    root.after(500, steal_focus)                
                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in {col}{current_row}",parent=entry)
                    xlEnq.range(f'{col}{current_row}').value = None
                    wb.activate(steal_focus=True)
                    xlEnq.range(f'{col}{current_row}').api.Activate()
                elif (col == "G" or col == "P") and xlEnq.range(f"{col}{current_row}").value is not None and  xlEnq.range(f"{col}{current_row}").value != "NA" and isinstance(xlEnq.range(f"{col}{current_row}").value, (int, float)): #show error msg if q is not int
                    if not(int(xlEnq.range(f"{col}{current_row}").value) == xlEnq.range(f"{col}{current_row}").value):
                        root.attributes("-topmost", True)
                        #   after 0.5 seconds focus will be stolen
                        root.after(500, steal_focus)               
                        messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in {col}{current_row}",parent=entry)
                        xlEnq.range(f'{col}{current_row}').value = None
                        wb.activate(steal_focus=True)
                        xlEnq.range(f'{col}{current_row}').api.Activate()
                else:
                    pass

    def pivot_refresher(wb):
        cx_db_sht = wb.sheets["Customer_Database"]
        inv_sht = wb.sheets["Inventory_Lists"]
        sp_db_sht = wb.sheets["Salesperson_Database"]
        cx_sht = wb.sheets["Customer_Database"]

        for sht in [cx_db_sht, inv_sht, sp_db_sht, cx_sht]:
            pivotCount = sht.api.PivotTables().Count
        
            
            for j in range(1, pivotCount+1):
                # sht.PivotTables(j).PivotCache().SourceData = f"'{inp_set_sht.name}'!R1C1:R{inp_set_last_row}C13" #13 for M col
                sht.api.PivotTables(j).PivotCache().Refresh()
        pass
##############NOW EDIT TABOVE FUNCTIONS
        
    try:
        root.after(500, steal_focus) 
          
        root.attributes("-topmost", True)
        messagebox.showinfo("Starting Process", "Opening Excel")
        root.after(500, steal_focus)
        wb = xw.Book(WORKBOOK)
        pivot_refresher(wb)
        xlEnq=wb.sheets[ENQSHEET]
        wb.activate(steal_focus=True)
        xlEnq.activate()
        xlEnq.range("B2:B5").value = None
        xlEnq.range("I4:I5").value = None
        xlEnq.range("D2").value = None
        xlEnq.range("B2").api.Activate()
        xlEnq.range("A8:U12").value = None
        xlEnq.range("I4").value = "NA"
        # define button event callback class
        xlButtonEvents=win32.WithEvents(xlEnq.api.OLEObjects("SubmitButton").Object,ButtonEvents)
        
        # eags_col_list = ["N","M"]
        
        # a while loop to wait until the button in excel is double-clicked
        keepOpen=True
        row =1
        state_dict = {}
        for key in range(8,13):
            state_dict[key] = 'Initial_State'
        # current_state = ["Initial_State", row]
        retrial = 0
        while keepOpen:
            try:
                if wb.name == 'Sales Form Layout With Database.xlsm':
                    if wb.app.api.ActiveSheet.Name == 'Enquiry Form':
                        current_row = wb.app.selection.row
                        if current_row <2 or current_row >=13:
                            current_row=2
                            continue
                        pass
                    else:
                        continue
                
                if xlButtonEvents.buttonClicked is not None:
                    if xlEnq.range("S8").value != 0 and xlEnq.range("A8").value != None:
                        pdf_loc = quote_generator(wb,xlEnq)
                        if pdf_loc!=None and pdf_loc!='':
                            excel_to_pdf(wb, pdf_loc)
                            xlButtonEvents.buttonClicked = None
                        else:
                           wb.activate()
                           xlEnq.activate()
                           xlButtonEvents.buttonClicked = None
                           continue
                else:
                    # date_row =  xlEnq.range('A2').end('down').row
                    # cx_row = xlEnq.range('C2').end('down').row
                    # cx_row = '4'
                    if wb.app.selection.sheet.name == 'Enquiry Form':
                        current_row = wb.app.selection.row
                        if current_row <2 or current_row >=13:
                            current_row=2
                            continue
                    else:
                        continue
                    #Condition for moving to validity and addtional details and then to 2nd row spec
                    if current_row == 8 and wb.app.selection.column == 21 and xlEnq.range('R8').value != None:
                        try:
                            xlEnq.range(f'I3').api.Activate()
                        except:
                            pass
                    elif current_row == 5 and wb.app.selection.column == 9 and xlEnq.range('I4').value != "NA":
                        try:
                            xlEnq.range(f'A9').api.Activate()
                        except:
                            pass

                    if current_row<=6:
                        if  (current_row==2 or current_row==3)and xlEnq.range('B2').value != None:
                        
                            if xlEnq.range(f'D2').value == None:
                                xlEnq.range(f'D2').value = date.today()
                        if wb.app.selection.column==2 and current_row==3:#B3 for cx name
                            if xlEnq.range(f'B4').value == None and xlEnq.range(f'B4').formula == '':
                                xlEnq.range(f'B4').formula = f"=XLOOKUP(B3,Customer_Database!$C$2:$C$141,Customer_Database!$E$2:$E$141)"
                                xlEnq.range(f'B5').formula = f"=XLOOKUP(B3,Customer_Database!$C$2:$C$141,Customer_Database!$D$2:$D$141)"
                                xlEnq.range(f'B6').formula = f"=XLOOKUP(B3,Customer_Database!$C$2:$C$141,Customer_Database!$H$2:$H$141)"
                        if xlEnq.range("A8").value == None and current_row != 3 and xlEnq.range(f'B4').value != None and xlEnq.range(f'B4').value != 0 and xlEnq.range(f'B2').value != None:
                            try:
                                xlEnq.range("A8").api.Activate()
                            except:
                                pass
                        #Condition for Prepared by
                        if (xlEnq.range(f'B2').value != None):
                            sp_db_sht = wb.sheets["Salesperson_Database"]
                            if sp_db_sht.range(f'C2').value != 'Not found':
                                try:
                                    current_lst = sp_db_sht.range(f'C2').expand('down')
                                except:
                                    time.sleep(0.5)
                                    current_lst = sp_db_sht.range(f'C2').expand('down')
                                if len(current_lst)==1:
                                    # if prev_val[key] != xlEnq.range(f'{key}{current_row}').value and xlEnq.range(f'{key}{current_row}').value != None:
                                    # m_prev_val == xlEnq.range(f'{key}{current_row}').value
                                    if xlEnq.range(f'B2').value != sp_db_sht.range(f'C2').value:
                                        xlEnq.range(f'B2').value = sp_db_sht.range(f'C2').value
                            elif sp_db_sht.range(f'C2').value == 'Not found' and xlEnq.range(f'B2').value != None:
                                root.lift()
                                root.overrideredirect(True)
                                root.attributes("-topmost", True)
                                root.attributes("-alpha", 0)
                                #   after 0.5 seconds focus will be stolen
                                root.after(500, steal_focus)  
                                messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in B2", parent=entry)
                                xlEnq.range(f'B2').value = None
                                wb.activate(wb.activate(steal_focus=True))
                                
                                xlEnq.range(f'B2').api.Activate()
                        #Condition for Customer
                        if (xlEnq.range(f'B3').value != None):
                            cx_db_sht = wb.sheets["Customer_Database"]
                            if cx_db_sht.range(f'J2').value != 'Not found':
                                current_lst = cx_db_sht.range(f'J2').expand('down')
                                if len(current_lst)==1:
                                    # if prev_val[key] != xlEnq.range(f'{key}{current_row}').value and xlEnq.range(f'{key}{current_row}').value != None:
                                    # m_prev_val == xlEnq.range(f'{key}{current_row}').value
                                    xlEnq.range(f'B3').value = cx_db_sht.range(f'J2').value
                            elif cx_db_sht.range(f'J2').value == 'Not found' and xlEnq.range(f'B3').value != None:
                                root.lift()
                                root.overrideredirect(True)
                                root.attributes("-topmost", True)
                                # root.attributes("-alpha", 0)
                                #   after 0.5 seconds focus will be stolen
                                root.after(500, steal_focus)
                                messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in B3",parent=entry)
                                xlEnq.range(f'B3').value = None
                                wb.activate(steal_focus=True)
                                xlEnq.range(f'B3').api.Activate()

                    
                    else:
                        list_formula_updater()
                        #Yes/No Condition
                        if current_row >7:
                            if xlEnq.range(f'H{current_row}').value == 'No':
                                if state_dict[current_row] != "No":
                                    state_dict[current_row] = "No"
                                    for i in range(ord('I'),ord('U')+1):#N till AC
                                        try:
                                            xlEnq.range(f'{chr(i)}{current_row}').api.Validation.Delete()   
                                        except:
                                            pass
                                        xlEnq.range(f'{chr(i)}{current_row}').value = "NA"
                                    # xlEnq.range(f'V4').value = "NA"#Validity 
                                    # xlEnq.range(f'W4').value = "NA"#Additional Comment
                                    # xlEnq.range(f'I4').value = "NA"#Additional Comment
                                    xlEnq.range(f'I{current_row}:U{current_row}').api.Validation.Add(Type=7, Formula1=f'=M{current_row}="Yes"')
                                    # xlEnq.range('N{current_row}:Z{current_row}').api.Validation.Add(Type=7, Formula1='=M{current_row}="Yes"')#, ErrorMessage = "Please No to Yes or Other in Quote Column")
                            
                            elif xlEnq.range(f'H{current_row}').value == 'Yes':
                                
                                if state_dict[current_row] != "Yes":
                                    state_dict[current_row] = "Yes"
                                    # for i in range(ord('N'),ord('T')):
                                        # if xlEnq.range(f'{chr(i)}{current_row}').value =="NA":
                                        #     xlEnq.range(f'{chr(i)}{current_row}').value = None
                                    for i in range(ord('I'),ord('T')+1):#N till AC
                                        try:
                                            xlEnq.range(f'{chr(i)}{current_row}').api.Validation.Delete()   
                                        except:
                                            pass
                                    # if xlEnq.range('N{current_row}').value =="NA":
                                    xlEnq.range(f'I{current_row}').value = None
                                    try:
                                        # xlEnq.range('N{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$A$2:$A$5")
                                        xlEnq.range(f'I{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$K$2#")
                                        xlEnq.range(f'I{current_row}').api.Validation.ShowError = False
                                    except:
                                        pass
                                    # if xlEnq.range('O{current_row}').value =="NA":
                                    xlEnq.range(f'J{current_row}').value = None
                                    try:
                                        # xlEnq.range('O{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$B$2:$B${current_row}")
                                        xlEnq.range(f'J{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$M$2#")
                                        xlEnq.range(f'J{current_row}').api.Validation.ShowError = False
                                    except:
                                        pass
                                    # if xlEnq.range('P{current_row}').value =="NA":
                                    xlEnq.range(f'K{current_row}').value = None
                                    try:
                                        # xlEnq.range('P{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$C$2:$C$7")
                                        xlEnq.range(f'K{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$O$2#")
                                        xlEnq.range(f'K{current_row}').api.Validation.ShowError = False
                                        xlEnq.range(f'K{current_row}').api.NumberFormat = "@"
                                    except:
                                        pass
                                    # if xlEnq.range('Q{current_row}').value =="NA":
                                    xlEnq.range(f'L{current_row}').value = None
                                    try:
                                        # xlEnq.range('Q{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$D$2:$D$9")
                                        xlEnq.range(f'L{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$Q$2#")
                                        xlEnq.range(f'L{current_row}').api.Validation.ShowError = False
                                    except:
                                        pass
                                    # if xlEnq.range('R{current_row}').value =="NA":
                                    xlEnq.range(f'M{current_row}').value = None
                                    try:
                                        # xlEnq.range('R{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$E$2:$E$17")
                                        xlEnq.range(f'M{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$S$2#")
                                        xlEnq.range(f'M{current_row}').api.Validation.ShowError = False
                                    except:
                                        pass
                                    # if xlEnq.range('S{current_row}').value =="NA":
                                    xlEnq.range(f'N{current_row}').value = None
                                    try:
                                        # xlEnq.range('S{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$F$2:$F$8")
                                        xlEnq.range(f'N{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$U$2#")
                                        xlEnq.range(f'N{current_row}').api.Validation.ShowError = False
                                    except:
                                        pass
                                    # if xlEnq.range('W{current_row}').value =="NA":
                                    xlEnq.range(f'R{current_row}').value = None
                                    try:
                                        xlEnq.range(f'R{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$J$2:$J$3")
                                        xlEnq.range(f'R{current_row}').api.Validation.ShowError = False
                                    except:
                                        pass
                                    # if xlEnq.range('T{current_row}').value =="NA":
                                    xlEnq.range(f'O{current_row}').value = 0 #LEngth
                                    # if xlEnq.range('U{current_row}').value =="NA":
                                    xlEnq.range(f'P{current_row}').value = 0 #Qty
                                    # if xlEnq.range('V{current_row}').value =="NA":
                                    xlEnq.range(f'Q{current_row}').value = 0 #selling Cost/LBS
                                    # if xlEnq.range('X{current_row}').value =="NA":
                                    xlEnq.range(f'S{current_row}').value = 0 #Selling Cost / UOM
                                    
                                    # if xlEnq.range('Y{current_row}').value =="NA":
                                    xlEnq.range(f'T{current_row}').value = 0 #Additonal Cost
                                    # if xlEnq.range('Z{current_row}').value =="NA":
                                    xlEnq.range(f'U{current_row}').value = 0#Final Priec
                                    # if xlEnq.range('AA{current_row}').value =="NA":
                                    if current_row == 8:
                                        xlEnq.range(f'I3').value = "7 Days" #Validity to be single entry and additonal comment bhi
                                    ####################yaha tak edit kiya hai######################################
                                    
                                else: #function for autofill if current_state is Yes
                                    inv_sht = wb.sheets["Inventory_Lists"]
                                    # map_dict = {"N":"K","O":"M","P":"O","Q":"Q","R":"S","S":"U","W":"I"}
                                    map_dict = {"I":"K","J":"M","K":"O","L":"Q","M":"S","N":"U","R":"I"}
                                    # if wb.app.selection.address[1] in map_dict.keys():
                                    #     wb.app.selection.value = None
                                    # prev_val = {"N":None,"O":None,"P":None,"Q":None,"R":None,"S":None,"W":None}
                                    prev_val = {"I":None,"J":None,"K":None,"L":None,"M":None,"N":None,"R":None}
                                    ######################EDITED TILL HERE
                                    for key in map_dict.keys():
                                        if inv_sht.range(f'{map_dict[key]}2').value != 'Not found':
                                            if inv_sht.range(f'{map_dict[key]}2').value !=None:
                                                # time.sleep(1)
                                                
                                                current_lst = inv_sht.range(f'{map_dict[key]}2').expand('down')
                                            if len(current_lst)==1:
                                                # if prev_val[key] != xlEnq.range(f'{key}{current_row}').value and xlEnq.range(f'{key}{current_row}').value != None:
                                                prev_val[key] == xlEnq.range(f'{key}{current_row}').value
                                                xlEnq.range(f'{key}{current_row}').value = inv_sht.range(f'{map_dict[key]}2').value
                                            elif str(xlEnq.range(f'B{current_row}').value) in current_lst.value and key == "K":#Check if cx Grade in current Eags list after slecting type
                                                prev_val[key] == xlEnq.range(f'{key}{current_row}').value
                                                xlEnq.range(f'{key}{current_row}').value = xlEnq.range(f'B{current_row}').value
                                                yield_key = list(map_dict)[list(map_dict).index(key)+1]
                                                yield_lst = inv_sht.range(f'{map_dict[yield_key]}2').expand('down')
                                                if len(yield_lst)!=1:
                                                    if xlEnq.range(f'C{current_row}').value in yield_lst.value:#Check if yeild exist in current EAGS List then fill yield value
                                                        prev_val[yield_key] == xlEnq.range(f'{yield_key}{current_row}').value
                                                        xlEnq.range(f'{yield_key}{current_row}').value = xlEnq.range(f'C{current_row}').value
                                                        od_key = list(map_dict)[list(map_dict).index(key)+2]
                                                        od_lst = inv_sht.range(f'{map_dict[od_key]}2').expand('down')

                                                        if len(od_lst)!=1:
                                                            if xlEnq.range(f'D{current_row}').value in od_lst.value:#Check if od exist in current EAGS List then fill od value
                                                                prev_val[od_key] == xlEnq.range(f'{od_key}{current_row}').value
                                                                xlEnq.range(f'{od_key}{current_row}').value = xlEnq.range(f'D{current_row}').value
                                                                id_key = list(map_dict)[list(map_dict).index(key)+3]
                                                                id_lst = inv_sht.range(f'{map_dict[id_key]}2').expand('down')
                                                            if len(id_lst)!=1:
                                                                if xlEnq.range(f'E{current_row}').value in id_lst.value:#Check if id exist in current EAGS List then fill od value
                                                                    prev_val[id_key] == xlEnq.range(f'{id_key}{current_row}').value
                                                                    xlEnq.range(f'{id_key}{current_row}').value = xlEnq.range(f'E{current_row}').value
                                                                    id_key = list(map_dict)[list(map_dict).index(key)+3]
                                                                    # id_lst = inv_sht.range(f'{map_dict[id_key]}2').expand('down')

                                        elif inv_sht.range(f'{map_dict[key]}2').value == 'Not found' and xlEnq.range(f'{key}{current_row}').value != None and xlEnq.range(f'{key}{current_row}').value != '':
                                            # root.lift()
                                            # root.overrideredirect(True)
                                            root.attributes("-topmost", True)
                                            #   after 0.5 seconds focus will be stolen
                                            root.after(500, steal_focus)
                                            messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in {key}{current_row}",parent=entry)
                                            # 
                                            
                                            
                                            
                                            
                                            xlEnq.range(f'{key}{current_row}').value = None
                                            wb.activate(steal_focus=True)
                                            xlEnq.range(f'{key}{current_row}').api.Activate()



                            elif xlEnq.range(f'H{current_row}').value == 'Other':
                                if state_dict[current_row] != "Other":
                                    state_dict[current_row] = "Other"
                                    xlEnq.range(f'Q{current_row}').value = 0  
                                    xlEnq.range(f'U{current_row}').value = 0  
                                    xlEnq.range(f'T{current_row}').value = 0  
                                    for i in range(ord('I'),ord('Q')):
                                        xlEnq.range(f'{chr(i)}{current_row}').api.Validation.Delete()
                                        xlEnq.range(f'{chr(i)}{current_row}').value = None
                                        try:
                                            xlEnq.range(f'R{current_row}').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$I$2#")
                                        except:
                                            pass
                                    # xlEnq.range(f'W{current_row}').api.Validation.Delete()
                                    xlEnq.range(f'R{current_row}').value = None
                                    xlEnq.range(f'M{current_row}').value = 0 
                                    xlEnq.range(f'N{current_row}').value = 0 
                                    xlEnq.range(f'O{current_row}').value = 0 
                            #Conditions for Yes/No/Other Drop
                            elif (xlEnq.range(f'H{current_row}').value != None) and (xlEnq.range(f'H{current_row}').value != 'Yes') and (xlEnq.range(f'H{current_row}').value != 'No') and (xlEnq.range(f'H{current_row}').value != 'Other'):
                                cx_db_sht = wb.sheets["Customer_Database"]
                                # m_prev_val = None
                                if cx_db_sht.range(f'P2').value != 'Not found':
                                    current_lst = cx_db_sht.range(f'P2').expand('down')
                                    if len(current_lst)==1:
                                        # if prev_val[key] != xlEnq.range(f'{key}{current_row}').value and xlEnq.range(f'{key}{current_row}').value != None:
                                        # m_prev_val == xlEnq.range(f'{key}{current_row}').value
                                        xlEnq.range(f'H{current_row}').value = cx_db_sht.range(f'P2').value
                                elif cx_db_sht.range(f'P2').value == 'Not found' and xlEnq.range(f'H{current_row}').value != None:
                                    root.lift()
                                    root.overrideredirect(True)
                                    root.attributes("-topmost", True)
                                    root.attributes("-alpha", 0)
                                    #   after 0.5 seconds focus will be stolen
                                    root.after(500, steal_focus)  
                                    messagebox.showerror("Wrong Value Entered", f"Please re-enter correct value in H{current_row}", parent=entry)
                                    xlEnq.range(f'H{current_row}').value = None
                                    wb.activate(steal_focus=True)
                                    xlEnq.range(f'H{current_row}').api.Activate()
                                    ##############################EDITED TILL HERE
                    
                        if xlEnq.range(f'H{current_row}').value != 'No':
                            formula_calculator()
                        numeric_col_checker()
            except Exception as e:
                retrial+=1
                if retrial == 3:
                    a = traceback.format_exc()
                    try:
                        wb.name
                        root.attributes("-topmost", True)
                        #   after 0.5 seconds focus will be stolen
                        root.after(500, steal_focus)  
                        messagebox.showerror("Exception Occured", f"{a}",parent=entry)
                        continue
                    except Exception as e:
                        raise e
                else:
                    continue
                # try:
                #     wb.name
                #     a = traceback.format_exc()
                #     root.attributes("-topmost", True)
                    
                #     messagebox.showerror("Exception Occured", f"{a}")
                #     keepOpen = False
                # except:
                #     keepOpen = False
                        

                

            pythoncom.PumpWaitingMessages()

            # print(xlButtonEvents.myvar)

        
    except Exception as e:
          
        print("Script finished - closing Excel")
        root.attributes("-topmost", True)
        root.focus_force()
        #   after 0.5 seconds focus will be stolen
        root.after(500, steal_focus)  
        messagebox.showinfo("Work Completed", "Excel Closed By User",parent=entry)
        # sys.exit()
        pass
        
    finally:
        try:
            wb.app.kill()
        except:
            pass
    # xw.Close(False) # False: do not save on close, True: save on close
    # wb.app.quit()

if __name__ == '__main__':
    main()
