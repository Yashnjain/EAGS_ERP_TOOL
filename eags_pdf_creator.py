import win32com.client as win32
import pythoncom
import sys, time
import xlwings as xw
from tkinter import messagebox
from tkinter.filedialog import asksaveasfilename
import tkinter as tk
import pandas as pd
from datetime import datetime, date

# use absolute pathes
WORKBOOK = "Sales Form Layout With Database.xlsm"
ENQSHEET = "Enquiry Form"
PDFSHEET = "Final Quote Layout"
# PDFSHEET = "DemoLayout"

root = tk.Tk()

root.withdraw()



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
        
        quote_sht = wb.sheets[PDFSHEET]
        cx_sht = wb.sheets["Customer_Database"]
        df = cx_sht.range("E1").options(pd.DataFrame, 
                             header=1,
                             index=False, 
                             expand='table').value ##df["cus_phone"] = int#
        loc_dict = {"Dubai":"$","Singapore":"$", "USA":"$","UK":"€","NA":""}
        location = xlEnq.range("N4").value
        currency = loc_dict[location]
        phone_dict = df.set_index(['cus_address'])["cus_phone"].to_dict()

        #Basic Details
        #Company(CX Name)
        cx_name = xlEnq.range("C4").value
        #Address
        cx_add = xlEnq.range("D4").value
        #City,Zip
        city_zip = xlEnq.range("N4").value + " ,12345"
        #Phone
        phone = phone_dict[cx_add]
        #Filling Basic Details
        quote_sht.range("B2").value = cx_name
        quote_sht.range("B3").value = cx_add
        quote_sht.range("B4").value = city_zip
        quote_sht.range("B5").value = phone

        #Second Basic Details
        form_date = xlEnq.range("B4").value
        quote_no = int(datetime.strftime(form_date,"%m%d%Y")+"001") #mmddyyyy01
        vaildity = xlEnq.range("AA4").value
        prep_by = xlEnq.range("A4").value
        pay_term = xlEnq.range("E4").value

        #updating Report Sheet
        report_wb = xw.Book("Report.xlsx")
        report_sht = report_wb.sheets["Report"]
        if report_sht.range("A2").value == None:
            report_sht.range("A2").value = quote_no
            report_sht.range("B2").options(transpose = True).value = xlEnq.range("A4:AE4").value
        else:
            latest_row = report_sht.range("A1").end('down').row
            prev_quote_no = int(report_sht.range(f"A{latest_row}").value)
            if str(prev_quote_no)[1:3] == str(quote_no)[1:3]:
                quote_no = prev_quote_no + 1
            report_sht.range(f"A{latest_row+1}").value = quote_no
            report_sht.range(f"B{latest_row+1}").value = xlEnq.range("A4:AE4").value
            report_sht.autofit()
        report_wb.save()
        report_wb.close()

        #Filling 2nd Basic Details
        quote_sht.range("F2").value = quote_no
        quote_sht.range("F3").value = form_date
        quote_sht.range("F4").value = vaildity
        quote_sht.range("F5").value = prep_by
        quote_sht.range("F6").value = pay_term

        #Size Variables
        cx_od = xlEnq.range("I4").value
        cx_id = xlEnq.range("J4").value
        #CX Grade Variables
        cx_grade = xlEnq.range("G4").value
        cx_yield = xlEnq.range("H4").value
        #CX Spec Varaibles
        cx_spec = xlEnq.range("F4").value
        #CX Qty Variables
        cx_qty = xlEnq.range("L4").value
        cx_len = xlEnq.range("K4").value
        #CX Size
        quote_sht.range("B9").value = f'{cx_od}I - {cx_id}I'
        #CX Grade
        quote_sht.range("B10").value = f'{cx_grade}-{cx_yield}'
        #CX Spec
        quote_sht.range("B11").value = cx_spec
        #CX Qty
        quote_sht.range("B12").value = f'{cx_qty} @ {cx_len}"'

        #EAGS details
        if xlEnq.range(f'M4').value == 'No':
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
            #Size Variable
            
            eags_od = xlEnq.range("R4").value
            eags_id = xlEnq.range("S4").value
            
            #CX Grade Variables
            eags_grade = xlEnq.range("P4").value
            eags_yield = xlEnq.range("Q4").value
            #CX Spec Varaibles
            eags_spec = xlEnq.range("F4").value
            #CX Qty Variables
            eags_qty = xlEnq.range("U4").value
            eags_len = xlEnq.range("T4").value
        #EAGS Heat variables
        # eags_heat = xlEnq.range("V4").value
        # #EAGS Tag Variables
        # eags_tag = xlEnq.range("W4").value

        #Filling EAGS Offer Details
        if xlEnq.range(f'M4').value == 'No':
            #EAEAGS Size
            quote_sht.range("B14").value = "NA"
            #EAGS Grade
            quote_sht.range("B15").value = "NA"
            #EAGS Spec
            quote_sht.range("B16").value = "NA"
            #EAGS Qty
            quote_sht.range("B17").value = "NA"
        else:
            #EAEAGS Size
            quote_sht.range("B14").value = f'{eags_od}I-{eags_id}I'
            #EAGS Grade
            quote_sht.range("B15").value = eags_grade+'-'+eags_yield
            #EAGS Spec
            quote_sht.range("B16").value = eags_spec
            #EAGS Qty
            quote_sht.range("B17").value = f'{eags_qty} @ {eags_len}"'
        #EAGS Heat
        # quote_sht.range("B18").value = eags_heat
        # #EAGS TAG
        # quote_sht.range("B18").value = eags_tag

        #Filling other columns
        if xlEnq.range(f'M4').value == 'No':
            #Qty
            quote_sht.range("C13").value = "NA"
            #UOM
            quote_sht.range("D13").value = "NA"
            #Price
            
            price = xlEnq.range("Z4").value
            quote_sht.range("E13").value = "NA"
            #Amount
            amount = "NA"
            quote_sht.range("F13").value = "NA"
            #Price Term
            quote_sht.range("G13").value = "NA"
            quote_sht.range("G14").value = None
        else:
            #Qty
            quote_sht.range("C13").value = f'{eags_qty}PC @ {eags_len}"'
            #UOM
            quote_sht.range("D13").value = xlEnq.range("W4").value
            #Price
            
            price = xlEnq.range("Z4").value
            quote_sht.range("E13").value = f"{price} {currency}"
            #Amount
            amount = float(price)*int(eags_qty)
            quote_sht.range("F13").value = f"{amount} {currency}"
            #Price Term
            quote_sht.range("G13").value = "Ex-works"
            quote_sht.range("G14").value = xlEnq.range("N4").value
        
        

        files = [('pdf', '*.pdf')]
        root.lift()
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0)
        pdf_loc = asksaveasfilename(filetypes = files, defaultextension = files, initialfile=quote_no)
        return pdf_loc

    def formula_calculator():
        try:
            #Calculating Wall thickness(R4-S4)/2 R4=OD, S4=ID
            if xlEnq.range('R4').value != None and xlEnq.range('R4').value != None and xlEnq.range('R4').value != "NA" and xlEnq.range('R4').value != "NA":   
                od = float(xlEnq.range('R4').value)
                id = float(xlEnq.range('S4').value)
                wt = (od-id)/2 #V column earlier
                #=((R4-V4)*V4*10.68)/12
                mid_formula = ((od-wt)*wt*10.68)/12 #W column earlier
                lbs_sell_cost = float(xlEnq.range('V4').value)
                len_col = float(xlEnq.range('T4').value)
                #Selling cost/UOM =X5*W5*T5 SellingCost/LBS * mid_formula * Length
                if xlEnq.range('W4').value == "Each":
                    uom_sell_cost = round((lbs_sell_cost * mid_formula * len_col),2)
                else:
                    uom_sell_cost = round((lbs_sell_cost * mid_formula),2)
                    
                #Filling Calculated values
                #Selling cost/UOM
                xlEnq.range('X4').value = uom_sell_cost
                #Final Price
                xlEnq.range('Z4').formula = "=Y4+X4"
            else:
                #Filling Calculated values
                #Selling cost/UOM
                if xlEnq.range('X4').value != None:
                    xlEnq.range('X4').value = None
                    #Final Price
                    xlEnq.range('Z4').formula = None
        except Exception as e:
            raise e
            

    messagebox.showinfo("Starting Process", "Opening Excel")
    wb = xw.Book(WORKBOOK)
    xlEnq=wb.sheets[ENQSHEET]
    wb.activate()
    xlEnq.activate()
    # define button event callback class
    xlButtonEvents=win32.WithEvents(xlEnq.api.OLEObjects("SubmitButton").Object,ButtonEvents)
    
    eags_col_list = ["N","M"]
    # a while loop to wait until the button in excel is double-clicked
    keepOpen=True
    current_state = "Initial_State"
    while keepOpen:
        try:
            if wb.name == 'Sales Form Layout With Database.xlsm':
                pass
            if xlButtonEvents.buttonClicked is not None:
                
                pdf_loc = quote_generator(wb,xlEnq)
                excel_to_pdf(wb, pdf_loc)
                xlButtonEvents.buttonClicked = None
            else:
                date_row =  xlEnq.range('A2').end('down').row
                cx_row = xlEnq.range('C2').end('down').row
                if xlEnq.range(f'B{date_row}').value == None:
                    xlEnq.range(f'B{date_row}').value = date.today()
                if xlEnq.range(f'D{cx_row}').value == None:
                    xlEnq.range(f'D{cx_row}').formula = "=XLOOKUP(C4,Customer_Database!$C$2:$C$141,Customer_Database!$E$2:$E$141)"
                    xlEnq.range(f'E{cx_row}').formula = "=XLOOKUP(C4,Customer_Database!$C$2:$C$141,Customer_Database!$D$2:$D$141)"
                #Yes/No Condition
                if xlEnq.range(f'M4').value == 'No':
                    if current_state != "No":
                        current_state = "No"
                        for i in range(ord('N'),ord('Z')+1):#N till AC
                            try:
                                xlEnq.range(f'{chr(i)}4').api.Validation.Delete()   
                            except:
                                pass
                            xlEnq.range(f'{chr(i)}4').value = "NA"
                        xlEnq.range(f'AA4').value = "NA"
                        xlEnq.range(f'AB4').value = "NA"
                        xlEnq.range('N4:Z4').api.Validation.Add(Type=7, Formula1='=M4="Yes"')
                        # xlEnq.range('N4:Z4').api.Validation.Add(Type=7, Formula1='=M4="Yes"')#, ErrorMessage = "Please No to Yes or Other in Quote Column")
                
                elif xlEnq.range(f'M4').value == 'Yes':
                    
                    if current_state != "Yes":
                        current_state = "Yes"
                        # for i in range(ord('N'),ord('T')):
                            # if xlEnq.range(f'{chr(i)}4').value =="NA":
                            #     xlEnq.range(f'{chr(i)}4').value = None
                        for i in range(ord('N'),ord('Z')+1):#N till AC
                            try:
                                xlEnq.range(f'{chr(i)}4').api.Validation.Delete()   
                            except:
                                pass
                        # if xlEnq.range('N4').value =="NA":
                        xlEnq.range('N4').value = None
                        try:
                            # xlEnq.range('N4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$A$2:$A$5")
                            xlEnq.range('N4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$K$2#")
                        except:
                            pass
                        # if xlEnq.range('O4').value =="NA":
                        xlEnq.range('O4').value = None
                        try:
                            # xlEnq.range('O4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$B$2:$B$4")
                            xlEnq.range('O4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$M$2#")
                        except:
                            pass
                        # if xlEnq.range('P4').value =="NA":
                        xlEnq.range('P4').value = None
                        try:
                            # xlEnq.range('P4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$C$2:$C$7")
                            xlEnq.range('P4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$O$2#")
                        except:
                            pass
                        # if xlEnq.range('Q4').value =="NA":
                        xlEnq.range('Q4').value = None
                        try:
                            # xlEnq.range('Q4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$D$2:$D$9")
                            xlEnq.range('Q4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$Q$2#")
                        except:
                            pass
                        # if xlEnq.range('R4').value =="NA":
                        xlEnq.range('R4').value = None
                        try:
                            # xlEnq.range('R4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$E$2:$E$17")
                            xlEnq.range('R4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$S$2#")
                        except:
                            pass
                        # if xlEnq.range('S4').value =="NA":
                        xlEnq.range('S4').value = None
                        try:
                            # xlEnq.range('S4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$F$2:$F$8")
                            xlEnq.range('S4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$U$2#")
                        except:
                            pass
                        # if xlEnq.range('W4').value =="NA":
                        xlEnq.range('W4').value = None
                        try:
                            xlEnq.range('W4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$J$2:$J$3")
                        except:
                            pass
                        # if xlEnq.range('T4').value =="NA":
                        xlEnq.range('T4').value = 0
                        # if xlEnq.range('U4').value =="NA":
                        xlEnq.range('U4').value = 0
                        # if xlEnq.range('V4').value =="NA":
                        xlEnq.range('V4').value = 0
                        # if xlEnq.range('X4').value =="NA":
                        xlEnq.range('X4').value = 0
                        
                        # if xlEnq.range('Y4').value =="NA":
                        xlEnq.range('Y4').value = 0
                        # if xlEnq.range('Z4').value =="NA":
                        xlEnq.range('Z4').value = 0
                        # if xlEnq.range('AA4').value =="NA":
                        xlEnq.range('AA4').value = "NA"
                        
                    else:
                        inv_sht = wb.sheets["Inventory_Lists"]
                        map_dict = {"N":"K","O":"M","P":"O","Q":"Q","R":"S","S":"U"}
                        # if wb.app.selection.address[1] in map_dict.keys():
                        #     wb.app.selection.value = None
                        prev_val = {"N":None,"O":None,"P":None,"Q":None,"R":None,"S":None}
                        for key in map_dict.keys():
                            if inv_sht.range(f'{map_dict[key]}2').value != 'Not found':
                                current_lst = inv_sht.range(f'{map_dict[key]}2').expand('down')
                                if len(current_lst)==1:
                                    # if prev_val[key] != xlEnq.range(f'{key}4').value and xlEnq.range(f'{key}4').value != None:
                                    prev_val[key] == xlEnq.range(f'{key}4').value
                                    xlEnq.range(f'{key}4').value = inv_sht.range(f'{map_dict[key]}2').value


                elif xlEnq.range(f'M4').value == 'Other':
                    if current_state != "Other":
                        current_state = "Other"
                        xlEnq.range('V4').value = 0  
                        xlEnq.range('Z4').value = 0  
                        xlEnq.range('Y4').value = 0  
                        for i in range(ord('N'),ord('V')):
                            xlEnq.range(f'{chr(i)}4').api.Validation.Delete()
                            xlEnq.range(f'{chr(i)}4').value = None
                            try:
                                xlEnq.range('W4').api.Validation.Add(Type=3, Formula1="=Inventory_Lists!$J$2:$J$3")
                            except:
                                pass
                        # xlEnq.range(f'W4').api.Validation.Delete()
                        xlEnq.range(f'W4').value = None
                        xlEnq.range('R4').value = 0 
                        xlEnq.range('S4').value = 0 
                        xlEnq.range('T4').value = 0 
                if xlEnq.range(f'M4').value != 'No':
                    formula_calculator()
        except:
            keepOpen = False
            

        pythoncom.PumpWaitingMessages()

        # print(xlButtonEvents.myvar)

    print("Script finished - closing Excel")
    messagebox.showinfo("Work Completed", "Excel Closed By User")
    sys.exit()
    # xw.Close(False) # False: do not save on close, True: save on close
    # wb.app.quit()

if __name__ == '__main__':
    main()
