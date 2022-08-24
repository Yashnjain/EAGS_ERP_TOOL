import xlwings as xw
import xlwings.constants as win32c
import time
import os
import pandas as pd
from RTools import resource_path

def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        raise e

def insert_rows(cell_values,working_sheet,working_workbook):
    try:
        last_row=working_sheet.range(f'A'+ str(working_sheet.cells.last_cell.row)).end('up').end('up').row
        #logic for multiple insertions of quotation rows for pdf creation
        working_sheet.api.Range(cell_values).Select()
        working_workbook.app.api.Selection.Copy()
        working_sheet.api.Range(f"A{last_row}").Select()
        working_workbook.app.api.Selection.Insert(Shift=win32c.InsertShiftDirection.xlShiftDown)
    except Exception as e:
        raise e

def chunks(lst, n):
    try:
        """Yield successive n-sized chunks from lst."""
        for i in range(0, len(lst), n):
            yield lst[i:i + n]
    except Exception as e:
                raise e

def pdf_generator(df):
    try:       
        # job_name = 'open_ar_v2' 
        input_sheet = resource_path('pdfCreator.xlsm')
        # input_sheet='pdfCreator.xlsm'
        retry=0
        while retry < 10:
            try:
                app = xw.App(visible=False)
                wb = app.books.open(input_sheet) 
                wb.app.visible = False
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        ws1=wb.sheets[1]  
        ws1.activate() 
                
        list1=list(chunks(range(0, (len(df))), 8))
        page_count=0
        for index,values in enumerate(list1):
                chunk_df=df.iloc[values]
                chunk_df.reset_index(inplace=True)
                chunk_df = chunk_df.astype({'C_OD':'float','C_ID':'float'})
                page_diff=70
                ws1.range(f'F{2+(page_count*page_diff)}').value=df['QUOTENO'][0]
                ws1.range(f'F{3+(page_count*page_diff)}').value=df['DATE'][0]
                ws1.range(f'F{4+(page_count*page_diff)}').value=df['VALIDITY'][0]
                ws1.range(f'F{5+(page_count*page_diff)}').value=df['PREPAREDBY'][0]
                ws1.range(f'F{6+(page_count*page_diff)}').value=df['PAYMENT_TERM'][0]
                
                ws1.range(f'B{2+(page_count*page_diff)}').value=df['CUS_NAME'][0]
                ws1.range(f'B{3+(page_count*page_diff)}').value=df['CUS_ADDRESS'][0]
                ws1.range(f'B{4+(page_count*page_diff)}').value=df['CUS_CITY_ZIP'][0]  #city/zipcode
                ws1.range(f'B{5+(page_count*page_diff)}').value=df['CUS_PHONE'][0]  #Phone
                ws1.range(f'B{6+(page_count*page_diff)}').value=df['CUS_EMAIL'][0]  #email
                ws1.range(f'A{18+(page_count*page_diff)}').value=df['ADD_COMMENTS'][0]
                for i in range(0,(len(chunk_df))):
                    diff=7
                    ws1.range(f'A{8+(diff*i+page_count*page_diff)}').value="Customer Requirement:"
                    customer_requirement=f'{round(float(chunk_df["C_OD"][i]),3)}"OD - {round(chunk_df["C_ID"][i],3)}"ID - {chunk_df["C_GRADE"][i]} - {chunk_df["C_YIELD"][i]} - {chunk_df["C_SPECIFICATION"][i]} - {chunk_df["C_QTY"][i]}@{chunk_df["C_LENGTH"][i]}"'
                    ws1.range(f'A{9+(diff*i+page_count*page_diff)}').value=customer_requirement
                    ws1.range(f'A{10+(diff*i+page_count*page_diff)}').value="EAGL Offer:"
                    
                    ws1.range(f'D{10+(diff*i+page_count*page_diff)}').value=chunk_df['E_UOM'][i]#UOM    [$$-en-US]#,##0.00   [$£-en-GB]#,##0.00
                    ws1.range(f'E{10+(diff*i+page_count*page_diff)}').value=float(chunk_df['E_FINAL_PRICE'][i])#Price
                    if chunk_df['CURRENCY'][i] == "$":
                        ws1.range(f'E{10+(diff*i+page_count*page_diff)}').api.NumberFormat = "[$$-en-US]#,##0.00"
                    elif chunk_df['CURRENCY'][i] == "£":
                        ws1.range(f'E{10+(diff*i+page_count*page_diff)}').api.NumberFormat = "[$£-en-GB]#,##0.00"
                    else:
                        pass
                    if chunk_df['E_FINAL_PRICE'][i] != "NA":
                        ws1.range(f'F{10+(diff*i+page_count*page_diff)}').value=float(chunk_df['E_FINAL_PRICE'][i])*float(chunk_df['E_QTY'][i])#amount
                        if chunk_df['CURRENCY'][i] == "$":
                            ws1.range(f'F{10+(diff*i+page_count*page_diff)}').api.NumberFormat = "[$$-en-US]#,##0.00"
                        elif chunk_df['CURRENCY'][i] == "£":
                            ws1.range(f'F{10+(diff*i+page_count*page_diff)}').api.NumberFormat = "[$£-en-GB]#,##0.00"
                        else:
                            pass    
                        ws1.range(f'H{10+(diff*i+page_count*page_diff)}').value=f"Ex-works"#delivery term
                        ws1.range(f'C{10+(diff*i+page_count*page_diff)}').value=f'{chunk_df["E_QTY"][i]}PC@{chunk_df["E_LENGTH"][i]}"'#QTY
                        ws1.range(f'H{11+(diff*i+page_count*page_diff)}').value=f"{chunk_df['E_LOCATION'][i]}"#delivery term
                    else:
                        ws1.range(f'F{10+(diff*i+page_count*page_diff)}').value = "NA"
                        ws1.range(f'H{10+(diff*i+page_count*page_diff)}').value="NA"#delivery term
                        ws1.range(f'C{10+(diff*i+page_count*page_diff)}').value="NA"#QTY
                    ws1.range(f'G{10+(diff*i+page_count*page_diff)}').value=chunk_df['LEAD_TIME'][i]#lead time                    
                    
                    if (chunk_df["C_QUOTE_YES/NO"][i]).lower()=='yes':
                        # chunk_df = chunk_df.astype({'E_OD1':'float','E_ID1':'float', 'E_FINAL_PRICE':'float', 'E_QTY':'float'})
                        eagl_offer=f'{round(float(chunk_df["E_OD1"][i]),3)}"OD - {round(float(chunk_df["E_ID1"][i]),3)}"ID - {chunk_df["E_GRADE"][i]} - {chunk_df["E_YIELD"][i]} - {chunk_df["C_SPECIFICATION"][i]} - {float(chunk_df["E_QTY"][i])}@{float(chunk_df["E_LENGTH"][i])}"'
                        ws1.range(f'A{11+(diff*i+page_count*page_diff)}').value= eagl_offer
                    if (chunk_df["C_QUOTE_YES/NO"][i]).lower()=='no' or (chunk_df["C_QUOTE_YES/NO"][i]).lower()=='other':
                        eagl_offer='NA'#f'{chunk_df["E_OD1"][i]}"OD - {chunk_df["E_ID1"][i]}"ID - {chunk_df["E_GRADE"][i]} - {chunk_df["E_YIELD"][i]} - {chunk_df["E_SPEC"][i]} - {chunk_df["E_QTY"][i]}@{chunk_df["E_LENGTH"][i]}"'
                        ws1.range(f'A{11+(diff*i+page_count*page_diff)}').value= eagl_offer
                    if i!=(len(chunk_df))-1:
                           last_row=8+(diff*(i+1)+page_count*page_diff) 
                           cell_values= f'A{8+(diff*i)+page_count*page_diff}:H{last_row-1}'                 
                           insert_rows(cell_values,ws1,wb)  
                if index!=(len(list1))-1: 
                    page_count+=1     
                    ws2=wb.sheets[0]  
                    ws2.activate() 
                    last_row = ws2.range(f'A'+ str(ws2.cells.last_cell.row)).end('up').row
                    last_column_letter=num_to_col_letters(ws2.range('A7').end('right').end('right').column)
                    ws2.range(f"A1:{last_column_letter}{last_row}").copy(ws1.range(f"A{page_count*page_diff+1}"))    
                    ws1.activate()
                   
        wb.api.ActiveSheet.PageSetup.Zoom=70
        current_work_dir = os.getcwd()
        filename=str(df['QUOTENO'][0]).replace("/","")
        pdf_path = os.path.join(current_work_dir, f"{filename}.pdf")
        # Save excel workbook to pdf file
        # print(f"Saving workbook as '{pdf_path}' ...")
        ws1.api.ExportAsFixedFormat(0, pdf_path)
        # Open the created pdf file
        # print(f"Opening pdf file with default application ...")
        # os.startfile(pdf_path)
        # return f"pdf generated succesfully"
        return pdf_path

    except Exception as e:
        raise e
    finally:
        try:
            wb.app.kill()
        except:
            pass
# df=pd.read_csv("result.csv")
# msg = pdf_generator(df)
# print(msg)
