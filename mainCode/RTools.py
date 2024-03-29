from shutil import ExecError
import pandas as pd
from datetime import datetime, date
from RsfTool import eagsQuotationuploader,getLatestQuote
import os, sys
import tkinter as tk
from tkinter import ttk
import customComboboxV2 #import myCombobox
from tkinter import messagebox





def resource_path(relative_path):
    try:
        """ Get absolute path to resource, works for dev and for PyInstaller """
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    except Exception as e:
                raise e







def dfMaker(specialList,cxList,otherList,pt,conn,quote_number, root):
    try:
        # colList = list(inpDict.keys())
        # for col in colList:
        #     for i in range(len(inpDict[col][0])):
        #         inpDict[col][0][i] = inpDict[col][0][i][0].get()
        # EAGS/Location/Year/Number
        # Location: USA/UK/DUB/SGP
        # Year: 2022
        # Number:00xxxx
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
            previous_quote_number=quote_number
            new_quoteNo,raw_data = getLatestQuote(conn,curr_quoteNo,previous_quote_number)

            otherList.append(quote_number)
            otherList.append(1)
            otherList.append(date.today())



            # columnList = ['QUOTENO', 'PREPAREDBY', 'DATE', 'CUS_NAME', 'PAYMENT_TERM', 'CURRENCY','CUS_ADDRESS', 'CUS_PHONE', 'CUS_EMAIL', 'CUS_CITY_ZIP', 'C_SPECIFICATION', 'C_TYPE',
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
            return sfDf,raw_data
        else:
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

def bakerMaker(specialList,cxList,otherList,ptBaker,conn, root):
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



            # columnList = ['QUOTENO', 'PREPAREDBY', 'DATE', 'CUS_NAME', 'PAYMENT_TERM', 'CURRENCY', 'CUS_ADDRESS', 'CUS_PHONE', 'CUS_EMAIL', 'CUS_CITY_ZIP', 'C_SPECIFICATION', 'C_TYPE',
            # 'C_GRADE', 'C_YIELD', 'C_OD', 'C_ID', 'C_LENGTH', 'C_QTY', 'C_QUOTE_YES/NO', 'E_LOCATION', 'E_TYPE', 'E_SPEC','E_GRADE', 'E_YIELD', 'E_OD1', 'E_ID1', 'E_OD2', 'E_ID2', 'E_LENGTH',
            # 'E_QTY', 'E_SELLING_COST/LBS', 'E_UOM', 'E_SELLING_COST/UOM', 'E_ADDITIONAL_COST', 'LEAD_TIME','E_FINAL_PRICE', 'VALIDITY', 'ADD_COMMENTS','PREVIOUS_QUOTE','REV_CHECKER', 'INSERT_DATE']

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
            # xlList = ["C_Specification","C_Type","C_Grade","C_Yield", "C_OD", "C_ID", "C_Length", "C_Qty", 'E_freightIncured', 'E_freightCharged','E_Margin_Freight', 'Lot_Serial_Number']
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
                bakerxlDf['DLV_DATE'] = bakerxlDf['DLV_DATE'].astype('datetime64[ns]').astype(str)
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
                        # print(specialList[col][0][i][0].get())
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



def specialCase(root, specialList,index,pt,df,row_num,quotedf):
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
        od1.insert(tk.END,quotedf['E_OD1'][row_num]) 

        id1Var = tk.StringVar()
        id1 = ttk.Entry(entryFrame1, textvariable=id1Var, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        id1['validatecommand'] = (id1.register(intFloat),'%P','%d')
        id1.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        id1.insert(tk.END,quotedf['E_ID1'][row_num])

        e_od2Var = tk.StringVar()
        global e_od2
        e_od2 = ttk.Entry(entryFrame2, textvariable=e_od2Var, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        e_od2['validatecommand'] = (e_od2.register(intFloat),'%P','%d')
        e_od2.grid(row=1,column=0,sticky=tk.EW,padx=5,pady=5)
        if quotedf['E_OD2'][row_num]:
            e_od2.insert(tk.END,quotedf['E_OD2'][row_num])
        # boxList["E_OD2"][0][index] = customComboboxV2.myCombobox(df,toproot,frame=entryFrame2,row=1,column=0,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = boxList,pt=pt,item_list=item_list)
        # boxList["E_OD2"][0][-1][0]['validate']='key'
        # boxList["E_OD2"][0][-1][0]['validatecommand'] = (boxList["E_OD2"][0][-1][0].register(intFloat),'%P','%d')
        # e_od[-1].config(textvariable="NA", state='disabled')
        # e_od2[-1][0]['validate']='key'
        # e_od2[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')

        e_id2Var = tk.StringVar()
        global e_id2
        e_id2 = ttk.Entry(entryFrame2, textvariable=e_id2Var, background = 'white',width = 10,validate = "key",
                validatecommand=(vcmd, '%P','%d'))
        e_id2['validatecommand'] = (e_id2.register(intFloat),'%P','%d')
        e_id2.grid(row=1,column=1,sticky=tk.EW,padx=5,pady=5)
        if quotedf['E_ID2'][row_num]:
           e_id2.insert(tk.END,quotedf['E_ID2'][row_num])

        # boxList["E_ID2"][0][index] = customComboboxV2.myCombobox(df,toproot,frame=entryFrame2,row=1,column=1,width=10,list_bd = 0,foreground='blue', background='white',sticky = "nsew",boxList = boxList,pt=pt,item_list=item_list)
        # boxList["E_ID2"][0][-1][0]['validate']='key'
        # boxList["E_ID2"][0][-1][0]['validatecommand'] = (boxList["E_ID2"][0][-1][0].register(intFloat),'%P','%d')
        # e_id1[-1][0]['validate']='key'
        # e_id1[-1][0]['validatecommand'] = (e_od1[-1][0].register(intFloat),'%P','%d')

        # od2Var = tk.StringVar()
        # od2 = ttk.Entry(entryFrame2, textvariable=od2Var, background = 'white',width = 15)
        # od2.grid(row=1,column=0)

        # id2Var = tk.StringVar()
        # id2 = ttk.Entry(entryFrame2, textvariable=id2Var, background = 'white',width = 15)
        # id2.grid(row=1,column=1)
        def exitTrue(close_check=False):
            od1=od1Var.get()
            id1 = id1Var.get()
            if not close_check and ((specialList["E_OD2"][0][index][1].get() == '') or (specialList["E_ID2"][0][index][1].get() == '') or od1 == '' or id1 == ''):
                root.attributes('-topmost', True)
                messagebox.showerror(title="Value Error",message="Please fill all values first",parent=root)
                root.attributes('-topmost', False)
                return
            elif close_check and ((specialList["E_OD2"][0][index][1].get() == '') or (specialList["E_ID2"][0][index][1].get() == '') or od1 == '' or id1 == ''):
                    toproot.attributes('-topmost', True)
                    messagebox.showerror(title="Closing with Blank",message="Please select type again as currently no data was provided",parent=toproot)
                    toproot.attributes('-topmost', False)
                    specialList['E_Type'][0][index][1].set("")
                    specialList['E_Grade'][0][index][1].set("")
                    specialList['E_Yield'][0][index][1].set("")
                    toproot.attributes('-topmost', False)
                    toproot.grab_release()
                    toproot.destroy()
            else:
                specialList["E_OD2"][0][index] = (e_od2.get(),e_od2.get())
                specialList["E_ID2"][0][index] = (e_id2.get(),e_id2.get())
                specialList['E_OD1'][0][index][0].configure(state='disabled')
                specialList['E_ID1'][0][index][0].configure(state='disabled')
                
                toproot.attributes('-topmost', False)
                toproot.grab_release()
                toproot.destroy()
                specialList['E_OD1'][0][index][1].set(float(od1))
                specialList['E_ID1'][0][index][1].set(float(id1))
                newDf = df[(df["site"] == specialList['E_Location'][0][index][0].get())
                                                    & (df["grade"]==specialList['E_Grade'][0][index][0].get())& (df["heat_condition"]==specialList['E_Yield'][0][index][0].get())
                                                    & (df["od_in"]==float(specialList['E_OD2'][0][index][0])) & (df["od_in_2"]==float(specialList['E_ID2'][0][index][0]))]
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
                specialList['E_Length'][0][index][0].focus()
                # toproot.destroy()
                check=True
        
        def on_closing():
            try:
                toproot.attributes('-topmost', True)
                if messagebox.askokcancel("Quit", "Do you want to quit?",parent=toproot):
                    toproot.attributes('-topmost', False)
                    close_check=True
                    exitTrue(close_check)
                root.attributes('-topmost', False)
            except Exception as e:
                raise e
        submitButton = tk.Button(submitFrame,text="Submit", command=exitTrue)
        # submitButton.place(relx=.5, rely=.5, anchor="center")
        submitButton.grid(row=0,column=1,pady=40)
        toproot.focus()
        od1.focus()
        toproot.protocol("WM_DELETE_WINDOW", on_closing)
        toproot.mainloop()
        if check:
            toproot.attributes('-topmost', False)
            toproot.grab_release()
            toproot.destroy()
            return od1.get(), id1.get(), specialList["E_OD2"][index][0].get(), specialList["E_ID2"][index][0].get()
        else:
            return None, None, None, None
    except Exception as e:
        raise e

    