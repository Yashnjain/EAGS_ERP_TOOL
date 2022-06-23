import pandas as pd
from datetime import datetime, date
from sfTool import eagsQuotationuploader,getLatestQuote
def dfMaker(specialList,cxList,otherList,pt,conn):

    # colList = list(inpDict.keys())
    # for col in colList:
    #     for i in range(len(inpDict[col][0])):
    #         inpDict[col][0][i] = inpDict[col][0][i][0].get()
    # EAGS/Location/Year/Number
    # Location: USA/UK/DUB/SGP
    # Year: 2022
    # Number:00xxxx
    locDict = {"DUBAI":"DUB", "SINGAPORE":"SGP", "USA":"USA","UK":"UK"}
    locVar = specialList['E_Location'][0][0][0].get()
    location = locDict[locVar.upper()]
    cxList[1]=datetime.strptime(cxList[1],"%m.%d.%Y").date()
    input_year=datetime.strftime(cxList[1],"%Y")
    
    #try to get latest quote number of same combination if not present then put 1 otherwise increament current contract
    curr_quoteNo = f"EAGS/{location}/{input_year}/000001"

    new_quoteNo = getLatestQuote(conn,curr_quoteNo)
    



    otherList.append(date.today())



    columnList = ['QUOTENO', 'PREPAREDBY', 'DATE', 'CUS_NAME', 'PAYMENT_TERM', 'CUS_ADDRESS', 'CUS_PHONE', 'CUS_EMAIL', 'CUS_CITY_ZIP', 'C_SPECIFICATION', 
    'C_GRADE', 'C_YIELD', 'C_OD', 'C_ID', 'C_LENGTH', 'C_QTY', 'C_QUOTE_YES/NO', 'E_LOCATION', 'E_TYPE', 'E_GRADE', 'E_YIELD', 'E_OD', 'E_ID', 'E_LENGTH',
     'E_QTY', 'E_SELLING_COST/LBS', 'E_UOM', 'E_SELLING_COST/UOM', 'E_ADDITIONAL_COST', 'LEAD_TIME','E_FINAL_PRICE', 'VALIDITY', 'ADD_COMMENTS','INSERT_DATE']
    row = []
    
    colList = list(specialList.keys())
    for i in range(len(specialList[colList[0]][0])):
        rowList = []
        rowList.append(new_quoteNo)
        rowList.extend(cxList)
        for col in colList:
            rowList.append(specialList[col][0][i][0].get()) #Insert jth column with ith index in rowList
        rowList.extend(otherList)#insert validity and additional comments
        row.append(rowList)#Append current ith row to row List
         #Empty rowList for fetching next row
    # print(row)
    print(len(columnList))
    print(len(row[0]))

    if specialList["E_Final Price"][0][0][0].get() != "":
        sfDf = pd.DataFrame(row, columns=columnList)
        print(sfDf)
        pt.model.df = sfDf
        pt.redraw()
        eagsQuotationuploader(conn, sfDf)
        
    print()


    # row = [[trade_date, flow_m1, ch_opis, ny_opis, flow[flow_m1][2], flow[flow_m1][3], flow[flow_m1][1],flow[flow_m1][0], in_date, up_date]]
    #     for i in f_keys[1:len(f_keys)]:
    #         row.append([trade_date, i, np.nan, np.nan, flow[i][2], flow[i][3], flow[i][1], flow[i][0], in_date, up_date])

    #     df = pd.DataFrame(row, columns=['TRADEDATE', 'FLOWMONTH', 'OPIS_CH', 'OPIS_NY', 'NYMEX_CH', 'NYMEX_NY', 'NYMEX_RBOB', 'CORN',
    #                                     'INSERT_DATE', 'UPDATE_DATE'], index=None)




# 'QuoteNo, PreparedBy, Date, Cus_Name, Payment_Term, Cus_Address, Cus_Phone, Cus_Email, Cus_city_zip, C_Specification, C_Grade, C_Yield, C_OD, C_ID, C_Length, C_Qty, C_Quote Yes/No, E_Location, E_Type, E_Grade, E_Yield, E_OD, E_ID, E_Length, E_Qty, E_Selling Cost/LBS, E_UOM, E_Selling Cost / UOM, E_Additional Cost, E_Final Price'
'quoteno, preparedby, date, cus_name, payment_term, cus_address, cus_phone, cus_email, cus_city_zip, c_specification, c_grade, c_yield, c_od, c_id, c_length, c_qty, c_quote yes/no, e_location, e_type, e_grade, e_yield, e_od, e_id, e_length, e_qty, e_selling cost/lbs, e_uom, e_selling cost / uom, e_additional cost, e_final price'
'QUOTENO, PREPAREDBY, DATE, CUS_NAME, PAYMENT_TERM, CUS_ADDRESS, CUS_PHONE, CUS_EMAIL, CUS_CITY_ZIP, C_SPECIFICATION, C_GRADE, C_YIELD, C_OD, C_ID, C_LENGTH, C_QTY, C_QUOTE YES/NO, E_LOCATION, E_TYPE, E_GRADE, E_YIELD, E_OD, E_ID, E_LENGTH, E_QTY, E_SELLING COST/LBS, E_UOM, E_SELLING COST / UOM, E_ADDITIONAL COST, E_FINAL PRICE, VALIDITY, ADD_COMMENTS'