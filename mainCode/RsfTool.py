from errno import ESTALE
import snowflake.connector
from snowflake.connector.pandas_tools import write_pandas
# from snowflake.sqlalchemy import URL
# from sqlalchemy import create_engine
import pandas as pd
# import datetime

# USER = "SVC_BUITDB_DEV"
# PASSWORD = "BUITDBDEV2022"
# ACCOUNT = 'OS54042.east-us-2.azure'
# WAREHOUSE = "BUIT_WH"
# # DATABASE = "BUITDB"
# DATABASE = "BUITDB_DEV"
# SCHEMA = "EAGS"
# ROLE = "OWNER_BUITDB_DEV"
# # ROLE = "OWNER_BUITDB"
USER =  "SVC_BUIT"#  "SVC_BUITDB_DEV"
PASSWORD = "BUIT2022"
ACCOUNT = 'OS54042.east-us-2.azure'
WAREHOUSE = "BUIT_WH"
DATABASE = "BUITDB"
# DATABASE = "BUITDB_DEV"
SCHEMA = "EAGS"
# ROLE = "OWNER_BUITDB_DEV"
ROLE = "OWNER_BUITDB"

EAGS_QUOTATION_TABLE = "EAGS_QUOTATION"
EAGS_BAKER_TABLE = "EAGS_BAKER"

def get_connection():
    try:
        # engine = bu_snowflake.get_engine(
        #         username=USER,
        #         password=PASSWORD,
        #         warehouse=WAREHOUSE,
        #         role=ROLE,
        #         schema=SCHEMA,
        #         database=DATABASE
        #         )
        # conn = engine.connect()
        conn=snowflake.connector.connect(

                user=USER,
                password=PASSWORD,
                account=ACCOUNT,
                warehouse=WAREHOUSE,
                database  = DATABASE,
                schema=SCHEMA,
                role =ROLE
            ) 
        
        return conn
    except Exception as e:
        raise e

def loginChecker(conn,table, user, pwd):
    try:
 
        # start_time = time.time()
        # engine = bu_snowflake.get_engine(
        #     username=USER,
        #     password=PASSWORD,
        #     warehouse=WAREHOUSE,
        #     role=ROLE,
        #     schema=SCHEMA,
        #     database=DATABASE
        #     )
        # conn = engine.connect()
        # end_time = time.time()
        # engineTime = end_time - start_time
        # print(engineTime)
        # engine.dispose()
        # start_time = time.time()
        # cnn=snowflake.connector.connect(

        #     user=USER,
        #     password=PASSWORD,
        #     account=ACCOUNT,
        #     warehouse=WAREHOUSE,
        #     database  = DATABASE,
        #     schema=SCHEMA,
        #     role =ROLE
        # ) 
        # end_time = time.time()
        # sfTime = end_time - start_time
        # print(sfTime)
        # cnn.close()

        # conn = cnn.cursor()
        # conn = engine.connect()
        query = f"SELECT NAME FROM {DATABASE}.{SCHEMA}.{table} WHERE USERNAME = '{user}' AND PASSWORD = '{pwd}'"
        
        # raw_data = conn.execute(f"SELECT NAME FROM {DATABASE}.{SCHEMA}.{table} WHERE USERNAME = '{user}' AND PASSWORD = '{pwd}'")
        cur = conn.cursor()
        cur.execute(query)
        # data = cur.fetch_pandas_all()
        data = cur.fetchall()
        
        # data = raw_data.fetchall()
        if len(data):
            return data[0][0]#data.iloc[0,0]
        return False
    except Exception as e:
        raise e
    finally:
        pass
        # engine.dispose()
        # conn.close()

# loginChecker("table","user", "pwd")

def get_cx_df(conn,table, customer= 'general'):
    try:
        # query = f"SELECT CUS_LONG_NAME, PAYMENT_TERM, CUS_ADDRESS, CUS_PHONE, CUS_EMAIL, CUS_CITY_ZIP FROM {DATABASE}.{SCHEMA}.{table}"
        if customer == 'baker':
            query = f"SELECT * FROM {DATABASE}.{SCHEMA}.{table} WHERE COMPANY='BAKER_HUGES'"
        else:
            query = f"SELECT * FROM {DATABASE}.{SCHEMA}.{table}"
        cur = conn.cursor()
        cur.execute(query)
        names = [ x[0] for x in cur.description]
        rows = cur.fetchall()
        df= pd.DataFrame( rows, columns=names)
        # df = cur.fetch_pandas_all()
        df.columns = map(str.lower  , df.columns)
        # df = pd.read_sql_query(query, conn)
        return df
    except Exception as e:
        raise e

def get_inv_df(conn, table):
    try:
        # init_time = datetime.datetime.now()
        # print(init_time)
        query = f"SELECT SITE, MATERIAL_TYPE, GLOBAL_GRADE, OD_IN, OD_IN_2, HEAT_CONDITION, ONHAND_PIECES, ONHAND_LENGTH_IN, ONHAND_DOLLARS_PER_POUNDS, RESERVED_PIECES, RESERVED_LENGTH_IN, AVAILABLE_PIECES, AVAILABLE_LENGTH_IN, DATE_LAST_RECEIPT, AGE FROM {DATABASE}.{SCHEMA}.{table}"
        cur = conn.cursor()
        cur.execute(query)
        names = [ x[0] for x in cur.description]
        rows = cur.fetchall()
        df= pd.DataFrame( rows, columns=names)
        # df = cur.fetch_pandas_all()
        df.columns = map(str.lower  , df.columns)
        df = df.fillna("")
        df['od_in'] = df['od_in'].astype('float')
        df['od_in_2'] = df['od_in_2'].astype('float')
        # df = pd.read_sql_query(query, conn)#conn.cursor.fetch_pandas
        # print(datetime.datetime.now())
        # timeTaken = datetime.datetime.now() - init_time
        # print(f"current time taken for fetch_all = {timeTaken}")
        
        # init_time = datetime.datetime.now()
        # print(init_time)
        # cur.execute(query)
        # df = cur.fetch_pandas_all()
        # timeTaken = datetime.datetime.now() - init_time
        # print(f"current time taken for pandas all = {timeTaken}")
        
        # init_time = datetime.datetime.now()
        # cur.execute(query)
        # for df in cur.fetch_pandas_batches():
        #     print(df)

        # # df = cur.fetch_pandas_all()
        # timeTaken = datetime.datetime.now() - init_time
        # print(f"current time taken for pandas all = {timeTaken}")

        return df
    except Exception as e:
        raise e


def get_qtylengthdf(conn, table, location, type, grade, od, id):
    try:
        query = f"""SELECT ONHAND_PIECES, ONHAND_LENGTH_IN, RESERVED_PIECES, RESERVED_LENGTH_IN, AVAILABLE_PIECES, AVAILABLE_LENGTH_IN FROM {DATABASE}.{SCHEMA}.{table} 
                    WHERE SITE={location} AND MATERIAL_TYPE={type} AND GLOBAL_GRADE={grade}  AND OD_IN = {od} AND OD_IN_2 = {id}"""
        cur = conn.cursor()
        
        cur.execute(query)
        names = [ x[0] for x in cur.description]
        rows = cur.fetchall()
        df= pd.DataFrame( rows, columns=names)
        # df = cur.fetch_pandas_all()
        df.columns = map(str.lower, df.columns)
        # df = pd.read_sql_query(query, conn)
        return df
    except Exception as e:
        raise e

#EAGS_Quotation   
def eagsQuotationuploader(conn,df,latest_revised_quote=None,baker=False):
    try:
                 #---------Updating previous rev checker to 0 ----------#
        if latest_revised_quote is not None:
            rev_query = f'''UPDATE {DATABASE}.{SCHEMA}.{EAGS_QUOTATION_TABLE} set REV_CHECKER = 0 Where QUOTENO = '{latest_revised_quote}';'''
            cur = conn.cursor()
            cur.execute(rev_query)
        # Write the data from the DataFrame to the table named "customers".
        # Write the data from the DataFrame to the table named "eags quotation".
        if baker:
            df.columns = map(str.upper, df.columns)
            df. columns = df. columns. str. replace(' ','_')
            success, nchunks, nrows, _ = write_pandas(conn, df, EAGS_BAKER_TABLE)
        else:
            success, nchunks, nrows, _ = write_pandas(conn, df, EAGS_QUOTATION_TABLE)
    except Exception as e:
        raise e

    # df.to_sql(name=EAGS_QUOTATION_TABLE, con=conn, index=False,if_exists='append', schema=SCHEMA)


def getLatestQuote(conn,curr_quoteNo,previous_quote_number=None, baker=False, newQuote=False):
    try:
        cx_init_name = curr_quoteNo.split("_")[0]
        if previous_quote_number is None:
            previous_quote_number = cx_init_name

        if previous_quote_number.rfind("R")!=-1:
            cx_init_name_previous=previous_quote_number[:previous_quote_number.rfind("R")]
        else:
             cx_init_name_previous=previous_quote_number   
        # query = f"SELECT QUOTENO FROM {DATABASE}.{SCHEMA}.{EAGS_QUOTATION_TABLE} WHERE INSERT_DATE IS NOT NULL ORDER BY INSERT_DATE DESC LIMIT 1"
        if baker:
            query = f"SELECT QUOTENO FROM {DATABASE}.{SCHEMA}.{EAGS_BAKER_TABLE} WHERE INSERT_DATE IS NOT NULL AND QUOTENO like '%{cx_init_name_previous}%' ORDER BY QUOTENO desc LIMIT 1"  
        elif baker and newQuote:
            query = f"SELECT QUOTENO FROM {DATABASE}.{SCHEMA}.{EAGS_QUOTATION_TABLE} WHERE INSERT_DATE IS NOT NULL AND QUOTENO not like '%R%' and QUOTENO like '%{cx_init_name_previous}%' ORDER BY QUOTENO  desc LIMIT 1"
        elif newQuote:
            query = f"SELECT QUOTENO FROM {DATABASE}.{SCHEMA}.{EAGS_QUOTATION_TABLE} WHERE INSERT_DATE IS NOT NULL AND QUOTENO not like '%R%' and QUOTENO like '%{cx_init_name_previous}%' ORDER BY QUOTENO  desc LIMIT 1"
            
        else:
            query = f"SELECT QUOTENO FROM {DATABASE}.{SCHEMA}.{EAGS_QUOTATION_TABLE} WHERE INSERT_DATE IS NOT NULL AND QUOTENO like '%{cx_init_name_previous}%' ORDER BY QUOTENO desc LIMIT 1"

        # data = conn.execute(query)
        # raw_data = data.fetchall()


        cur = conn.cursor()
        # cur.execute(query)
        # raw_data = cur.fetch_pandas_all()
        cur.execute(query)
        names = [ x[0] for x in cur.description]
        rows = cur.fetchall()
        raw_data= pd.DataFrame( rows, columns=names)


        if len(raw_data) and "_" in raw_data.iloc[0,0]:
            raw_data = raw_data.iloc[0,0]#raw_data[0][0]
            # revIndex = raw_data.rfind("/")
            # data = raw_data[:revIndex]

            # curr_revIndex = curr_quoteNo.rfind("/")
            # curr_data = curr_quoteNo[:curr_revIndex]
            # if curr_data==data:
            #     inpNum = raw_data.split("/")[-1]
            #     nextNum = str(int(inpNum)+1).zfill(6)
            #     newData = data+"/"+nextNum
            # else:
            #     newData = curr_quoteNo

            #########New quote no logic##############
            data = raw_data.split("_")[0]
            curr_data = curr_quoteNo.split("_")[0]
            

            if curr_data==data:
                # inpNum = raw_data.split("_")[0]
                if raw_data.split("_")[-1].rfind("R")!=-1:
                    curr_revIndex = raw_data.split("_")[-1].rfind("R")
                    nextNum=str(int(raw_data.split("_")[-1][curr_revIndex+1:])+1)
                    old_name=raw_data.split("_")[-1][:curr_revIndex+1]
                    newData = data +"_"+ old_name + nextNum
                else:    
                    newData = raw_data + "R1"
            else:
                newData = curr_quoteNo

        else:
            newData = curr_quoteNo
        # print(data)
        return newData, raw_data
    except Exception as e:
        raise e
def getallquotes(conn,quote_number):
    try:

        # conn = cnn.cursor()
        # conn = engine.connect()
        if 'Baker' in quote_number:
            tablename=EAGS_BAKER_TABLE
        else :
               tablename= EAGS_QUOTATION_TABLE
        query = f"select QUOTENO,DATE,PREPAREDBY from {DATABASE}.{SCHEMA}.{tablename} WHERE contains(QUOTENO, '{quote_number}');"
        # raw_data = conn.execute(f"SELECT NAME FROM {DATABASE}.{SCHEMA}.{table} WHERE USERNAME = '{user}' AND PASSWORD = '{pwd}'")
        cur = conn.cursor()
        cur.execute(query)
        # data = cur.fetch_pandas_all()
        names = [ x[0] for x in cur.description]
        rows = cur.fetchall()
        quote_numbers= pd.DataFrame( rows, columns=names)
        # data = raw_data.fetchall()
        # if len(data):
        #     return data[0][0]#data.iloc[0,0]
        # return False
        return quote_numbers
    except Exception as e:
        raise e
    finally:
        pass
#Driver
# curr_quoteNo = 'EAGS/USA/2022/000001'
# conn = get_connection()
# getLatestQuote(conn,curr_quoteNo)
def getfullquote(conn,quote_number):
    try:

        # conn = cnn.cursor()
        # conn = engine.connect()
        if 'Baker' in quote_number:
            tablename=EAGS_BAKER_TABLE
        else :
               tablename= EAGS_QUOTATION_TABLE
        query = f"select * from {DATABASE}.{SCHEMA}.{tablename} WHERE QUOTENO='{quote_number}';"

         
        # raw_data = conn.execute(f"SELECT NAME FROM {DATABASE}.{SCHEMA}.{table} WHERE USERNAME = '{user}' AND PASSWORD = '{pwd}'")
        cur = conn.cursor()
        cur.execute(query)
        # data = cur.fetch_pandas_all()
        names = [ x[0] for x in cur.description]
        rows = cur.fetchall()
        quotedf= pd.DataFrame( rows, columns=names)
        # data = raw_data.fetchall()
        # if len(data):
        #     return data[0][0]#data.iloc[0,0]
        # return False
        return quotedf
    except Exception as e:
        raise e
    finally:
        pass

