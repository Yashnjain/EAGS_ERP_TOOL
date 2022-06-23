import bu_snowflake
import snowflake.connector
from snowflake.sqlalchemy import URL
from sqlalchemy import create_engine
import pandas as pd

# USER = "AKSHATA"
# PASSWORD = "Biourja@2022"
USER = "SVC_BUITDB_DEV"
PASSWORD = "BUITDBDEV2022"
ACCOUNT = 'OS54042.east-us-2.azure'
WAREHOUSE = "BUIT_WH"
DATABASE = "BUITDB_DEV"
SCHEMA = "EAGS"
ROLE = "OWNER_BUITDB_DEV"

def get_connection():
    engine = bu_snowflake.get_engine(
            username=USER,
            password=PASSWORD,
            warehouse=WAREHOUSE,
            role=ROLE,
            schema=SCHEMA,
            database=DATABASE
            )
    conn = engine.connect()
    return conn

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
        
        raw_data = conn.execute(f"SELECT NAME FROM {DATABASE}.{SCHEMA}.{table} WHERE USERNAME = '{user}' AND PASSWORD = '{pwd}'")
        data = raw_data.fetchall()
        if len(data):
            return True
        return False
    except Exception as e:
        raise e
    finally:
        pass
        # engine.dispose()
        # conn.close()

# loginChecker("table","user", "pwd")

def get_cx_df(conn,table):
    # query = f"SELECT CUS_LONG_NAME, PAYMENT_TERM, CUS_ADDRESS, CUS_PHONE, CUS_EMAIL, CUS_CITY_ZIP FROM {DATABASE}.{SCHEMA}.{table}"
    query = f"SELECT * FROM {DATABASE}.{SCHEMA}.{table}"
    df = pd.read_sql_query(query, conn)
    return df

def get_inv_df(conn, table):
    query = f"SELECT SITE, MATERIAL_TYPE, GLOBAL_GRADE, OD_IN, OD_IN_2, HEAT_CONDITION, ONHAND_PIECES, ONHAND_LENGTH_IN, RESERVED_PIECES, RESERVED_LENGTH_IN, AVAILABLE_PIECES, AVAILABLE_LENGTH_IN FROM {DATABASE}.{SCHEMA}.{table}"
    df = pd.read_sql_query(query, conn)
    return df


def get_qtylengthdf(conn, table, location, type, grade, od, id):
    query = f"""SELECT ONHAND_PIECES, ONHAND_LENGTH_IN, RESERVED_PIECES, RESERVED_LENGTH_IN, AVAILABLE_PIECES, AVAILABLE_LENGTH_IN FROM {DATABASE}.{SCHEMA}.{table} 
                WHERE SITE={location} AND MATERIAL_TYPE={type} AND GLOBAL_GRADE={grade}  AND OD_IN = {od} AND OD_IN_2 = {id}"""
    df = pd.read_sql_query(query, conn)
    return df


# conn = get_connection()
# get_cx_df(conn,table= "EAGS_CUSTOMER")
