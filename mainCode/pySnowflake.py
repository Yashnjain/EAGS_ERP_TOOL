import pandas as pd
import time
import bu_snowflake
import snowflake.connector
from snowflake.connector.pandas_tools import pd_writer
from snowflake.connector.pandas_tools import write_pandas



USER =  "SVC_BUIT"#  "SVC_BUITDB_DEV"
PASSWORD = "BUIT2022"
ACCOUNT = 'OS54042.east-us-2.azure'
WAREHOUSE = "BUIT_WH"
DATABASE = "BUITDB"

SCHEMA = "EAGS"
ROLE = "OWNER_BUITDB"
# TABLE = 'EAGS_INVENTORY'
TABLE = 'EAGS_CUSTOMER'
# TABLE = 'EAGS_SALESPERSON'


def insert_data_to_db(df):
    
    conn=snowflake.connector.connect(

            user=USER,
            password=PASSWORD,
            account=ACCOUNT,
            warehouse=WAREHOUSE,
            database  = DATABASE,
            schema=SCHEMA,
            role =ROLE
        )  

    # engine = bu_snowflake.get_engine(
    #     username=USER,
    #     password=PASSWORD,
    #     warehouse=WAREHOUSE,
    #     role=ROLE,
    #     schema=SCHEMA,
    #     database=DATABASE
    #     )
        
    # con = engine.connect()
    df.columns = map(str.upper, df.columns)
    success, nchunks, nrows, _ = write_pandas(conn, df, TABLE)
    # df.to_sql(name=TABLE, con=cnn, index=False,if_exists='append', schema=SCHEMA, method=pd_writer)
    print("Data successfully uploaded")
    conn.close()




# df = pd.read_csv(r'C:\Users\imam.khan\Downloads\Inventory_eags.csv')

#Customer dataframe
# df = pd.read_excel(r'cxDatabase.xlsx')
# df['PAYMENT_TERM'] = df['PAYMENT_TERM'].astype(str)
# df['CUS_CITY_ZIP'] = df['CUS_CITY_ZIP'].astype(str)

#SalespersonDataframe
df = pd.read_excel("Customer info.xlsx", sheet_name="Data", usecols="A:H")
insert_data_to_db(df)
# df.to_csv("Inv.csv",index=False,header=False)

print(df)