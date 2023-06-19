import pandas as pd
import re
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.automap import automap_base
from sqlalchemy.engine.url import make_url
from decimal import Decimal
import  configVariables


from sqlalchemy import update, text



def update_mail_sent_status(connection_string, table_name, dataframe):
    '''update mailSent column value as true so that it doesn't need to fetch at finalreport collection time
    
    connection_string   :   dbconnection string
    table_name          : table name that need to update
    dataframe           : collected dataframe
    '''

    try:
        # Create the engine using the connection string
        engine = create_engine(connection_string)

        Session = sessionmaker(bind=engine)
        session = Session()

        # Update the 'mailsent' column in the MySQL table
        with engine.begin() as connection:
            for index, row in dataframe.iterrows():
                voucherno_legal_entity = row['voucher_legalentity']
                # voucherno_legal_entity ='15MLB201'
                query = text(f"UPDATE {table_name} SET mail_sent = :value WHERE voucher_legalentity = :entity")
                connection.execute(query, {"value": True, "entity": voucherno_legal_entity})

        # Close the session
        session.close()
        return "success"
    except Exception as e:
        print("exception found is", str(e))
        return "failed"




def fetch_data_from_database(connection_string, senderid, inputdata_table, filepath):
    '''fetching data from database for run report
    senderid            :   mailsender id
    connection_string   :   dbconnection string

    return              :df if success, else return 'failed' 
    '''


    
    # Parse the connection string to create the URL object
    try:
        df = None
        mailSentStatus =False
        mysql_query = f"SELECT * FROM {inputdata_table} WHERE mail_Sent = 0 AND mail_id='{senderid}'"
        url = make_url(configVariables.dbConnectionString)

        # Create the engine using the URL object
        engine = create_engine(url)

        # Create a session factory
        Session = sessionmaker(bind=engine)

        # Create a session
        session = Session()

        # Reflect the database schema and create mapping
        Base = automap_base()
        Base.prepare(engine, reflect=True)
        filepath = str(filepath)+'/finalreport.xlsx'
        # Extract the table name from the query using regular expressions
        table_name_match = re.search(r'FROM\s+(\w+)', mysql_query, re.IGNORECASE)
        if table_name_match:
            table_name = table_name_match.group(1)
        else:
            raise ValueError("Table name not found in the query.")

        # Choose the appropriate table based on the extracted table name
        if table_name in Base.classes:
            Table = Base.classes[table_name]
        else:
            raise ValueError(f"Table '{table_name}' not found in the database schema.")

        # Fetch the column names from the database table
        column_names = [column.name for column in Table.__table__.columns]

        # Execute the query and fetch all the rows
        rows = session.query(Table).all()

        # Close the session
        session.close()

        # Convert the list of objects into a DataFrame
        df = pd.DataFrame([row.__dict__ for row in rows])
        
        # columns_to_drop = ["run_status", "mail_id", "mail_sent"]
        # df = df.drop(columns=columns_to_drop)       
        # tablename='malabartable'
        updationPart = update_mail_sent_status(configVariables.dbConnectionString, table_name, df)
        print(updationPart)

        # Reorder the DataFrame columns based on the table column order
        df = df[column_names]

        # Drop the '_sa_instance_state' column
        df = df.drop('_sa_instance_state', axis=1, errors='ignore')
        print(df)
        df = df.drop('run_status', axis=1, errors='ignore')
        df = df.drop('mail_id', axis=1, errors='ignore')
        df = df.drop('mail_sent', axis=1, errors='ignore')
        print(df)
        df['debit'] = df['debit'].apply(lambda x: float(x))
        df['debit'] = df['debit'].round(2)

        df.to_excel(filepath, index=False)
        # return df
        return "success"
    
    except Exception as e:
        print("Exception occurred and the message is:", str(e))
        return "failed"





# #fetch_data_from_database('mysql+pymysql://root:Password.123@127.0.0.1:3306/malabarprocessdb', 'Nehamathew@gmail.com', 'project_petty_table', 'C:/Users/Q0037/Documents/Malabar/VoucherVerificationProcess/InputFolder/Nehamathew@gmail.com')


# #fetch_data_from_database('mysql+pymysql://root:Password.123@127.0.0.1:3306/malabarprocessdb','Nehamathew@gmail.com','proect_petty_table','C:/Users/Q0037/Documents/Malabar/VoucherVerificationProcess/InputFolder/Nehamathew@gmail.com/mysql_result.xlsx')



# Example usage

# Connection string for MySQL. replace your conenction string with this.
#mysql_connection_string = "mysql+pymysql://root:root@127.0.0.1:3306/1mgfulldata"

# Connection string for SQL Server
# sql_server_connection_string = "mssql+pyodbc://username:password@dsn=dsn_name"

# Fetch data from MySQL
# mysql_query = "SELECT * FROM malabartable;" where mailid='' AND mailSent =False"
# mysql_query = f"SELECT * FROM malabartable where mailid='{senderid}' AND mailSent =False"

#senderId= 'abcd@gmail.com'
#mysql_df = fetch_data_from_database(mysql_connection_string,senderId)

# Convert 'debit' column to float and round to 2 decimal places
# mysql_df['Amount'] = mysql_df['Amount'].apply(lambda x: float(x))
# mysql_df['Amount'] = mysql_df['Amount'].round(2)

# mysql_df['Amount'] = mysql_df['Amount'].apply(lambda x: float(x))
# mysql_df['Amount'] = mysql_df['Amount'].round(2)


# Write the MySQL result to an Excel file
#mysql_df.to_excel("mysql_result.xlsx", index=False)

