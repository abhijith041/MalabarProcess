import pandas as pd
from sqlalchemy import create_engine
import pymysql 
import pyodbc
from sqlalchemy.orm import sessionmaker
from openpyxl import load_workbook
from sqlalchemy import update, text
import configVariables
import shutil


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
            connection.execute(f"DELETE FROM {table_name} WHERE status = 'Invalid'")

        # Close the session
        session.close()
        return "success"
    except Exception as e:
        print("exception found is", str(e))
        return "failed"



def FetchDB_For_FinalReport(inputdata_table, senderid):
    
    """
    This is used to fetch data from database having status columns New
    then all sorted data is returned based on the legal entity.

    """
    try:
            
        #engine = create_engine('mysql+pymysql://root:Password.123@127.0.0.1:3306/malabarprocessdb')
        engine = create_engine(configVariables.dbConnectionString)

        mysql_query = f"SELECT * FROM {inputdata_table} WHERE mail_id='{senderid}' AND mail_Sent = 0"
        #query = "SELECT * FROM project_petty_table WHERE status = 'New'"

        # Execute the query and fetch data into a DataFrame
        df = pd.read_sql(mysql_query, con=engine)   

        #invoke mail sent status update function to make it true
        mail_sent_status = update_mail_sent_status(configVariables.dbConnectionString, inputdata_table, df)

        df_list = df.to_dict(orient='records')
        return df_list
    except Exception as e:
         print("An error occurred:", str(e))




def update_status_toExcel_for_report(df_list, filename):
    try:
        #filename = 'C:/Users/Q0037/Downloads/MalabarProjectGit/MalabarProjectGit/InputFolder/nibiya.monas@quadance.com/input.xlsx'    
        driver = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
        conn_str = f"Driver={{{driver}}};DBQ={filename};readOnly=0;"
        print(df_list)
       
        # Connect to Excel file using OLE DB
        cnxn = pyodbc.connect(conn_str, autocommit=True)

        # Open a cursor to execute SQL commands
        cursor = cnxn.cursor()
        #assign values here
         
         
        # Update a cell in the worksheet
        worksheet_name = "Sheet1"       
        for data in df_list:

            new_status = data['status']
            new_comments = data['comments']
            new_voucher_number = data['voucher_number']


            update_query = f"UPDATE [{worksheet_name}$A8:Z] SET status = ?, comments = ? WHERE voucher_number = ?"
            #update_query = f"UPDATE [Sheet1$A8:M] SET status = 'invalid', comments = 'test comments' WHERE voucher_number = '75'"
           # update_sql = f"UPDATE [{worksheet_name}$A8:J] SET duplication = ? WHERE voucher_number = ?"
            cursor.execute(update_query, [new_status, new_comments, new_voucher_number])

            # Define the SQL statement to update the cell
            #update_sql = f"UPDATE [{worksheet_name}$A8:J] SET status = ? WHERE voucher_number = ?"
            # update_query = f"UPDATE [{worksheet_name}$A8:J] SET status = '{new_status}', comments = '{new_comments}' WHERE voucher_number = '{new_voucher_number}'"          #'{voucher_number}'

            # # Execute the SQL statement to update the cell
            # cursor.execute(update_query)
        

        # Commit the changes
        cnxn.commit()

        cursor.close()
        cnxn.close()
        return  True
    except Exception as e:
         print("An error occurred:", str(e))
         return  False
    #getprocessed_data('project_petty_table', 'nibiya.monas@quadance.com', 'C:/Users/Q0037/Downloads/MalabarProjectGit/MalabarProjectGit/InputFolder/nibiya.monas@quadance.com/input.xlsx')



def move_to_proccessed_folder(source_folder, destination_folder):
    '''for moving folder from niput folder to processed folder. Processed folder must be there in mail directory. '''
    try:
        shutil.move(source_folder, destination_folder)

        print("Folder moved successfully.")

    except Exception as e:

        print("Error occurred while moving the folder:", str(e))