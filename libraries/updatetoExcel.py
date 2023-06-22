import pyodbc
import pandas as pd


def update_excel_duplication(filename,where_value, New_Status):
    # Connection parameters for Excel file (adjust as necessary)     
    #filename = "C:/Users/Q0037/Documents/Robots/Malabar_Gold_Accounts_Automation3/input/challenge.xlsx"
    driver = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
    conn_str = f"Driver={{{driver}}};DBQ={filename};readOnly=0;"

    # Connect to Excel file using OLE DB
    cnxn = pyodbc.connect(conn_str, autocommit=True)

    # Open a cursor to execute SQL commands
    cursor = cnxn.cursor()

    # Update a cell in the worksheet
    worksheet_name = "Sheet1"
#--------------------------------------------------Select Query based on where value condition-----------------------------------------
    select_sql = """
        SELECT TOP 1 *
        FROM [Sheet1$]
        WHERE [Voucher_Status] = ?
    """

    cursor.execute(select_sql, where_value)
    selected_row = cursor.fetchone()

    if selected_row:
        column_names = [column[0] for column in cursor.description]

        # Retrieve the column values based on column names
        status_column_index = column_names.index('Voucher_Status')
        voucher_status = selected_row[status_column_index]

        voucher_no_column_index = column_names.index('voucher_No')
        voucher_no = selected_row[voucher_no_column_index]

        # Print the selected values for verification
        print("Selected Voucher_Status:", voucher_status)
        print("Selected Voucher_No:", voucher_no)

#------------------------------------------------------Update a paricular row having where value----------------------------------------
        # Define the SQL statement to update the cell
        update_sql = f"""
           UPDATE [{worksheet_name}$]
           SET [Voucher_Status] = ?
           WHERE [voucher_No] = ?
        """

        # Execute the SQL statement to update the cell
        cursor.execute(update_sql, New_Status, voucher_no)

        cnxn.commit()  # Commit the changes to the worksheet

        print("Cell updated successfully.")

    # Close the cursor and connection
    cursor.close()
    cnxn.close()
    return  voucher_no


#update_excel_cell('New', 'Inprogress')

# def duplication_process(LE):

#     df = pd.read_excel('C:/Users/Q0037/Documents/Malabar/VoucherVerificationProcess/InputFolder/Nehamathew@gmail.com/input.xlsx', skiprows=7)

#     for index, row in df.iterrows():
#         voucher_no = row['voucher_number']      
#         vr_le = LE+'_'+voucher_no
#         print(vr_le)

# duplication_process('MGPP')

