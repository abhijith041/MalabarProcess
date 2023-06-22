dbConnectionString  =   'mysql+pymysql://root:root@127.0.0.1:3306/malabarprocessdb'

pymysqlConntionString = "host='127.0.0.1',user='root',password='root',database='malabarprocessdb',port=3306"
                            # host='127.0.0.1',user='root',password='root',database='malabarprocessdb',port=3306

host            =   '127.0.0.1'
user            =   'root'
password        =   'root'
port            =   3306

databaseName = 'malabarprocessdb'
erpExtractionTable = 'erp_transaction_table'
projectPettyTable = 'project_petty_table'




# ---------------------- voucherExtraction.py file variables ----------------------

# update this regex if you want to try other combinations
voucher_number_pattern = [r'(?<=VR NO:)(\w+-\w+)',r'VR NO:(\S+)', r'\bKRK-[A-Z]{2}\d{3}\b', r'V\.No\s*:\s*(\d+)']

amount_patterns        = [r'Amount\s*\n\s*₹\s*([\d.,]+)',r'₹\s*([\d.,]+)',r'Amount\s+₹\s+(\d+\.\d+)',
                            r'Amount\s+(\d+\.\d+)\s*₹?',r'₹\s+(\d+\.\d+)\s+In Words',r'Amount\s+(\d+\.\d+)',]

date_pattern           = [r'Date:\s*([0-9]{2}-[A-Za-z]{3}-[0-9]{4})',r'Date:\s*([0-9]{2}-[A-Za-z]+-[0-9]{4})',
                            r'Date:\s*([0-9]{2}-[A-Za-z]{3,}-[0-9]{4})']

issued_to_pattern      = [r'(?<=Issued to:)\s*(.*?)\s*(?=_)',r'(?<=Issued to:)\s*(.*?)(?=\n)',
                            r'Issued to:\s*(.*?)(?=\s*Description:)',r'Issued to:\s*(.*?)(?=\s*-\.)',
                            r'Issued to:\s*(.*?)(?=\s*Date:)',]





# ---------------------- excelFileFormatCheck.py file variables  ----------------------

# this headers are used to identify input excel file
headers_to_check        = ['voucher_number', 'date', 'expense_category', 'type', 'issued_to', 
                           'description', 'debit', 'credit', 'closing_balance']







# ---------------------- update_ExcelToDB.py file variables  ----------------------
# column_header           = ['voucher_number', 'date', 'expense_category', 'type', 'issued_to', 
#                            'description', 'debit', 'credit', 'closing_balance','duplication','status']


# we only take only some headers from input excel file to update in DB. that header fields are here
column_header           = ['voucher_number', 'date', 'expense_category', 'type', 'issued_to', 
                           'description', 'debit', 'credit', 'closing_balance','duplication','status']
