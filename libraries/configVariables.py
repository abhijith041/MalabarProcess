dbConnectionString  =   'mysql+pymysql://root:root@127.0.0.1:3306/malabarprocessdb'
pymysqlConntionString = "host='127.0.0.1',user='root',password='root',database='malabarprocessdb',port=3306"



# ---------------------- voucherExtraction.py file variables ----------------------


voucher_number_pattern = [r'(?<=VR NO:)(\w+-\w+)',r'VR NO:(\S+)', r'\bKRK-[A-Z]{2}\d{3}\b', r'V\.No\s*:\s*(\d+)']

amount_patterns        = [r'Amount\s*\n\s*₹\s*([\d.,]+)',r'₹\s*([\d.,]+)',r'Amount\s+₹\s+(\d+\.\d+)',
                            r'Amount\s+(\d+\.\d+)\s*₹?',r'₹\s+(\d+\.\d+)\s+In Words',r'Amount\s+(\d+\.\d+)',]

date_pattern           = [r'Date:\s*([0-9]{2}-[A-Za-z]{3}-[0-9]{4})',r'Date:\s*([0-9]{2}-[A-Za-z]+-[0-9]{4})',
                            r'Date:\s*([0-9]{2}-[A-Za-z]{3,}-[0-9]{4})']

issued_to_pattern      = [r'(?<=Issued to:)\s*(.*?)\s*(?=_)',r'(?<=Issued to:)\s*(.*?)(?=\n)',
                            r'Issued to:\s*(.*?)(?=\s*Description:)',r'Issued to:\s*(.*?)(?=\s*-\.)',
                            r'Issued to:\s*(.*?)(?=\s*Date:)',]





# ---------------------- excelFileFormatCheck.py file variables  ----------------------

headers_to_check        = ['voucher_number', 'date', 'expense_category', 'type', 'issued_to', 
                           'description', 'debit', 'credit', 'closing_balance']