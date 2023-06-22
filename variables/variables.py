
#name of database tables mentioned here
INPUT_DB_TABLE = "project_petty_table"
ERP_DB_TABLE = "erp_transaction_table"
CONNECTION_STRING = "mysql+pymysql://root:root@127.0.0.1:3306/malabarprocessdb"
PYMYSQL_CONN_STRING = "host='127.0.0.1',user='root',password='root',database='malabarprocessdb',port=3306"
# CONFIG_FILE = "data/config.xlsx"
CONFIG_FILE =   r'C:/malabar_config/config.xlsx'


TRANSACTION_ITEM = ""
SYSTEM_EXCEPTION = ""
BUSINESS_EXCEPTION = ""
TRANSACTION_NUMBER = 0
CONFIG = {}
#CONFIG_FILE = "config.xlsx"
CONFIG_SHEETS = ["Settings", "Constants"]
RETRY_NUMBER = 0
TRANSACTION_FIELD_1 = ""
TRANSACTION_FIELD_2 = ""
TRANSACTION_ID = ""
DT_TRANSACTION_DATA = {}
CONSECUTIVE_SYSTEM_EXCEPTION = 0
BROWSER = ""



# ---------------------- email fetching.robot file variables ----------------------

emailSourceFolder               =   'Inbox'         #   not using currently
emailDestinationFolder          =   'readedMail'    #   not using currently
EMAIL_SUBJECT                   =   ['Project Petty','marketing petty']
#
#  Define the duration in seconds (10 minutes = 600 seconds) -email polling duration
polling_duration                        =   0
# subject_line                    =   ['abc''project petty transaction',]




# ----------------------voucherDataExtraction.robot file variables ----------------------

systemadminMailId               =   'abhijith.p@quadance.com'