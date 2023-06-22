from sqlalchemy import create_engine, Column, Integer, String, Date, Float
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from datetime import datetime
import configVariables


def add_data_row_to_table(voucherData, engine):
    """used to add extracted voucher data to a table in Database

    voucherData  : voucher details dictionary
    engine_connection_string : from config file
    
    return       : success or failure
    """

    # voucherNoLegalEntity = '13_MLB'
    
    try:

        # Define the database connection string
        #db_connection_string = 'mysql+pymysql://'+username+':'+password+'@localhost/1mgfulldata'
        db_connection_string = configVariables.dbConnectionString

        # Create the database engine and session
        engine = create_engine(db_connection_string)
        Session = sessionmaker(bind=engine)
        session = Session()

        # Define the table structure
        Base = declarative_base()

        class MalabarTable(Base):
            # update tablename here
           # __tablename__ = 'erp_transaction_table'
            __tablename__ = configVariables.erpExtractionTable
            voucher_legalentity = Column(String, primary_key=True)
            #voucherno_legal_entity = Column(String, primary_key=True)
            voucher_number = Column(String)            
            legal_entity = Column(String)
            date = Column(String)
            amount = Column(Float)
            supporting_doc_present = Column(String)
            support_doc_amount = Column(Float)
            signature_mismatch = Column(String)
            issued_to = Column(String)

        # Create an instance of your table and assign values to the columns
        # working code
        # new_row = MalabarTable(voucherno_legal_entity='13_ML203', voucher_no='13', date='05/02/2023',
        #                         Amount=150.0, Support_doc=True, support_doc_amount=150.0)


        #  return dictionary key
        # 'voucherNo','amount','issuedTo','date','SupportDoc','SupportDocAmount'
        voucher_date = datetime.strptime(voucherData['date'], '%d-%b-%Y').strftime('%Y-%m-%d')
        new_row = MalabarTable(voucher_legalentity=voucherData['VoLegalEntity'], voucher_number=voucherData['voucherNo'],
                                date=voucher_date,amount=voucherData['amount'], supporting_doc_present=voucherData['supportingDocFound'],
                                support_doc_amount=voucherData['supportingDocAmount'],legal_entity=voucherData['LegalEntity'],
                                signature_mismatch=True,issued_to=voucherData['issuedTo'])



        # Add the new row to the session and commit the changes
        session.add(new_row)
        session.commit()

        # Close the session
        session.close()

        return "success"
    
    except Exception as e:

        print("exception is :" +str(e))
        # print("".join(traceback.format_exception(*sys.exc_info())))
        return "failure"