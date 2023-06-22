*** Settings ***
Documentation       Template robot main suite.

Library             Collections
Library             MyLibrary 
Resource            VoucherDataExtraction.robot        
# variables           variables.py
Variables            variables.py
Resource            emailFetchingPy.robot


*** Tasks ***
petty cash processing bot

    ${config_status}=    Run Keyword And Return Status    Read config file 
    
    IF    ${config_status} == $True
         ${mail_status}=    Run Keyword And Return Status    email fetching based on subject         
    END

    IF    ${mail_status} == $True
        ${DB_upload_Status}=    Run Keyword And Return Status    processing each folders
    END
    moving folder from input folder
    # processing each folders
    Log  Done
