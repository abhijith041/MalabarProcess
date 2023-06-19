*** Settings ***
Library    AddVoucherDataRow.py
Library    voucherExtraction.py
Library    RPA.RobotLogListener
Library    String
Library    update_exceltoDB
Resource    VoucherDataExtraction.robot


*** Variables ***

${input_file}    E:\\malabarProcess\\Voucher 31\\Voucher 47.pdf
# ${input_file}    E:/malabarProcess/Petty Samples 1/Petty Samples/Project Petty/MGD INDORE PHEONIX 73-120/73-120/116.pdf
${username}    root
${password}    root
# ${azure_key}    34729f664d044096bc9e06e162f7a47e
# ${endpoint}    https://quadanceocrgroup.cognitiveservices.azure.com/


*** Keywords ***
extracting voucher data from vouchers
    [Arguments]    ${input_File}    ${endpoint}    ${azure_key}
    # collect extracted voucjer data in a dictionary 'voucherData'
    ${voucherData}    extract voucher data    ${input_File}    ${endpoint}    ${azure_key}    ${voucher_legalentity}     ${legal_entity_code}   
    
    
    #  return dictionary key
    # 'voucherNo','amount','issuedTo','date','SupportDoc','SupportDocAmount'
    IF   "${voucherData}" != "error" and "${voucherData}" != "${None}"
        Log To Console    ${\n}
        Log To Console    dict is: ${voucherData}
        ${voucher_amount}    Replace String    ${voucherData['amount']}    .00    ${EMPTY}
        Set Global Variable    ${voucher_amount}
        Log To Console    Amount is :${voucher_amount}
        Log To Console    Voucher number is :${voucherData['voucherNo']}
        Log To Console    Voucher date is :${voucherData['date']}
        Log To Console    Voucher issued to is :${voucherData['issuedTo']}

        # Log To Console    Supporting Document found :${voucherData['SupportDoc']}
        
        RETURN    ${voucherData}
    ELSE
        ${status}     Set Variable    Invalid
        ${comment}     Set Variable    Failed to extracting the pdf files
        update_status_to_DB     ${status}      ${comment}     ${voucher_legalentity}     ${pymysql_connection}
        RETURN    'error'
    END

check voucher data matching based on input file
    # for checking extracted voucher data matching based on imput excel file. If  any mismatch found, it marked as invalid and 
    # otherwise mark it as valid
    # input     dictionary:    ${voucherData} - extracted dictionary
    # input     dictionary:    ${input_data_table_element} - input excel sheet data item based on voucher number. 

    [Arguments]    ${voucherData}    ${input_data_table_element}
    
    
    ${inputSheetDate}    Set Variable    ${input_data_table_element['date']}

    

    ${voucherNumber}    Set Variable    ${input_data_table_element['voucher_number']}
    ${voucherNumber}    Convert To Upper Case    ${voucherNumber}

    ${debit_amount}=    Convert To String    ${input_data_table_element['debit']}
    ${debit_amount}=    Replace String    ${debit_amount}    .0    ${EMPTY}

    
    ${issued_to}    Set Variable    ${input_data_table_element['issued_To']}
    ${issued_to}    Convert To Upper Case    ${issued_to}

    # converting input sheet date to another date format '%d-%b-%Y'
    ${convertedDate}     Convert Date To Custom Format    ${inputSheetDate}
    ${comparison_required}=    Set Variable    ${True}
    Log To Console    ${comparison_required}
    ${comment}    Set Variable    ${EMPTY}
    #Set Global Variable    ${comment}

        # check supporting document available or not. If not comparison required
    IF    ${voucherData['supportingDocFound']}
        Log To Console    supporting document found

        # comparing amount from voucher and amount from supporting
        IF    ${comparison_required} == $True
            ${comment}    Set Variable    extracted amount from voucher and amount from supporting document is not matching
            comparing extracted data    ${voucher_amount}   ${voucherData['supportingDocAmount']}    ${voucherData}      ${comment}
        END

        # comparing voucher number
        IF    ${comparison_required} == $True
            ${comment}    Set Variable    Voucher number mismatched
            comparing extracted data    ${voucherData['voucherNo']}     ${voucherNumber}    ${voucherData}      ${comment}
    
        END
        
        # comparing amount
        IF    ${comparison_required} == $True
            ${comment}    Set Variable    Value mismatch: Amount
            comparing extracted data    ${voucher_amount}      ${debit_amount}     ${voucherData}    ${comment}
            
        END
    
        # comparing date : checking with converted date ${convertedDate}
        IF    ${comparison_required} == $True
            ${comment}    Set Variable    Value mismatch: Date
            comparing extracted data    ${voucherData['date']}    ${convertedDate}    ${voucherData}    ${comment}
            
        END

        # comparing issuedTo 
        IF    ${comparison_required} == $True
            ${comment}    Set Variable    Value mismatch: Issued to
            comparing extracted data    ${voucherData['issuedTo']}    ${issued_to}    ${voucherData}    ${comment} 
            
        END
        IF    ${comparison_required} == $True
            ${status}    Set Variable    Valid
            ${comment}    Set Variable    Voucher validation process is success
            add data item to database      ${voucherData}
            update_status_to_DB     ${status}      ${comment}     ${voucher_legalentity}     ${pymysql_connection}   
        END


    ELSE
        ${status}    Set Variable    Invalid
        ${comment}    Set Variable    Supporting document is not found
        Log To Console    supporting document not found
        add data item to database      ${voucherData}
        update_status_to_DB     ${status}      ${comment}     ${voucher_legalentity}     ${pymysql_connection}
    
    END

comparing extracted data
    # comparing voucher number,date, amount, issuedTo
    # input    string : voucherDataItem    = extracted voucher data item from dictionary
    # input    string : inputDataItem      = input excel data item from db

    [Arguments]    ${voucherDataItem}    ${inputDataItem}    ${voucherData}    ${comment}

    IF    '${voucherDataItem}' == '${inputDataItem}'
        Log To Console    '${voucherDataItem}' = '${inputDataItem}' Match found
    ELSE
        
        Log To Console    ${inputDataItem} and ${voucherDataItem} not matching  
        Log    ${inputDataItem} and ${voucherDataItem} no match found      
        
        ${status}    Set Variable    Invalid
        # call keyword to update input table data item's status as invalid.        
        update_status_to_DB     ${status}      ${comment}     ${voucher_legalentity}     ${pymysql_connection}
        add data item to database      ${voucherData}
        ${comparison_required}=    Set Variable    ${False}
        Set Global Variable    ${comparison_required}
    END

add data item to database
    # add extracted voucher data, other details to database
    # input     dictionary:    ${voucherData} - extracted dictionary
    [Arguments]    ${voucherData}


    ${dataAddReturn}    Add Data Row To Table    ${voucherData}     ${engine_str}  
    Log To Console    uploading the extracted data to db table is ${dataAddReturn}
    
    



    

# *** Tasks ***
# extract all details
#     extracting voucher data from vouchers