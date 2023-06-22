*** Settings ***
Documentation       Process Title : Process description

Library             Collections
Library             MyLibrary
Library             String
Library             signaturechecking.py
Library            voucherExtraction
Library             update_exceltoDB.py
Library            RPA.Excel.Files
Variables      ../variables/variables.py
Library        Collections
Library        RPA.Archive
Library        RPA.FileSystem
Library        RPA.PDF
#Library        fetchFromDBTable.py 
Library        excelFileFormatCheck.py
# Resource        emailFetching.robot
Library        runreportCreation.py
#Library        OperatingSystem 
Resource        DocExtractAndDBUpddate.robot
Resource        emailFetchingPy.robot

 

*** Keywords ***
Read config file 
    #here read the bot will read the config file and return the values as dictionary and declared ${config} as global variable
    &{out_config}=  Create Dictionary

    Open Workbook      ${CONFIG_FILE}
    FOR    ${sheet}    IN    @{CONFIG_SHEETS}
        #Log Message   Reading worksheet: ${sheet}
        ${table}=  Read Worksheet As Table    ${sheet}  header=${True}
        FOR    ${row}    IN    @{table}
            IF    "${row['Name']}" != "${null}"
                Set To Dictionary    ${out_config}  ${row['Name']}  ${row['Value']}
            END
        END
    END

    #Assigning all the global variables here 
    Set Global Variable    ${CONFIG}    ${out_config}
    ${engine_str}=     Set Variable    ${CONFIG['engine_connection_string']}
    ${pymysql_connection}=    Set Variable    ${CONFIG['pymysql_connection_string']}
    #${inputfolderpath}=    Set Variable        ${CONFIG['RootFolder']}
    
    Set Global Variable    ${engine_str}
    Set Global Variable     ${pymysql_connection}

    Close Workbook

moving folder from input folder

    ${path}    get_init_details   
    ${temp_path}    Catenate    ${path}${CONFIG['RootFolder']}
    # ${proccessed_folder}    Catenate    ${path}${CONFIG['Processfolder']}
    ${proccessed_folder}    Catenate    ${CONFIG['Processfolder']}
    #${folderpath}    Catenate     ${temp_path}
    ${input_folders}=    List Directories In Directory     ${temp_path}

    FOR    ${folder}    IN    @{input_folders}
        Log    ${folder}
        # Move To Proccessed Folder    ${folder}     ${proccessed_folder}
        Move To Processed Folder    ${folder}    ${proccessed_folder}   
    END



processing each folders
#here bot invoke the code to load the data to database and fetch each row from db then check for pdf then extracting required data
    # ${path}    Set Variable    ${CURDIR}
    # ${newPath}     Evaluate    os.path.dirname('''${path}''')
    ${path}    get_init_details   
    ${temp_path}    Catenate    ${path}${CONFIG['RootFolder']}
    # ${proccessed_folder}    Catenate    ${path}${CONFIG['Processfolder']}
    ${proccessed_folder}    set variable    ${CONFIG['Processfolder']}
    #${folderpath}    Catenate     ${temp_path}
    ${input_folders}=    List Directories In Directory     ${temp_path}
   

    ${count}    Get Length    ${input_folders}
    Log   total folder in input fodler: ${count} 

    FOR  ${folder}  IN  @{input_folders}
        Log To Console    Current processing foler: ${folder}
        set Global Variable    ${folder}
        ${mail_id}=    Get File Name    ${folder}
        ${Excel_files}=    List Files In Directory    ${folder}
        ${file_count}    Get Length    ${Excel_files}
        Continue For Loop If     '${file_count}' == '0'
        # IF     '${file_count}' == '0'           
        #     Continue For Loop            
        # END
        ${run_required}    Set Variable    ${True}
        Set Global Variable     ${run_required}
        ${run_required}    Get the excel file and zip file    ${Excel_files}    ${folder} 
        Continue For Loop If     ${run_required} == $False
                                                                            
        Fetching each row from DatabaseProcess    ${folder}
        report fetching and send mail     
        # Move To Proccessed Folder    ${folder}     ${proccessed_folder}    
        Log    ${folder}
      
    END

 
collect email id from folder name 
    [Arguments]    ${folder_name}
    
    # spliting folder name w.r.to underscore 
    ${splitResult}    Split String    ${folder_name}    _
    # Log To Console    ${splitResult}
    ${last_item}    Remove From List    ${splitResult}    -1
    # ${concatenated}    Catenate    SEPARATOR=    ${splitResult}
    ${mail_id}    Evaluate    ''.join(${splitResult})   
    Log To Console    Mail id is: ${mail_id}
    RETURN    ${mail_id}


Get the excel file and zip file
#here getting the excel file and load to database, also unzip the zip file and store files in a folder
    [Arguments]    ${Excel_files}    ${folder}

     FOR    ${file}    IN    @{Excel_files}
            ${extension_xl}=    Get File Extension    ${file}
            #load the excel file to database
            IF    '${extension_xl}' == '.xlsx'               
                ${input_excel_path}    Set Variable    ${file}
                Set Global Variable     ${input_excel_path} 
                 ${folder_name}=    Get File Name    ${folder}
                 ${mail_id}    collect email id from folder name    ${folder_name}
                 Set Global Variable    ${mail_id}
                
                 #read the horizontal aligned data from input file
                 ${legal_entity_code}     ${cost_center_code}     ${emp_code}     ${interior_code}     ${work_code}        read_horizontaldata    ${input_excel_path}
                 #checking that if any of those horizontaly aligned data is missing
                 ${empty_primary_data}    ${comments}     check_empty_horizontaldata    ${legal_entity_code}     ${cost_center_code}     ${emp_code}     ${interior_code}     ${work_code}        
                  IF    ${empty_primary_data} != $True
                     
                   ${subject}    Set Variable    Alert: Missing mandatory Fields in excel file
                    ${body}    set variable    Hi Team, \n\nplease find the attached excel file that not containining expected data fields.\n${comments}.\n\nThank you,\nBot
                    # Dear [Recipient's Name],\n\n${messageBody}\n\nSincerely,\n[Your Name]
                    send mail with excel report    ${mail_id}    ${input_excel_path}    ${subject}    ${body}
                    ${run_required}    Set Variable    ${False} 
                    RETURN    ${run_required}
                 END

                ${excel_format}    check_excel_headers    ${input_excel_path}
                IF    '${excel_format}' == 'no_match'
                    ${Mailsubject}    Set Variable    input data sheet template is not correct
                    # send mail if exception occures    ${mail_id}    ${Mailsubject}    $sendingTime
                    ${subject}    Set Variable    Alert: Missing mandatory columns in excel file
                    ${body}    set variable    Hi Team, \n\nplease find the attached excel file that not containining expected columns.\n\nThank you,\nBot
                    # Dear [Recipient's Name],\n\n${messageBody}\n\nSincerely,\n[Your Name]
                    send mail with excel report    ${mail_id}    ${input_excel_path}    ${subject}    ${body}
                    log   send mail- excel column format mismatch
                    ${run_required}    Set Variable    ${False}                    
                    # Exit For Loop
                    RETURN    ${run_required} 

                END
                 #create columns in the input excel file [duplicate, status, comment]
                create_columns    ${input_excel_path}
                 
                 #here checking the duplication occurs if so upadte the status to excel sheet
                 duplication_checking_process     ${input_excel_path}    ${legal_entity_code}

                #upload the input sheet data to database
                 ${no_data_in_df}    ${inputdata_uploaded}      ${voucherNo_empty}   upload_input_values_DB    ${legal_entity_code}    ${cost_center_code}     ${emp_code}     ${interior_code}     ${work_code}     ${input_excel_path}    ${engine_str}     ${mail_id}                  
                 Set Global Variable    ${legal_entity_code}

                     
                    log to console    No data in dataframe after removing duplicates :${no_data_in_df}
                    Log To Console     Data upload to DB: ${inputdata_uploaded} 
                    log    No data in dataframe after removing duplicates :${no_data_in_df}
                    Log     Data upload to DB: ${inputdata_uploaded}   

                    # check if data available after removing duplicate data from input excel
                    IF    ${no_data_in_df} == $True
                        Log    Failed to upload input file to db
                        ${subject}    Set Variable    Alert: Duplicate Data found.
                        ${body}    set variable    Hi, \nUnable to continue running with this input data as all the data from the input files has already been stored in the database.\n\nThank you,\nBot
                        send mail with excel report    ${mail_id}    ${input_excel_path}    ${subject}    ${body}
                        ${run_required}    Set Variable    ${False}   
                        # Exit For Loop
                        RETURN    ${run_required}
                    END


                    IF   '${voucherNo_empty}' == 'invalid'
                        Log To Console    voucher number is missing in one of the rows
                        Log    voucher number is missing in one of the rows
                        Exit For Loop                      
                        
                    END
                    IF    ${inputdata_uploaded} == True
                        Log To Console    Sucessfully completed DB updation  
                    ELSE
                    # flg
                        Log To Console    Failed to upload input file to db
                        Log    Failed to upload input file to db
                        ${subject}    Set Variable    Alert: Database updation failed.
                        ${body}    set variable    Hi, \n\nCurrently unable to update data to database due to technical error Please fix that.\n\nThank you,\nBot
                        send mail with excel report    ${systemadminMailId}    ${input_excel_path}    ${subject}    ${body}
                        ${run_required}    Set Variable    ${False}   
                        # Exit For Loop
                        RETURN    ${run_required} 

                    END                        
            END
            #unzipping the zip files
            IF    '${extension_xl}' == '.zip'
                Log     zip file is extracted: ${file}
                Extract Archive     ${file}    ${folder}
                Remove File    ${file}
                
            END

        END
        #[RETURN]     ${input_excel_path}

check_empty_horizontaldata
    [Arguments]    ${legal_entity_code}     ${cost_center_code}     ${emp_code}     ${interior_code}     ${work_code}
    ${empty_primary_data}    Set Variable    ${True}
    ${comments}    Set Variable    missing item :

    IF    '${legal_entity_code}' == 'Not found'
        ${comments}    Set Variable    ${comments}legal entity name is missing,
        ${empty_primary_data}    Set Variable    ${False}      
    END

    IF    '${cost_center_code}' == 'Not found'
        ${comments}    Set Variable     ${comments}cost center code name is missing,  
        ${empty_primary_data}    Set Variable    ${False}  
    END

    IF    '${emp_code}' == 'Not found'
        ${comments}    Set Variable    ${comments}emp code name is missing,
        ${empty_primary_data}    Set Variable    ${False}      
    END
    IF    '${interior_code}' == 'Not found'
        ${comments}    Set Variable    ${comments}interior code name is missing,  
        ${empty_primary_data}    Set Variable    ${False}      
    END 

    IF    '${work_code}' == 'Not found'
        ${comments}    Set Variable    ${comments}work code name is missing,
        ${empty_primary_data}    Set Variable    ${False}        
    END
    RETURN    ${empty_primary_data}    ${comments}

Fetching each row from DatabaseProcess
    [Arguments]    ${folder}
#----------------------------fetching each row from database and store to list---------------------------------------------------------------
      
    ${input_data_table} =   read_data_from_database
    Log    ${input_data_table}

    FOR    ${input_data_table_element}    IN    @{input_data_table}

       
        ${vr_no}=    Set Variable    ${input_data_table_element['voucher_number']}
        ${voucher_legalentity}=    Set Variable    ${input_data_table_element['voucher_legalentity']}
        Set Global Variable     ${voucher_legalentity}
        ${vr_name}=    Set Variable    ${vr_no}.pdf             
        #${pdf_exist}=    Loop through input folders     ${vr_name}    ${folder}
        ${pdf_exist}=   Find Matching PDF Files     ${folder}    ${vr_name}    ${input_data_table_element}
        IF     ${pdf_exist} == $False  

            ${status}=     Set Variable    Invalid
            ${comments}=    Set Variable     No voucher pdf exist         
            Log To Console    ${vr_no}-Update to DB- No voucher document exist
            Log To Console    voucher: ${vr_no} is Invalid
            ${update_status}    update_status_to_DB    ${status}    ${comments}   ${voucher_legalentity}      ${pymysql_connection}
            Log To Console     Db loading status is ${update_status} 

        # ELSE
        #     ${status}=     Set Variable    Valid
        #     ${comments}=    Set Variable      voucher pdf exist   
        #      ${update_status}     update_status_to_DB    ${status}    ${comments}   ${voucher_legalentity}      ${pymysql_connection}
           
        #     Log To Console    ${vr_no}-Update to DB-  voucher document exist
        END  
    END
    Log    key done-Fetching each row from DatabaseProcess




Find Matching PDF Files
#------------------------Below code is for, Once file unzipped bot needs to find the matching voucher pdf in the folder------------------------

    [Arguments]    ${folder}    ${vr_name}    ${input_data_table_element}
    ${pdf_exist}=    Set Variable   ${False}
        ${pdf_files}=    List Files In Directory     ${folder}
        ${skip_subfolder}    Set Variable    ${False}       

        FOR    ${element}    IN    @{pdf_files}
                ${extension}=    Get File Extension    ${element}
                IF    '${extension}' == '.pdf'
                    ${skip_subfolder}    Set Variable    ${True}             
                     Exit For Loop    
                END
        END
    #----------------------------------------Case2 pdf checking---------------------------------------------------------------------
        IF    ${skip_subfolder} != $True
            ${sub_folders}=    List Directories In Directory    ${folder}
            #${zip_sub_folders}    List Directories In Directory    ${sub_folders}
            FOR    ${sub_folder}    IN    @{sub_folders}
                ${filename}=    Get File Name    ${sub_folder}
                ${zip_sub_folders}    List Directories In Directory    ${sub_folder}
                ${getlength}    Get Length    ${zip_sub_folders}
                IF    ${getlength} > 0
                    ${pdf_files}=    List Files In Directory        ${zip_sub_folders}[0]      
                    Log To Console    pdf file exist in the child folder of subfodler  
                ELSE
                    ${pdf_files}=    List Files In Directory    ${sub_folder}
                    Log To Console    pdf file exist in the sub folder       

                END        

            END

        END

        FOR    ${element}    IN    @{pdf_files}
            ${extension}=    Get File Extension    ${element}
            IF    '${extension}' == '.pdf'               

                ${pdf_file_name}=    Get File Name    ${element}     
                IF    '${vr_name}' == '${pdf_file_name}'  
                    ${pdf_file_path}=    Set Variable     ${element}                  
                    ${pdf_exist} =    Set Variable    ${True}
                    # ${azure_key}     Set Variable        34729f664d044096bc9e06e162f7a47e
                    # ${endpoint}     Set Variable      https://quadanceocrgroup.cognitiveservices.azure.com/
                    
                    # collecting azure endpoint ans subscription key from Config file
                    ${azure_key}    Set Variable    ${CONFIG}[SubscriptionKey]
                    ${endpoint}    Set Variable    ${CONFIG}[EndPoint]

                    #Here need to pass the matched pdf file name to next keyword for voucher signature match
                    ${sign_status}    sign_checking    ${pdf_file_path}                    
                   
                     IF    '${sign_status}' == 'error'
                        ${status}=     Set Variable    Invalid
                        ${comments}=    Set Variable    'issues in signature detection ${vr_name}'
                        ${update_status}=     update_status_to_DB    ${status}    ${comments}   ${voucher_legalentity}      ${pymysql_connection}
                        Exit For Loop

                    ELSE IF    ${sign_status} == $True
                        ${status}=     Set Variable    Invalid
                        ${comments}=    Set Variable    'Duplicate signature detected'
                        ${update_status}=     update_status_to_DB    ${status}    ${comments}   ${voucher_legalentity}      ${pymysql_connection}
                        Exit For Loop

                    END
                    #below code is to extract voucher pdf data
                     ${voucherData}    extracting voucher data from vouchers      ${pdf_file_path}    ${endpoint}    ${azure_key}
                     Exit For Loop If     ${voucherData} == 'error'
                    #below code is to compare the values with input file data
                    ${status}    ${comments}    Run Keyword And Ignore Error     check voucher data matching based on input file     ${voucherData}    ${input_data_table_element}
                    Run Keyword If   '${status}' == 'FAIL'   update_status_to_DB    ${status}    ${comments}   ${voucher_legalentity}      ${pymysql_connection}

                    #Exit For Loop   

                ELSE
                    Log    pdf doesn't exists for ${vr_name}

                    #${pdf_exist} =    Set Variable    'False'                 
                END
            END
        END
    Log    keyword done- Find Matching PDF Files
    RETURN     ${pdf_exist}
 
 
report fetching and send mail

    
    # ${senderid} =     Get File Name    ${folder}
    ${df_list}    FetchDB_For_FinalReport    ${INPUT_DB_TABLE}    ${mail_id}
    Log To Console    final data list returned from db:${df_list}
    ${is_updated}    update_status_toExcel_for_report      ${df_list}    ${input_excel_path}
    Log To Console    updating to the final report ${is_updated}    
    ${subject}    Set Variable    Success: Run report
    ${body}    Set Variable    Hi Team, \n\nplease find the attached excel file containing final run report.\n\nThank you,\nBot
    send mail with excel report     ${mail_id}     ${input_excel_path}  ${subject}    ${body}
    # send mail with excel report    $senderId    $reportpath    $subject





    # ${senderid} =     Get File Name    ${folder}
    # fetch_data_from_database    ${engine_str}    ${senderid}    ${INPUT_DB_TABLE}    ${folder}
    # ${reportpath}    Catenate    ${folder}/finalreport.xlsx
    # ${subject}    Set Variable    Success: Run report
    # ${body}    Set Variable    Hi Team, \n\nplease find the attached excel file containing final run report.\n\nThank you,\nBot
    # send mail with excel report     ${senderid}     ${reportpath}  ${subject}    ${body}
    # # send mail with excel report    $senderId    $reportpath    $subject