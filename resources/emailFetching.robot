*** Settings ***

Library    RPA.Outlook.Application
Library    Collections
Library    RPA.FileSystem
Library    RPA.Cloud.Azure 
Variables    variables.py
Library    voucherExtraction
Library       MyLibrary
Library    emailFetchingPython.py
# Library    RPA.Excel.Application

*** Variables ***
# @{subjects}    email testing    email marketing

*** Keywords ***
email fetching based on subject old
    Open Application
    FOR    ${subject}    IN    @{EMAIL_SUBJECT}
        Log To Console  mail subject is: ${subject}
        log     mail subject is:${subject}

        # ${emails} =    Get Emails     folder_name=Inbox  email_filter=[Subject]=${subject}    # ${emails} =    Get Emails     folder_name=Inbox       email_filter=[unread]='${True}'
        ${emails} =    Get Emails     folder_name=Inbox  email_filter=[Subject]=${subject}
        ${emailLength}    Get Length    ${emails}
        
        Log To Console    Total email with given subject: '${subject}' is ${emailLength}
        Log     Total email with given subject: '${subject}' is ${emailLength}
        
        FOR    ${email}    IN    @{emails}
            ${attachmentDownloadFolder}     get_init_details
            Log To Console    ${email}
            ${senderId}    Get From Dictionary    ${email}    Sender
            ${sendingTime}    Get From Dictionary    ${email}    ReceivedTime
        
            Create Directory    ${attachmentDownloadFolder}/InputFolder/${senderId}    parents=${True}
            
            # collecting all attchment's name
            ${attachmentNameList}    get attachments from email    ${email}
            
            # check if multiple .xlsx file present or not
            ${excelFileList}    Check Attachments With Multiple Excel    ${attachmentNameList}
            # ${filtered_list}
            ${excelAttachmentLength}    Get Length    ${excelFileList}
            # ${excelAttachmentLength}    Convert To Integer    ${excelAttachmentLength}

            ${zipFileList}    Check Attachments With Multiple Zip    ${attachmentNameList}

            ${zipAttachmentLEngth}    Get Length    ${zipFileList}
            
            Log To Console    excel attachments is: ${excelFileList}
            Log    excel attachments is: ${excelFileList}
            Log To Console    zip attachments is: ${zipFileList} 
            Log    zip attachments is: ${zipFileList} 


            # IF    '${excelAttachmentLength}' == '1'
            #     Log To Console    excel attachmentLength is equal to 1
            #     Log    excel attachmentLength is equal to 1
            #     FOR    ${attachment}    IN    @{email}[Attachments]
            #         Log    ${attachment}
            #         # Save Email Attachment
            #         IF  ".xlsx" in "${attachment}[filename]"
            #             Save Email Attachments    ${attachment}    ${attachmentDownloadFolder}${/}InputFolder${/}${senderId}${/}  
            #         END
            #     END
            # ELSE
            #     Log To Console    excel attachmentLength is not equal to 1
            #     Log    excel attachmentLength is not equal to 1
            #     ${mailSubject}    Set Variable    Multiple excel file or no excel file in email found  
            #     # newly added code
            #     send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}
            #     Mark Email As Read    ${email}        
            # END

            # # check zip attachments and downloading
            # IF    '${zipAttachmentLEngth}' =='1'
            #     Log    zip attachmentLength is equal to 1
            #     FOR    ${attachment}    IN    @{email}[Attachments]
            #         Log    ${attachment}
            #         IF  ".zip" in "${attachment}[filename]"
            #             Save Email Attachments    ${attachment}    ${attachmentDownloadFolder}${/}InputFolder${/}${senderId}${/}  
            #         END
            #     END
            # ELSE
            #     ${mailSubject}    Set Variable    Multiple zip files or no zip file found in email
            #     send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}
            #     Mark Email As Read    ${email}
            # END
            



        # new code

            IF    '${excelAttachmentLength}' == '1' and '${zipAttachmentLength}' == '1'
                Log To Console    excel attachment length is 1 and zip attachment length is 1
                Log    excel attachment length is 1 and zip attachment length is 1
            FOR    ${attachment}    IN    @{email}[Attachments]
                Log    downloading: ${attachment}
                # Save Email Attachment
                IF    ".xlsx" in "${attachment}[filename]" or ".zip" in "${attachment}[filename]"
                    Save Email Attachments    ${attachment}    ${attachmentDownloadFolder}${/}InputFolder${/}${senderId}${/}
                END
            END
            ELSE IF    '${zipAttachmentLength}' > '1' and '${excelAttachmentLength}' > '1'
                Log    Multiple zip files and excel files found in email
                ${mailSubject}    Set Variable    Alert: Multiple zip files and excel files found in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}
                Mark Email As Read    ${email}
            
            ELSE IF    '${zipAttachmentLength}' > '1'
                Log    Multiple zip files in the email
                ${mailSubject}    Set Variable    Alert: Multiple zip files in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}
                Mark Email As Read    ${email}
            
            ELSE IF    '${excelAttachmentLength}' > '1'
                Log    Multiple excel files in email
                ${mailSubject}    Set Variable    Alert: Multiple excel files in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}
                Mark Email As Read    ${email}
            
            ELSE IF    '${zipAttachmentLength}' == '0' and '${excelAttachmentLength}' == '0'
                Log    No zip files and excel files found in email
                ${mailSubject}    Set Variable    Alert: No zip files and excel files found in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}
                Mark Email As Read    ${email}
            
            ELSE IF    '${zipAttachmentLength}' == '0'
                Log    No zip files in email
                ${mailSubject}    Set Variable    Alert: No zip files in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}
                Mark Email As Read    ${email}
            
            ELSE IF    '${excelAttachmentLength}' == '0'
                Log    No excel files in email
                ${mailSubject}    Set Variable    Alert: No excel files in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}
                Mark Email As Read    ${email}
            END


        # new code ends here


        move email from inbox to readedMail folder    ${subject} 
        END
        # Move Emails     source_folder=${emailSourceFolder}  target_folder=${emailDestinationFolder}    email_filter=[Subject]=${subject}
    END
    Quit Application
    



email fetching based on subject python
    
    ${filtered_emails}    Fetch Email Based On Subject old    tedst
    Log To Console    email collected
    FOR    ${email}    IN    @{filtered_emails}
        ${attachmentNameList}    get attachments from email    ${email}
        Log To Console    list of attachments is : ${attachmentNameList}
        
    END



move email from inbox to readedMail folder     
    [Arguments]    ${subject}
    Run Keyword And Ignore Error    Move Emails    source_folder=${emailSourceFolder}    target_folder=${emailDestinationFolder}    email_filter=[Subject]=${subject}
       


get attachments from email
    [Arguments]     ${email}



    ${attachments}    Get From Dictionary    ${email}    Attachments
    ${filenames}    Create List
    FOR    ${attachment}    IN    @{attachments}
        ${filename}    Get From Dictionary    ${attachment}    filename
        Append To List    ${filenames}    ${filename}
    END
    Log To Console    ${filenames}
    Log    ${filenames}
    
    RETURN    ${filenames}


send mail if exception occures
    [Arguments]    ${senderId}    ${mailSubject}    ${sendingTime}

    ${body}    Set Variable    Hi team,\n\nThe email you send on ' ${sendingTime} ' is not in proper format.\nPlease send email in correct formal.\nThank you,\nBot
    Send Message    ${senderId}    ${mailSubject}    ${body}



send mail with excel report 23
    [Arguments]    ${senderId}    ${reportpath}    ${subject}    ${body}

     open Application
    
    # ${subject}    Set Variable    Run report
    # Send Message    ${senderId}    ${mailSubject}    ${body}
    Send Email    
    ...    ${senderId}   
    ...    ${subject}    
    ...    ${body}
    ...    attachments=${reportpath}
    Quit Application


send mail if input template not matching
    open Application
    [Arguments]    ${senderId}    ${reportpath}
    ${subject}    Set Variable    Run report
    ${body}    Set Variable    Hi Team, input excel file is not in correct format.
    # Send Message    ${senderId}    ${mailSubject}    ${body}
    Send Email    
    ...    ${senderId}    
    ...    ${subject}    
    ...    ${body}
    ...    attachments=${reportpath}
    Quit Application




# *** Tasks ***
# email collection process
#     # email fetching based on subject
#     email fetching based on subject python
#     # send mail with final report  
#     # attachments=${CURDIR}${/}test.xlsx