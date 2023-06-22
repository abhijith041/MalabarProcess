*** Settings ***
Library     emailFetchingPython
Library    Collections
Library    RPA.Outlook.Application
Variables    variables.py
Library       MyLibrary
Library    RPA.FileSystem
Library    RPA.Windows
Library    DateTime
Library    String



*** Keywords ***
Email Fetching Based On Subject
    

    # Define the duration in seconds (10 minutes = 600 seconds) -email polling duration
    ${duration}    Set Variable    ${polling_duration}
    
    # collecting process starting time for email polling
    ${startTimeInSeconds}    process run starting time convert into seconds
    
    FOR    ${subject}    IN    @{EMAIL_SUBJECT}
    # Fetch Email Based On Subject New    subject_line

        ${filteredEmails}    Fetch Email Based On Subject    ${subject}    ${duration}     ${startTimeInSeconds}   
        ${emailLength}    Get Length    ${filteredEmails}
        
        Log To Console    \nSearching for email with subject : ${subject}
        Log    Searching for email with subject : ${subject}

        Log To Console    \nTotal email with given subject is: ${emailLength}
        Log    Total email with given subject is: ${emailLength}

        ${attachmentDownloadFolder}     get_init_details

        FOR    ${email}    IN    @{filteredEmails}
        # iterating throgh each mail 
            ${senderId}    Find Sender Email    ${email}
            # ${senderId}    Set Variable    abhijith.p@quadance.com
            Log To Console    Email subject from collected mail is: ${email.Subject}
            Log To Console    mail seder id : ${senderId}   
            ${sendingTime}    email Sending time    ${email}
            Log To Console    '${sendingTime}'
            # ${sendingTime}    Evaluate    ${sendingTime.strip()}
            # Log To Console    '${sendingTime}'

            ${attachmentCount}    Get Email Attachment Count    ${email}
            IF    ${attachmentCount} > 0
                FOR    ${index}    IN RANGE    ${attachmentCount}

                    # disable this if you are running code using .bat file
                    ${attachmentNames}    Get Email Attachments Name    ${email}    ${index}
                    
                    # disable this if you are running code using IDE
                    # ${attachmentNames}    Get Email Attachments Name In List    ${email}
                    
                    log    ${attachmentNames}
                    # Append To List    ${attachmentNames}    ${attachment_name}
                END
                Log To Console    ${attachmentNames}
            END
            
            IF    ${attachmentCount} == 0
                
                Log    the email doesnot containing any attachments
                ${mailSubject}    Set Variable    Alert: The email doesn't containing any attachments
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}    ${email.Subject}
                # Mark Email As Read    ${email}
                Mark Unread Email As Read    ${email}
            
                Continue For Loop
            END

            # check if multiple .xlsx file present or not
            ${excelFileList}    Check Attachments With Multiple Excel    ${attachmentNames}
            # ${filtered_list}
            ${excelAttachmentLength}    Get Length    ${excelFileList}
            # ${excelAttachmentLength}    Convert To Integer    ${excelAttachmentLength}

            ${zipFileList}    Check Attachments With Multiple Zip    ${attachmentNames}

            ${zipAttachmentLEngth}    Get Length    ${zipFileList}

            Log To Console    excel attachments is: ${excelFileList}
            Log    excel attachments is: ${excelFileList}
            Log To Console    zip attachments is: ${zipFileList} 
            Log    zip attachments is: ${zipFileList}    
            
            # concatenating senderId and sendingTime to variable ${senderId_sendingTime} 
            # ${voucher_amount}    Replace String    ${voucherData['amount']}    .00    ${EMPTY}
            ${sendingTime}    Replace String    ${sendingTime}    :    -  
            ${senderId_sendingTime}    Catenate    ${senderId}_   ${sendingTime}

            ${attachmentDownloadPath}    Set Variable    ${attachmentDownloadFolder}${/}InputFolder${/}${senderId_sendingTime}${/}
            IF    '${excelAttachmentLength}' == '1' and '${zipAttachmentLength}' == '1'
                Log To Console    excel attachment length is 1 and zip attachment length is 1
                Log    excel attachment length is 1 and zip attachment length is 1
                
                # log to console    downloading attachments: ${attachmentNames}
                FOR    ${attachment}    IN    @{attachmentNames}
                    Log To Console    downloading: ${attachment}
                    Log    downloading: ${attachment}

                    # downloading files using python code
                    Download Attachments    ${email}    ${attachmentDownloadPath}

                    # IF    ".xlsx" in "${attachment.FileName}" or ".zip" in "${attachment.FileName}"
                    #     Download Attachments    ${attachment}    ${attachmentDownloadPath}
                    # # Save Email Attachment
                    # # IF    ".xlsx" in "${attachment}[filename]" or ".zip" in "${attachment}[filename]"
                    # #     Save Email Attachments    ${attachment}    ${attachmentDownloadFolder}${/}InputFolder${/}${senderId}${/}
                    # END
                END
                Mark Unread Email As Read    ${email}



            ELSE IF    '${zipAttachmentLength}' > '1' and '${excelAttachmentLength}' > '1'
                Log    Multiple zip files and excel files found in email
                ${mailSubject}    Set Variable    Alert: Multiple zip files and excel files found in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}    ${email.Subject}
                Mark Unread Email As Read    ${email}
            
            ELSE IF    '${zipAttachmentLength}' > '1'
                Log    Multiple zip files in the email
                ${mailSubject}    Set Variable    Alert: Multiple zip files in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}    ${email.Subject}
                # Mark Email As Read    ${email}
                Mark Unread Email As Read    ${email}
            
            ELSE IF    '${excelAttachmentLength}' > '1'
                Log    Multiple excel files in email
                ${mailSubject}    Set Variable    Alert: Multiple excel files in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}    ${email.Subject}
                # Mark Email As Read    ${email}
                Mark Unread Email As Read    ${email}
            
            ELSE IF    '${zipAttachmentLength}' == '0' and '${excelAttachmentLength}' == '0'
                Log    No zip files and excel files found in email
                ${mailSubject}    Set Variable    Alert: No zip files and excel files found in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}    ${email.Subject}
                # Mark Email As Read    ${email}
                Mark Unread Email As Read    ${email}
            
            ELSE IF    '${zipAttachmentLength}' == '0'
                Log    No zip files in email
                ${mailSubject}    Set Variable    Alert: No zip files in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}    ${email.Subject}
                # Mark Email As Read    ${email}
                Mark Unread Email As Read    ${email}
            
            ELSE IF    '${excelAttachmentLength}' == '0'
                Log    No excel files in email
                ${mailSubject}    Set Variable    Alert: No excel files in email
                send mail if exception occures    ${senderId}    ${mailSubject}    ${sendingTime}    ${email.Subject}
                # Mark Email As Read    ${email}
                Mark Unread Email As Read    ${email}
            END

        
        END  

    Quit Application

    END


process run starting time convert into seconds
    ${startingTime}    Get Time
    Log To Console    \nProcess starting time is :${startingTime}
    ${startingTimeSeconds}    Run Keyword And Return Status    Should Match Regexp    ${startingTime}    \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}
    Run Keyword If    ${startingTimeSeconds}    Run Keyword And Continue On Failure    Fail    Invalid starting time format

    ${startTimeParts}    Split String    ${startingTime}    separator=.
    ${startTime}    Set Variable    ${startTimeParts}[0]

    ${startTimeObj}    Evaluate    datetime.datetime.strptime($startTime, "%Y-%m-%d %H:%M:%S")
    ${startTimeInSeconds}    Evaluate    $startTimeObj.timestamp()

    Log To Console   Starting Time in Seconds: ${startTimeInSeconds}
    # Set Global Variable    ${startTimeInSeconds}
    RETURN    ${startTimeInSeconds}


Get Email Attachments Name
    [Arguments]    ${email}    ${index}
    ${attachmentNames}    Create List
    # Log    ${email.Attachments}
    # Log To Console   ${email.Attachments}
    ${emailAttachments}    Set Variable    ${email.Attachments}
    # Log    ${emailAttachments}

    FOR    ${attachment}    IN    @{email.Attachments}
        Append To List    ${attachmentNames}    ${attachment.FileName}
    END
    ${attachment_name}    Set Variable    ${attachmentNames[${index}]}
    # Log To Console    ${attachmentNames}
    [Return]    ${attachmentNames}


Get Email Attachment Count
    [Arguments]    ${email}
    ${attachmentCount}    Set Variable    ${email.Attachments.Count}
    [Return]    ${attachmentCount}

send mail if exception occures
    [Arguments]    ${senderId}    ${mailSubject}    ${sendingTime}    ${email.Subject}
    Open Application
    
    ${body}    Set Variable    Hi team,\n\nThe email you send on '${sendingTime}' with subject line '${email.Subject}' is not in proper format.\nPlease send email in correct format.\nThank you,\nBot
    # Send Message    ${senderId}    ${mailSubject}    ${body}
    Send Email    ${senderId}    ${mailSubject}    ${body}

Get Attachment Filename
    [Arguments]    ${attachment}
    ${filename}    Set Variable    ${attachment.FileName}
    [Return]    ${filename}



send mail with excel report
    [Arguments]    ${senderId}    ${reportpath}    ${subject}    ${body}

    open Application
    
    # ${subject}    Set Variable    Run report
    # Send Message    ${senderId}    ${mailSubject}    ${body}
    Send Email    
    ...    ${senderId}   
    ...    ${subject}    
    ...    ${body}
    ...    attachments=${reportpath}
    # Quit Application




# *** Tasks ***

# run this Tasks    
#     Email Fetching Based On Subject
#     Quit Application