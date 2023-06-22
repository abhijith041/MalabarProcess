import os
import win32com.client
import datetime
import random,uuid
import time



outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
 
today = datetime.datetime.today().date()


yesterday = datetime.datetime.today().date() - datetime.timedelta(days=1)


# working code and currently in use, but collect all unread emails from inbox folder.
def fetch_email_based_on_subject(subject_line,duration,start_time):    
    filtered_emails = []
    # duration    =int(duration)
    subject_line = str(subject_line)
    sent_items_folder = namespace.GetDefaultFolder(6)
    filter_criteria = "[UnRead] = True"
    filtered_emails_from_inbox = sent_items_folder.Items.Restrict(filter_criteria)
    
    for email in filtered_emails_from_inbox:
        if email.UnRead and subject_line.lower() in email.Subject.lower():
            filtered_emails.append(email)


    while time.time() - start_time < duration:
        # Sleep for 20 seconds
        time.sleep(20)
        fetch_email_based_on_subject(subject_line,duration,start_time)

    return filtered_emails



def get_email_attachments_name_in_list(email):
    '''collect email attachments name '''
    attachmentsNames = []
    for attachments in email.Attachments:
        # print(attachments.filename)
        filename = attachments.Filename
        attachmentsNames.append(filename)
    return attachmentsNames




# working code, not in use but collect all emails from inbox folder.
def fetch_email_based_on_subject_old2(subject_line):
    subject_line = str(subject_line)
    sent_items_folder = namespace.GetDefaultFolder(6)
    filtered_emails = []
    for email in sent_items_folder.Items:
        if email.UnRead and subject_line.lower() in email.Subject.lower():
            filtered_emails.append(email)
    return filtered_emails


def download_attachments(email, save_location):
    try:        
        os.makedirs(save_location, exist_ok=True)  # Create the save_location directory if it doesn't exist
        for attachment in email.Attachments:
            file_extension = os.path.splitext(attachment.FileName)[1]
            if file_extension.lower() in ['.zip', '.xlsx']:
                attachment.SaveAsFile(os.path.join(save_location, attachment.FileName))
                print(f"Downloaded attachment: {attachment.FileName}")
    
    except Exception as e:
        print(str(e))


# not in use
def fetch_email_based_on_subject_old(subject_line):
    # filter_criteria = "[UnRead] = True"
    filter_criteria = f"[UnRead] = True AND '[Subject]' LIKE '%{subject_line}%'"
    
    sent_items_folder = namespace.GetDefaultFolder(6) 
    filtered_emails = sent_items_folder.Items.Restrict(filter_criteria)
    if len(filtered_emails) > 0:

        return filtered_emails
    else:
        return  None
    

def email_Sending_time(email):
    try:
        sent_time = email.SentOn
        sending_time_formated = sent_time.strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        print(str(e))
        sending_time_formated = None
    return sending_time_formated


def find_sender_email(email):
    '''collect sender emailid from email'''
    try:
        sender_email_address = email.Sender.GetExchangeUser().PrimarySmtpAddress
    except AttributeError:
        sender_email_address = email.SenderEmailAddress
    return sender_email_address


def check_attachments_with_multiple_excel(attachment_list):
    '''check number of excel files in email attachments '''
    
    # abcd = ['abc.xlsx', 'test.xlsx', 'bcd.xls', 'hai.zip','word.docsx']
    excel_list = [item for item in attachment_list if ".xl" in item] 
    # print(excel_list)
    return excel_list


def check_attachments_with_multiple_zip(attachment_list):
    '''check number of  files in email attachments'''
    
    zip_list = [item for item in attachment_list if ".zip" in item]
    # print(zip_list)
    return zip_list



def mark_unread_email_as_read(email):
    try:
        email.UnRead = False
        email.Save()
    except Exception as e:
        print(str(e))
    




# subject_line ='abc'
# fetch_email_based_on_subject(subject_line)









# filter_criteria = "[UnRead] = True AND [SentOn] >= '" + yesterday.strftime("%d/%m/%Y") + "'"
    # filter_criteria = "[UnRead] = True AND [Subject] LIKE '%" + subject_line + "%' AND [SentOn] >= '" + yesterday.strftime("%d/%m/%Y") + "'"
    
    # filter_criteria = "@SQL=" + f"\"urn:schemas:httpmail:subject\" LIKE '%{subject_line}%'"
    
    
    

    # filter_criteria = "@SQL=" + f"\"http://schemas.microsoft.com/mapi/proptag/0x0E07001F\" = True AND \"urn:schemas:httpmail:subject\" LIKE '%{subject_line}%'"
