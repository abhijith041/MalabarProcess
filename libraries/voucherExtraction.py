from azure.cognitiveservices.vision.computervision import ComputerVisionClient
from azure.cognitiveservices.vision.computervision.models import OperationStatusCodes
from msrest.authentication import CognitiveServicesCredentials
import time
from pdf2image import convert_from_path
import os
import re
import datetime
# from variables import variables
from configVariables import *



def supporting_doc_checking(extracted_dict):
    '''used to check if supporting document available or not and extract supporting document amount value
    
    extracted_dict         :   extrcated dictionary from function 'extract_voucher_data'.
    
    return                 :   extracted dictionary + supporting document details with keys 'supportingDocFound'
                                and supportingDocFound
    
    '''

    try:
            
        amount_check = extracted_dict['amount']
        amount_check = amount_check.replace(',', '')
        amount_check = int(float(amount_check))
        print(amount_check)
        # print(extracted_result)
        extracted_result_split = extracted_result.split("Authorised Sign")
        print(extracted_result_split)
        # print("hi")
        # after_split = extracted_result_split[1]
        # after_split = after_split.split("\n")
        after_split = extracted_result_split[1].split("\n")
        
        supporting_doc_amount =""
        if len(after_split) > 5 :
            print("supporting doc found")
            # updating dictionary with value supporting document as true.
            voucherData.update({'supportingDocFound':True})

            for item in reversed(after_split):
                item = item.upper()
                if 'RS' in item or '.00' in item or '/-' in item or '/.' in item or ',' in item:
                    item = item.replace('RS', '').replace('.00','').replace('/-','').replace('/.','').replace(',','').strip()
                # item =item.strip()
                if item ==str(amount_check):
                    print("supporting doc match found, value is :",item)
                    supporting_doc_amount =item
                    voucherData.update({'supportingDocAmount':supporting_doc_amount}) 
                    break
    
            if supporting_doc_amount=="":
                after_split= extracted_result_split[0].split("\n")
                count=0
                # for item in reversed(after_split):
                #     if after_split.index(item) < -10:
                #         if '=' in item or '.' in item:
                for reverse_index, item in enumerate(reversed(after_split)):
                    if reverse_index < 5:
                        if '=' in item or '.' in item:
                            # Use regex to extract the integer part
                            match = re.search(r'\d+', item)
                            
                            if match:     
                            # supporting_doc_amount = re.search(r'\d+', item).group()
                            
                                if supporting_doc_amount ==str(amount_check):
                                    count+1
                                    print("supporting doc match found, value is :",supporting_doc_amount)
                                    voucherData.update({'supportingDocAmount':supporting_doc_amount})
                                    break
                    else:
                        voucherData.update({'supportingDocAmount':"Not found"})
                        break            
        else:
            print("supporting doc not found")
            voucherData.update({'supportingDocFound':False})

        return voucherData
    
    except Exception as e:

        print("exception is :" +str(e))

        # print("".join(traceback.format_exception(*sys.exc_info())))
        return "fail"



def voucher_number_Extraction(extracted_result):
    '''voucher number extraction from extracted text'''
    try:
        for vNo in voucher_number_pattern:
                voucher_number_match = re.search(vNo, extracted_result)
    
                if voucher_number_match :        
                    voucher_number_match = voucher_number_match.group(0)
                    voucher_number_match = voucher_number_match.upper()    
                    print("voucher number found:",voucher_number_match) 

                    break
                else:
                    print("no match found for voucher number")
    except Exception as e:
        print(str(e))
        voucher_number_match ='not found'
    return voucher_number_match


def amount_Extraction(extracted_result):
    '''amount extraction from extracted text'''
    try:
        for amount in amount_patterns:
            amount_match = re.search(amount, extracted_result)
            if amount_match:
                amount_found = amount_match.group(1)
                amount_found = amount_found.replace(',', '')
                #amount_found = str(int(amount_match.group(1)))
                print("amount found :",amount_found)
                break
            else:
                print("amount not found")
    except Exception as e:
        print(str(e))
        amount_found ='not found'
    return amount_found


def date_extraction(extracted_result):
    '''date value extraction from extracted text'''
    try:
        for date in date_pattern:
            date_match = re.search(date, extracted_result)
            if date_match:
                date_found = date_match.group(1)
                print("date found :",date_found)
                break
        else:
            print("date not found")
    except Exception as e:
        print(str(e))
        date_found ='not found'
    return date_found


def issued_to_extraction(extracted_result):
    try:
        for item in issued_to_pattern:
            issued_to_match = re.search(item, extracted_result)
            if issued_to_match:
                issued_to_found = issued_to_match.group(1)
                issued_to_found = issued_to_found.upper()
                print("date found :",issued_to_found)
                break
            else:
                print("date not found")
    except Exception as e:
        print(str(e))
        issued_to_found ='not found'
    return issued_to_found


def extract_voucher_data(voucher_pdf,endpoint,subscription_key,vrLegalEntity,legalentity):
    """
    for extracting voucher data from voucher and returning voucher data as a dictionary.

    voucher_pdf         :   input pdf file 
    subscription_key    :   azure ocr engine subscription_key
    endpoint            :   endpoint(url) of the azure ocr engine.

    return              : extracted data dictionary as variable 'voucherData'. Else return value as false.
    """

    

    global voucherData
    voucherData = {}
    # subscription_key = "34729f664d044096bc9e06e162f7a47e"

    # endpoint = 'https://quadanceocrgroup.cognitiveservices.azure.com/'


    computervision_client = ComputerVisionClient(endpoint, CognitiveServicesCredentials(subscription_key))

    # voucher_pdf = r"E:\malabarProcess\Voucher 31\Voucher 31.pdf"
    # voucher_pdf = r"E:\malabarProcess\Petty Samples 1\Petty Samples\Project Petty\MGD INDORE PHEONIX 73-120\73-120\116.pdf"
    # "E:\malabarProcess\Petty Samples 1\Petty Samples\Project Petty\MGD INDORE PHEONIX 73-120\73-120\116.pdf"


    try:    
        # Convert PDF to images
        images = convert_from_path(voucher_pdf)


        global extracted_result
        extracted_result = ""
        # Process each image
        for image in images:
            # Save the image to a temporary file
            image_path = r"temp_image.png"
            image.save(image_path)

            # Call the API
            with open(image_path, "rb") as image_file:
                read_response = computervision_client.read_in_stream(image_file, raw=True)

            # Get the operation location (URL with an ID at the end)
            read_operation_location = read_response.headers["Operation-Location"]
            # Grab the ID from the URL
            operation_id = read_operation_location.split("/")[-1]

            # Retrieve the results            
            while True:
                read_result = computervision_client.get_read_result(operation_id)
                print(read_result)

                time.sleep(1)

                if read_result.status in [OperationStatusCodes.not_started, OperationStatusCodes.running]:
                    continue
                elif read_result.status == OperationStatusCodes.succeeded:
                    for page_result in read_result.analyze_result.read_results:
                        for line in page_result.lines:
                            # print(line.text)
                            extracted_result=extracted_result+"\n"+line.text

                else:
                    print("Text extraction failed.")

                break
                
            # Delete the temporary image file
            os.remove(image_path)
        print(extracted_result)




        voucher_number_match = voucher_number_Extraction(extracted_result)
        amount_found         = amount_Extraction(extracted_result)
        date_found           = date_extraction(extracted_result)
        issued_to_found      = issued_to_extraction(extracted_result)

        # voucher number pattern matching regex
        # voucher_number_pattern = [r'(?<=VR NO:)(\w+-\w+)',r'VR NO:(\S+)', r'\bKRK-[A-Z]{2}\d{3}\b', r'V\.No\s*:\s*(\d+)']

        # for vNo in voucher_number_pattern:
        #     voucher_number_match = re.search(vNo, extracted_result)
 
        #     if voucher_number_match :        
        #         voucher_number_match = voucher_number_match.group(0)
        #         voucher_number_match = voucher_number_match.upper()    
        #         print("voucher number found:",voucher_number_match) 

        #         break
        #     else:
        #         print("no match found for voucher number")

        # amount pattern matching regex 
        # amount_patterns = [
        #     r'Amount\s*\n\s*₹\s*([\d.,]+)',r'₹\s*([\d.,]+)',r'Amount\s+₹\s+(\d+\.\d+)',
        #     r'Amount\s+(\d+\.\d+)\s*₹?',r'₹\s+(\d+\.\d+)\s+In Words',r'Amount\s+(\d+\.\d+)',]

        # for amount in amount_patterns:
        #     amount_match = re.search(amount, extracted_result)
        #     if amount_match:
        #         amount_found = amount_match.group(1)
        #         amount_found = amount_found.replace(',', '')
        #         #amount_found = str(int(amount_match.group(1)))
        #         print("amount found :",amount_found)
        #         break
        #     else:
        #         print("amount not found")
                

        # date pattern matching regex
        # date_pattern = [r'Date:\s*([0-9]{2}-[A-Za-z]{3}-[0-9]{4})',r'Date:\s*([0-9]{2}-[A-Za-z]+-[0-9]{4})',
        #                 r'Date:\s*([0-9]{2}-[A-Za-z]{3,}-[0-9]{4})']

        # for date in date_pattern:
        #     date_match = re.search(date, extracted_result)
        #     if date_match:
        #         date_found = date_match.group(1)
        #         print("date found :",date_found)
        #         break
        #     else:
        #         print("date not found")


        # issued_to_pattern = [r'(?<=Issued to:)\s*(.*?)\s*(?=_)',r'(?<=Issued to:)\s*(.*?)(?=\n)',
        #     r'Issued to:\s*(.*?)(?=\s*Description:)',r'Issued to:\s*(.*?)(?=\s*-\.)',
        #     r'Issued to:\s*(.*?)(?=\s*Date:)',]

        # for item in issued_to_pattern:
        #     issued_to_match = re.search(item, extracted_result)
        #     if issued_to_match:
        #         issued_to_found = issued_to_match.group(1)
        #         issued_to_found = issued_to_found.upper()
        #         print("date found :",issued_to_found)
        #         break
        #     else:
        #         print("date not found")


        #updating dictionary with extracted result 
        if voucher_number_match != "":
            #voucher legal entity updation
            voucherData.update({"VoLegalEntity": vrLegalEntity})

            #legalentity updation
            voucherData.update({"LegalEntity": legalentity})

            voucherData.update({"voucherNo":voucher_number_match})

            # checking if voucher number contains VR NO as prefix in it ie VR No:KRK-220
            if "VR NO" in voucher_number_match:
                voucher_number_match = voucher_number_match.split(":")
                voucher_number = voucher_number_match[1]
                voucherData.update({"voucherNo":voucher_number})
            else:
                voucherData.update({"voucherNo":voucher_number_match})
        else:
            voucherData.update({"voucherNo":"Not found"})

        if amount_found !="":
            voucherData.update({'amount':amount_found})
        else:
            voucherData.update({'amount':"Not found"})
        
        if issued_to_found != "":
            voucherData.update({'issuedTo':issued_to_found})
        else:
            voucherData.update({'issuedTo':"Not found"})
        
        if date_found != "":
            voucherData.update({'date':date_found})
        else:
            voucherData.update({'date':"Not found"})
        
        # return voucher data dictionary
        # print(voucherData)

        # calling supporitng document extraction function

        # voucherData = supporting_doc_checking(voucherData)
        supportingDocResult = supporting_doc_checking(voucherData)

        if supportingDocResult =='fail':
            
            print("supporting doc extraction incomplete")
            voucherData.update({'supportingDocFound':False})
            voucherData.update({'supportingDocAmount':"Not found"})
            print(voucherData)
        else:
            print(voucherData)
            # voucherData.update({'supportingDocFound':False})
            # voucherData.update({'supportingDocAmount':"Not found"})
        return  voucherData
    
    except Exception as e:

        print("exception is :" +str(e))

        # print("".join(traceback.format_exception(*sys.exc_info())))
        return 'error'





# def convert_date_to_custom_format(timestamp):
#     date_obj = datetime.datetime.strptime(timestamp, "Timestamp('%Y-%m-%d %H:%M:%S')")
#     formatted_date = date_obj.strftime('%d-%b-%Y')
#     print(formatted_date)
#     return formatted_date


def convert_date_to_custom_format(timestamp):
    date_obj = datetime.datetime.strptime(timestamp.strftime('%Y-%m-%d %H:%M:%S'), "%Y-%m-%d %H:%M:%S")
    formatted_date = date_obj.strftime('%d-%b-%Y')
    return formatted_date


def convert_date_to_custom_format1(timestamp):
    date_obj = datetime.datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
    formatted_date = date_obj.strftime('%d-%b-%Y')
    print(formatted_date)
    return formatted_date



# abc = '2023-01-07 16:23:39'
# convert_date_to_custom_format(abc)
# Example usage
# timestamp = "Timestamp('2023-03-23 08:30:20')"
# custom_format = convert_date_to_custom_format(timestamp)
# print(forma)






