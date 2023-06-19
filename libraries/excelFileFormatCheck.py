import openpyxl
import configVariables


def print_excel_headers(file_path):
    '''Print the column headers in the Excel file'''

    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    # Assuming headers are located in row 8
    header_row = 8

    # Get all the column headers in row 8
    column_headers = [cell.value for cell in sheet[header_row]]

    # Print the column headers
    for header in column_headers:
        print(header)

    return column_headers


def check_excel_headers(file_path):
    '''Check Excel file format before processing'''

    headers_to_check    = configVariables.headers_to_check
    # headers_to_check = ['voucher_number', 'date', 'expense_category', 'type', 'issued_to', 'description', 'debit', 'credit', 'closing_balance']
    try:
        columns_in_file = print_excel_headers(file_path)
        columns_in_file = [column.lower() if column is not None else None for column in columns_in_file]

        print(columns_in_file)

        # Iterate over each column header in the file
        for header in headers_to_check:
            if header.lower() not in columns_in_file:
                print('No match found')
                return 'no_match'

        print('All headers matched successfully')
        return 'match'
    
    except Exception as e:
        print(f'An error occurred while checking the headers: {str(e)}')
        return 'no_match'

# Provide the file path
#file_path = 'C:/Users/Q0037/Documents/Malabar/VoucherVerificationProcess/InputFolder/nibiyamonas@quadance.com/input.xlsx'

# Check the Excel file headers
#check_excel_headers(file_path)
