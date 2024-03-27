import sys
from os.path import dirname
print(dirname(__file__))
print('\n')
print(sys.path)
# sys.path.append("/Library/Frameworks/Python.framework/Versions/3.10/lib/python3.10/site-packages")
print('\n')
print(sys.prefix)
print('\n')
# input()

import pdfplumber
import os
import re
import pandas as pd

def get_file_list(file_dir):
    file_list = []
    for root, _, files in os.walk(file_dir):
        for file in files:
            if file.endswith('.pdf') or file.endswith('.PDF'):
                file_list.append(os.path.join(root, file))
    print(len(file_list))
    return file_list

def extract_text_from_pdf(file_path):
    text_all = ['\n \n']
    with pdfplumber.open(file_path) as pdf:
        pages = pdf.pages
        for page in pages:
            text = page.extract_text()
            text_all.append(text)
    return ''.join(text_all)

def save_text_to_file(text, output_file):
    with open(output_file, 'a', encoding='utf-8') as text_file:
        text_file.write(text)

def get_cols(n):
    cols = ['INVOICE #', 'DATE', 'P.O. No.', 'TERMS', 'REP.', 'SHIP', 'VIA', 'S.O. NO.', 'ORDERED BY', 'Subtotal',
            'Total']

    # Generate columns for quantities, items, descriptions, U/Ms, prices, and amounts up to 20
    for i in range(1, n):
        cols.extend([
            f'Quantity{i}',
            f'Item{i}',
            f'Description{i}',
            f'U/M{i}',
            f'Price Each{i}',
            f'Amount{i}'
        ])
    #print('get_cols result:', cols, '\n')
    return cols

def extract_invoice_info(text):
    text_lines = text.splitlines()
    result = []

    # Get invoice number
    inv_index = [i for i, x in enumerate(text_lines) if 'INVOICE' in x][1]
    inv = text_lines[inv_index].split()[-1]
    result.append(inv)

    # Get DATE
    inv_index = [i for i, x in enumerate(text_lines) if 'DATE' in x][0]
    inv = text_lines[inv_index].split()[-1]
    result.append(inv)

    # Get position P.O. No. TERMS REP. SHIP VIA S.O. NO. ORDERED BY:
    inv_index = [i for i, x in enumerate(text_lines) if 'P.O. No.' in x][0]
    inv = text_lines[inv_index + 1].split()

    if inv[-1].isdigit():
        inv.append(' ')

    # Regular expression pattern to match dates in the "YYYY-MM-DD" format
    date_pattern = r'\d{4}-\d{2}-\d{2}'
    ship = re.findall(date_pattern, ' '.join(inv))
    ship_index = [i for i, x in enumerate(inv) if ship[0] in x][0]

    po = inv[0]
    result.append(po)
    term = ' '.join(inv[1: ship_index - 1])
    result.append(term)
    rep = inv[ship_index - 1]
    result.append(rep)
    result.append(ship[0])
    via = ' '.join(inv[ship_index + 1: -2])
    result.append(via)
    so = inv[-2]
    result.append(so)

    if len([i for i, x in enumerate(text_lines) if 'ORDERED BY' in x]) > 1:
        ob_index = [i for i, x in enumerate(text_lines) if 'ORDERED BY' in x][1]
        ob = text_lines[ob_index]
    elif len([i for i, x in enumerate(text_lines) if 'ORDERED' in x]) > 1:
        ob_index = [i for i, x in enumerate(text_lines) if 'ORDERED' in x][1]
        ob = text_lines[ob_index]
    elif len([i for i, x in enumerate(text_lines) if 'ORDER' in x]) > 1:
        ob_index = [i for i, x in enumerate(text_lines) if 'ORDER' in x][1]
        ob = text_lines[ob_index]
    else:
        ob = ''
    result.append(ob)

    # get Subtotal
    inv_index = [i for i, x in enumerate(text_lines) if 'Subtotal' in x][0]

    st = text_lines[inv_index].split()[-1]
    result.append(st)

    # get total
    inv_index = [i for i, x in enumerate(text_lines) if 'Total' in x][-1]

    tt = text_lines[inv_index].split()[-1]
    result.append(tt)

    return result


def extract_item_info(text):
    text_lines = text.splitlines() ##
    start_index = text_lines.index('Quantity Item Description U/M Price Each Amount')
    if len([i for i, x in enumerate(text_lines) if 'ORDERED BY' in x]) > 1:
        end_index = [i for i, x in enumerate(text_lines) if 'ORDERED BY' in x][1]
    elif len([i for i, x in enumerate(text_lines) if 'ORDERED' in x]) > 1:
        end_index = [i for i, x in enumerate(text_lines) if 'ORDERED' in x][1]
    elif len([i for i, x in enumerate(text_lines) if 'ORDER' in x]) > 1:
        end_index = [i for i, x in enumerate(text_lines) if 'ORDER' in x][1]
    else:
        end_index = [i for i, x in enumerate(text_lines) if 'Subtotal' in x][0]

    num_ranges = []
    current_range = []

    for i in range(start_index + 1, end_index):
        line = text_lines[i]
        parts = line.split()

        if parts[0].isnumeric() and (parts[-2].replace('.', '', 1).isdigit() or parts[-2].replace('.', '', 1).replace(',', '', 1).isdigit()) and (parts[-1].replace('.', '', 1).isdigit() or parts[-1].replace('.', '', 1).replace(',', '', 1).isdigit()):
            if current_range:
                current_range[1] += 1
                num_ranges.append(current_range)
            current_range = [i, i]
        elif current_range:
            current_range[1] = i

    # Append the last range
    if current_range:
        current_range[1] += 1
        num_ranges.append(current_range)

    extracted_data = []

    quant = ''
    item_number = ''
    item_description = ''
    unit_measurement = ''
    price_each = ''
    amount = ''

    for i in range(0, len(num_ranges)):
        for k in range(num_ranges[i][0], num_ranges[i][1]):
            line = text_lines[k]
            parts = line.split()
            if k == num_ranges[i][0]:
                extracted_data.append(
                    f"{quant} {item_number} {item_description} {unit_measurement} {price_each} {amount}")
                quant = parts[0]
                item_number = parts[1]
                item_description = ' '.join(parts[2:-3])
                unit_measurement = parts[-3]
                price_each = parts[-2]
                amount = parts[-1]
            else:
                item_description = item_description + ' ' + ' '.join(parts)

    extracted_data.append(f"{quant} {item_number} {item_description} {unit_measurement} {price_each} {amount}")
    extracted_data = extracted_data[1:]

    return extracted_data # [[],[]]



def merge_extracted(result, extracted_data):
    for i in range(0, len(extracted_data)):
        items = extracted_data[i].split(' ')

        qua = items[0]
        result.append(qua)
        ite = items[1]
        result.append(ite)
        des = ' '.join(items[2:-3])
        result.append(des)
        um = items[-3]
        result.append(um)
        pe = items[-2]
        result.append(pe)
        am = items[-1]
        result.append(am)
    return result



def modify_result(result, cols):
    while len(result) < len(cols):
        result.append('')

    return result



def main():
    file_dir = "/Users/quedonglin/Downloads/Aurora invoices (Special Order) 2"
    output_file = "all text.txt"
    file_list = get_file_list(file_dir)
    cols = get_cols(20)
    #print(len(cols))

    df = pd.DataFrame([cols])
    df.to_excel('invoice_data all.xlsx', index=False, header = False)

    for file_path in file_list:
        text_all = extract_text_from_pdf(file_path)
        
        print(text_all)
        
        save_text_to_file(text_all, output_file)
        
        if 'INVOICE TO' in text_all:

            header_info = extract_invoice_info(text_all)
            item_info = extract_item_info(text_all)
            #print(header_info)
            #print(item_info)

            result = merge_extracted(header_info, item_info)
            #print(result)
            #print(len(result))

            result = modify_result(result, cols)

            # Load the existing Excel file into a DataFrame
            existing_df = pd.read_excel('invoice_data all.xlsx')
            # Create a DataFrame for the new row of data
            new_row = pd.DataFrame([result], columns=existing_df.columns)
            # Concatenate the existing DataFrame with the new row
            updated_df = pd.concat([existing_df, new_row], ignore_index=True)
            # Save the updated DataFrame to the same Excel file
            updated_df.to_excel('invoice_data all.xlsx', index=False)
        else:
            print ('na')

if __name__ == "__main__":
    main()







