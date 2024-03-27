import pytesseract
from pytesseract import Output
from PIL import Image, ImageDraw
import pdf2image
import re
import os
import pandas as pd

# print('\n\n', os.environ['PATH'])
# os.environ['PATH'] += os.pathsep + '/opt/homebrew/bin/tesseract'

def get_cols():
    cols = ['file name', 'INVOICE #', 'DATE', 'P.O. No.', 'TERMS', 'REP.', 'SHIP', 'VIA', 'S.O. NO.', 'ORDERED BY', 'Subtotal', 'Total']
    cols.extend(['Quantity', 'Item', 'Description', 'U/M', 'Price Each', 'Amount'])
    return cols

def get_file_list(file_dir):
    file_list = []
    for root, _, files in os.walk(file_dir):
        for file in files:
            if file.endswith('.pdf') or file.endswith('.PDF'):
                file_list.append(os.path.join(root, file))
    print(len(file_list))
    return file_list
    
def ocr_pdf(file_path, output_image_dir):
   # Ensure the output directory exists
    if not os.path.exists(output_image_dir):
        os.makedirs(output_image_dir)
    # Extract the filename without extension from the file_path
    pdf_filename = os.path.splitext(os.path.basename(file_path))[0]
    # Convert PDF to list of images
    images = pdf2image.convert_from_path(file_path, fmt='png', dpi=200)
    
    # Perform OCR on each image
    text_info = []
            
    for i, img in enumerate(images):

        full_img_text = pytesseract.image_to_string(img, config='--psm 6 --oem 3', lang='eng') + '\n'
        #print("full_img_text", full_img_text)
        save_text_to_file(full_img_text, 'compare.txt')

        if ("INVOICE") in full_img_text:
            head = crop_head(img)
            items = crop_inv(img)
            #print(items)
            save_img(img, i, output_image_dir, pdf_filename)
            text_info.append(full_img_text)
            return text_info, head, items, pdf_filename
        else:
            print(f"Skipped file {file_path} as it it not a e-receipt")
        #print('text_info\n',text_info) 
            return [[],'','','']

def save_img(img, i, output_image_dir, pdf_filename):
    # Get word-level bounding box information
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    #print(data, '\n')
    n_boxes = len(data['level'])
    draw = ImageDraw.Draw(img)
    for i in range(n_boxes):
        if data['level'][i] >= 4:  # Level 4 corresponds to word level
            (x, y, w, h, text) = (data['left'][i], data['top'][i], data['width'][i], data['height'][i], data['text'][i])
            text=text.encode()
            #print(i, text, '\n')
            draw.rectangle(((x, y), (x + w, y + h)), outline='green')
            draw.text((x, y-12), text, fill='red')  # Adjust y position to place text above the bounding box
    img.save(os.path.join(output_image_dir, f'{pdf_filename}_{i}.png'))


def crop_head(img):
    x1, y1, x2, y2 = (60,780,1700,845)
    img_out = img.crop((x1, y1, x2, y2))
    part_text_page = pytesseract.image_to_string(img_out, config='--psm 6 --oem 3', lang='eng') + '\n'
    part_text_page = part_text_page.replace('|', '').replace('\n', '')
    #print("crop_head\n", part_text_page)
    return part_text_page

def crop_inv(img):
    x1, y1, x2, y2 = (30,900,1700,1750)
    img_out = img.crop((x1, y1, x2, y2))
    part_text_page = pytesseract.image_to_string(img_out, config='--psm 6 --oem 3', lang='eng') + '\n'
    #save_img(img_out, 0, '/Users/quedonglin/Downloads/Aurora invoices (Special Order) 2/t/img', img)
    #print("crop_inv\n", part_text_page)
    return part_text_page

    
def save_text_to_file(text, output_file):
    with open(output_file, 'a', encoding='utf-8') as text_file:
        text_file.write(text)

def extract_detailed_invoice_info(text): # ocr_text_list [[full],[inv],[head],[file name]]
    text_lines = text[0][0].splitlines()
    result = []
    result.append(text[-1])
    
    #head = text[2].splitlines()

    # Get invoice number
    inv_index = [i for i, x in enumerate(text_lines) if 'INVOICE' in x][0]
    inv = text_lines[inv_index].split()[-1]
    result.append(inv)

    # Get DATE
    inv_index = [i for i, x in enumerate(text_lines) if 'DATE' in x][0]
    inv = text_lines[inv_index].split()[-1]
    result.append(inv)

    #('result', result)

    head_info = text[1]
    inv = head_info.split()
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
    
    #print('result\n', result)

    if not inv[-1].isdigit():
        ob = inv[-1]
    elif len([i for i, x in enumerate(text_lines) if 'ORDERED BY' in x]) > 1:
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

    st = text_lines[inv_index].split()[-1].replace('$', '')
    result.append(st)

    # get total
    inv_index = [i for i, x in enumerate(text_lines) if 'Total' in x][-1]

    tt = text_lines[inv_index].split()[-1].replace('$', '')
    result.append(tt)

    return result

def get_end_index(text_lines, start_index):
    for i in range(start_index,len(text_lines)):
        #print('i\n', text_lines)
        if 'ORDERED BY' in text_lines[i]:
            return i
        elif 'ORDERED' in text_lines[i]:
            return i
        elif 'ORDER' in text_lines[i]:
            return i
    return len(text_lines) - 1

def extract_item_info(text, filename):
    text_lines = text.splitlines() #
    text_lines = [line for line in text_lines if line.strip()]
    #print(text_lines, 'text_lines')
    start_index = 0
    end_index = get_end_index(text_lines, start_index)

    #print('start_index, end_index', start_index, end_index)
    
    num_ranges = []
    current_range = []

    for i in range(start_index, end_index):
        line = text_lines[i]
        parts = line.split()

        if len(parts) > 3 and ((len(parts) > 3 and parts[0].isnumeric() and (parts[-2].replace('.', '', 1).replace(',', '').replace(':', '', 1).replace('$', '', 1).isdigit()) and (parts[-1].replace('.', '', 1).replace(',', '').replace(':', '', 1).replace('$', '', 1).isdigit())) or (len(parts) > 3 and parts[-2].replace('.', '', 1).replace(',', '').replace(':', '', 1).replace('$', '', 1).isdigit() and (parts[-1].replace('.', '', 1).replace(':', '', 1).replace(',', '').replace('$', '', 1).isdigit())) or (parts[0].replace('.', '', 1).replace(',', '').isdigit() and (parts[-1].replace('.', '', 1).replace(':', '', 1).replace(',', '').replace('$', '', 1).isdigit()))):
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
    
    if len(num_ranges) == 0:
        print('fail to save', filename, '.pdf\n')
        save_text_to_file(filename + '\n' + text + '\n\n', "file failed to save.txt")
        return [{'quant': 'N/A', 'prod': 'N/A', 'item_description': 'N/A', 'unit_measurement': 'N/A', 'price_each': 'N/A', 'amount': 'N/A'}]
        
    for i in range(0, len(num_ranges)):
        for k in range(num_ranges[i][0], num_ranges[i][1]):
            line = text_lines[k]
            parts = line.split()
            # print('parts\n', parts)
            
            if k == num_ranges[i][0]:
                quantity = parts[0].replace('.', '', 1) if parts[0].replace('.', '', 1).isdigit() else '999'
                num = parts[1] if parts[0].replace('.', '', 1).isdigit() else parts[0]
                unit = parts[-3] if len(parts[-3]) < 4 else 'ea'

                if parts[0].replace('.', '', 1).isdigit():
                    des_end_index = -3 if len(parts[-3]) < 4 else -2
                    des = ' '.join(parts[2:des_end_index])
                else:
                    des_end_index = -3 if len(parts[-3]) < 4 else -2
                    des = ' '.join(parts[1:des_end_index])
                
                item = {
                    'quantity': quantity,
                    'item_number': num,
                    'item_description': des,  # assuming last 6 parts are not part of description
                    'unit_measurement': unit,
                    'price_each': parts[-2],
                    'amount': parts[-1]
                }
                extracted_data.append(item)
            else:
                extracted_data[-1]['item_description'] += (' ' + ' '.join(parts))
                # item_description = item_description + ' ' + ' '.join(parts)

    return extracted_data # [{},{}]


def main():
    file_dir = '/Users/quedonglin/Downloads/Aurora invoices (Special Order) 2'
    #'/Users/quedonglin/Downloads/Noble Invoices (Special Order)' #"/Users/quedonglin/Library/CloudStorage/OneDrive-UniversityofToronto/02 work/F&S DA/Noble Invoices (Special Order)"
    output_file = "all text1.txt"
    output_image_dir = file_dir + '/img'
    data_file = 'inv data no duplication1.xlsx'
    
    file_list = get_file_list(file_dir)
    
    photocopy = 0
    er = 0
    aa=0

    cols = get_cols()
    df = pd.DataFrame([cols])
    df.to_excel(data_file, index=False, header = False)
    result = []
    for file_path in file_list:
        all = ocr_pdf(file_path, output_image_dir)
        ocr_text_list = all[0]
        item_info = all[2]
        
        aa+=1
        save_text_to_file('\n' + str(aa) +'\n' + str(file_path) + ':\n', output_file)
        print(aa)
        if len(ocr_text_list) > 0:
            er+=1
            list_text = ocr_pdf(file_path, output_image_dir)
            print('list_text\n',list_text)
            info_lst = extract_detailed_invoice_info(list_text) #['123934', '2023-07-27', 'SO#706961', '2% 10 NET 30', 'Jc', '2023-07-14', 'OUR TRUCK', '120610', ' ', '$386.15', '$436.35']
            #print('info_lst\n',info_lst)

            #for current_text in item_info:
            result = []
            result+=info_lst
            t = extract_item_info(item_info, list_text[3])
            #print('t\n',t)
            for i in range(0,len(t)):
                details_list = list(t[i].values())
                    # Load the existing Excel file into a DataFrame
                existing_df = pd.read_excel(data_file)
                    # Create a DataFrame for the new row of data
                new_row = pd.DataFrame([info_lst+details_list], columns=existing_df.columns)
                    # Concatenate the existing DataFrame with the new row
                updated_df = pd.concat([existing_df, new_row], ignore_index=True)
                    # Save the updated DataFrame to the same Excel file
                updated_df.to_excel(data_file, index=False)
            save_text_to_file(item_info, output_file)                       
                # Extract detailed invoice details from OCR text
        else:
            photocopy +=1
        
    print('photocopy:', photocopy)
    print('er:', er)

if __name__ == "__main__":
    main()