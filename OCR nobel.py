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
    cols = ['CUSTOMER NUMBER','INVOICE NUMBER', 'INVOICE DATE', 'P.O. No.', 'TERMS', 'SHIP DATE', 'Total', 'G.S.T./H.S.T.', 'Invoice Total', 'Cash Discount']
    # 'SHIPPING INSTRUCTIONS', 'VIA', 'SHIP POINT'
    cols.extend(['LN #', 'PRODUCT', 'DESCRIPTION', 'ORIG. INV', 'CUSTOMER PROD', 'Superseded Prod', 'Interchange Prod', 'ORDER QTY', 'QTY B.O.', 'QTY SHIP', 'UOM', 'UNIT PRICE', 'UNIT', 'DISC. MULT.', 'AMOUNT (NET)'])
    return cols

def get_file_list(file_dir):
    file_list = []
    for root, _, files in os.walk(file_dir):
        for file in files:
            if file.endswith('.pdf') or file.endswith('.PDF'):
                file_list.append(os.path.join(root, file))
    print(len(file_list))
    return file_list

def crop_text(img, coordinates):
    x1, y1, x2, y2 = coordinates
    img_out = img.crop((x1, y1, x2, y2))
    part_text_page = pytesseract.image_to_string(img_out, lang='eng')
    
    return part_text_page, img_out
    
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
        #d = pytesseract.image_to_data(img, output_type=Output.DICT)
        #df = pd.DataFrame(d)
        #print("d", d)
        #print("df", df)
        #df1 = df[(df.conf != '-1') & (df.text != ' ') & (df.text != '')] # 删除空项
        #print("df1", df1)
        #sorted_blocks = df1.groupby('block_num').first().sort_values('top').index.tolist() # list of blocks num
        #print("sorted_blocks", sorted_blocks) 
        
        #for block in sorted_blocks:
        #    curr = df1[df1['block_num'] == block]
        #    sel = curr[curr.text.str.len() > 3]
        #    char_w = (sel.width / sel.text.str.len()).mean()
        #    prev_par, prev_line, prev_left = 0, 0, 0
        #    text = ''
        #    for ix, ln in curr.iterrows():
        #        if prev_par != ln['par_num']:
        #            text += '\n'
        #            prev_par = ln['par_num']
        #            prev_line = ln['line_num']
        #            prev_left = 0
        #        elif prev_line != ln['line_num']:
        #            text += '\n'
        #            prev_line = ln['line_num']
        #            prev_left = 0
                
        #        added = 0
        #        if ln['left'] / char_w > prev_left + 1:
        #            added = int(ln['left'] / char_w) - prev_left
        #            text += ' ' * added
        #        text += ln['text'] + ' '
        #        prev_left += len(ln['text']) + added + 1
        #    text += '\n'
        #    print(text)
        #    save_text_to_file(text, 'structure_data.txt')

        full_img_text = pytesseract.image_to_string(img, config='--psm 6 --oem 3', lang='eng') + '\n'
        #print("full_img_text", full_img_text)
        save_text_to_file(full_img_text, 'compare.txt')

        if "Please Remit" in full_img_text:
            save_img(img, i, output_image_dir, pdf_filename)
            text_info.append(full_img_text)
        else:
            print(f"Skipped file {file_path} as it it not a e-receipt")
        #print('text_info\n',text_info) 
    return text_info

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

def save_img1(img, i, output_image_dir, pdf_filename):
    # OCR processing to get bounding boxes
    boxes = pytesseract.image_to_boxes(img)

    # Draw bounding boxes on the image
    draw = ImageDraw.Draw(img)
    for b in boxes.splitlines():
        b = b.split(' ')
        
        (x0, y0, x1, y1, text) = (int(b[1]), int(b[2]), int(b[3]), int(b[4]), b[0])
        text=text.encode()

        # Adjusting the y-coordinates correctly
        y0_new = img.height - y0
        y1_new = img.height - y1

        # Drawing the rectangle with the corrected coordinates
        draw.rectangle(((x0, y1_new), (x1, y0_new)), outline="green")  # Note the swap here

        # Drawing text above the top-left corner of the bounding box
        text_position = (x0, y1_new - 12)  # shifting text up by its height + 5 pixels
        draw.text(text_position, text, fill="red")  # 'fill' is the color of the text

    # Construct a unique filename for the image: PDF filename + page number
    image_filename = os.path.join(output_image_dir, f'{pdf_filename}_page_{i+1}.png')
        
    # Save the image
    img.save(image_filename, 'PNG')
    
def save_text_to_file(text, output_file):
    with open(output_file, 'a', encoding='utf-8') as text_file:
        text_file.write(text)

def extract_detailed_invoice_info(ocr_text):
    # Dictionary to hold extracted invoice details
    invoice_details = {}

    # Regular expressions to match specific invoice details
    patterns = {
        'cus_num': r'CUSTOMER NUMBER\s*:\s*(\d+)',
        'invoice_number': r'INVOICE NUMBER:\s*(\S+)',
        'invoice_date': r'INVOICE DATE:\s*(\S+)',
        'po_number': r'P\.O\. NUMBER:\s*(\S+)',
        'terms': r'TERMS:\s*([\S\s]+?)\n',
        'ship_date': r'SHIP DATE:\s*(\S+)',
        'total':r'Total\s+(\d{1,3}(?:,\d{3})*\.\d{2})',
        'gsthst': r'G.S.T/H.S.T.\s*(\S+)',
        'total_amount': r'Invoice Total\s*([\d,]+\.\d{2})',
        'cash_discount': r'Cash Discount\s*([\d,]+\.\d{2})'
    }
    # ['CUSTOMER NUMBER','INVOICE NUMBER', 'INVOICE DATE', 'P.O. No.', 'TERMS', 'SHIP DATE', 'Total', 'G.S.T./H.S.T.', 'Invoice Total', 'Cash Discount',
    # 'LN #', 'PRODUCT', 'DESCRIPTION', 'ORDER QTY', 'QTY B.O.', 'QTY SHIP', 'UOM', 'UNIT PRICE', 'UNIT', 'DISC. MULT.', 'AMOUNT (NET)']

    # Search for matches and extract details
    for key, pattern in patterns.items():
        match = re.search(pattern, ocr_text, re.DOTALL)
        if match:
            invoice_details[key] = match.group(1).strip()  # Use strip to remove leading/trailing whitespace
        else:
            invoice_details[key] = 'N/A'

    return invoice_details

def get_end_index(text_lines, start_index):
    for i in range(start_index,len(text_lines)):
        if 'EFT' in text_lines[i]:
            return i
        elif 'Cash' in text_lines[i]:
            return i
        elif 'Past' in text_lines[i]:
            return i
        elif 'Join' in text_lines[i]:
            return i
        elif 'Total' in text_lines[i]:
            return i
        elif 'Product' in text_lines[i]:
            return i
        elif 'TERMS' in text_lines[i]:
            return i
                    
    return -3

def extract_item_info(text):
    text_lines = text.splitlines() ##
    #print(text_lines)
    start_index = next(i for i, line in enumerate(text_lines) if 'LN#' in line) + 1
    #print(start_index)       
    end_index = get_end_index(text_lines, start_index)
    #print(end_index)

    num_ranges = []
    current_range = []

    for i in range(start_index, end_index):
        line = text_lines[i]
        parts = line.split()

        if len(parts) > 9 and ((parts[-1].replace('.', '', 1).replace(',', '', 1).replace('-', '', 1).replace(':', '', 1).isdigit()) and ((parts[-8].replace('.', '', 1).replace(',', '', 1).replace('-', '', 1).isdigit()) or (parts[-7].replace('.', '', 1).replace(',', '', 1).replace('-', '', 1).isdigit()) or (parts[-6].replace('.', '', 1).replace(',', '', 1).replace('-', '', 1).isdigit()))) and (((not '/' in parts[0]) and (parts[0].isdigit() or parts[0].replace('.', '', 1).isdigit()) and (parts[-1].replace('.', '', 1).replace(',', '', 1).replace('-', '', 1).isdigit()) and (parts[-4].replace('.', '', 1).replace(',', '', 1).isdigit()) and (not any('ORIG' in part for part in parts))) or ((not '/' in parts[0]) and (parts[0].isdigit() or parts[0].replace('.', '', 1).isdigit()) and (parts[-4].replace('.', '', 1).replace(',', '', 1).isdigit()) and (not any('ORIG' in part for part in parts))) or ((not '/' in parts[0]) and (parts[-1].replace('.', '', 1).replace(',', '', 1).replace('-', '', 1).isdigit()) and (parts[-4].replace('.', '', 1).replace(',', '', 1).isdigit()) and (not any('ORIG' in part for part in parts)))):
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

    for i in range(0, len(num_ranges)):
        for k in range(num_ranges[i][0], num_ranges[i][1]):
            line = text_lines[k]
            parts = line.split()
            #print("parts\n", parts)
            if k == num_ranges[i][0]:
                item = {
                    'quantity': parts[0].replace('.', '', 1),
                    'product': ' '.join(parts[1:-8]),  # assuming last 6 parts are not part of description
                    'description': 'N/A',
                    'originv': 'N/A',
                    'customerprod': 'N/A',
                    'interchange': 'N/A',
                    'superseded': 'N/A',
                    'product_number': parts[-8].replace('-', '', 1),
                    'qty_bo': parts[-7].replace('-', '', 1),
                    'qty_ship': parts[-6].replace('-', '', 1),
                    'uom': parts[-5],
                    'price_each': parts[-4],
                    'unit': parts[-3],
                    'unit_mult': parts[-2],
                    'amount': parts[-1].replace('-', '', 1)  # assuming the amount is the last part
                }
                extracted_data.append(item)
            else:
                patterns = {
                    'customerprod': r'Customer Prod:\s*(\S+)',
                    'originv': r'ORIG. INV. #:\s*(\S+)',
                    'superseded': r'Superseded Prod:\s*(\S+)',
                    'interchange': r'Interchange Prod:\s*(\S+)'
                }
                for key, pattern in patterns.items():
                    match = re.search(pattern, line, re.DOTALL)
                    if match:
                        extracted_data[-1][key] = match.group(1).strip()  # Use strip to remove leading/trailing whitespace
                    else:
                        if extracted_data[-1]['description'] == 'N/A':
                            extracted_data[-1]['description'] = line
                        #else:
                        #    extracted_data[-1]['description'] += line
    return extracted_data # [{},{}]

def main():
    file_dir = '/Users/quedonglin/Downloads/Noble Invoices (Special Order)' #"/Users/quedonglin/Library/CloudStorage/OneDrive-UniversityofToronto/02 work/F&S DA/Noble Invoices (Special Order)"
    output_file = "text from imgs no duplication.txt"
    output_image_dir = file_dir + '/img'
    data_file = 'invoice_data no duplication.xlsx'
    
    file_list = get_file_list(file_dir)
    
    photocopy = 0
    er = 0
    aa=0

    cols = get_cols()
    df = pd.DataFrame([cols])
    df.to_excel(data_file, index=False, header = False)
    result = []
    for file_path in file_list:
        ocr_text_list = ocr_pdf(file_path, output_image_dir)
        aa+=1
        save_text_to_file('\n' + str(aa) +'\n' + str(file_path) + ':\n', output_file)
        print(aa)
        if len(ocr_text_list) > 0:
            er+=1
            item_info = extract_detailed_invoice_info(ocr_text_list[-1])
            values_list = list(item_info.values())
            for current_text in ocr_text_list:
                result = []
                result+=values_list
                t = extract_item_info(current_text)
                
                for i in range(0,len(t)):
                    details_list = list(t[i].values())

                    # Load the existing Excel file into a DataFrame
                    existing_df = pd.read_excel(data_file)
                    # Create a DataFrame for the new row of data
                    new_row = pd.DataFrame([values_list+details_list], columns=existing_df.columns)
                    # Concatenate the existing DataFrame with the new row
                    updated_df = pd.concat([existing_df, new_row], ignore_index=True)
                    # Save the updated DataFrame to the same Excel file
                    updated_df.to_excel(data_file, index=False)
                save_text_to_file(current_text, output_file)                       
                # Extract detailed invoice details from OCR text
        else:
            photocopy +=1
        
    print('photocopy:', photocopy)
    print('er:', er)

if __name__ == "__main__":
    main()