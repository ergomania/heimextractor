import sys
import re
import csv
from datetime import datetime
import PyPDF2

# Get the PDF file name from the command-line arguments
pdf_file_name = 'test.pdf'
try:
    pdf_file_name = sys.argv[1]
except IndexError:
    print('Looking for test.pdf in the current directory')

positions = [['position', 'sku', 'name', 'qty', 'price']]

try:
    # Open the PDF file in read-binary mode
    with open(pdf_file_name, 'rb') as pdf_file:
        # Create a PDF reader object
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        # Get the number of pages in the PDF file
        num_pages = len(pdf_reader.pages)

        # Loop through each page and extract the text
        for page_num in range(num_pages):
            # Get the page object
            page_obj = pdf_reader.pages[page_num]

            # Extract the text from the page
            page_text = page_obj.extract_text()

            # Print the text
            page_lines = page_text.split('\n')

            for line in page_lines:
                matches = re.search(r'^([0-9]*) ([0-9]*) (.*(?= Stk)) Stk ([^ ]*)', line, re.MULTILINE)
                if matches:
                    position = matches.group(1)
                    sku = matches.group(2)
                    name_with_qty = matches.group(3).split(' ')
                    qty = name_with_qty.pop()
                    name = ' '.join(name_with_qty)
                    price = matches.group(4)
                    price = price.replace('’', '')

                    positions.append([
                        position,
                        sku,
                        name,
                        qty,
                        price
                    ])

                    print('Found position {} with sku {} and name {}'.format(
                        position,
                        sku,
                        name
                    ))

    print('Found {} positions'.format(len(positions)))

    now = datetime.now()
    with open('{}_heim_extraction.csv'.format(now.strftime("%Y%m%d")), 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        for position in positions:
            writer.writerow(position)

except FileNotFoundError:
    print('File not found.')
