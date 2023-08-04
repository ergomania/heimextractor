import sys
import re
import csv
from datetime import datetime
from tkinter import filedialog
import PyPDF2

positions = [['position', 'sku', 'name', 'qty', 'price']]

def main():
    # Select pdf files
    pdf_file_names = filedialog.askopenfilenames(filetypes=[('PDF', '*.pdf')])
    if (len(pdf_file_names) == 0):
        print('No file selected.')
        sys.exit()

    # Loop through each PDF file
    for pdf_file_name in pdf_file_names:
        extract_positions_from_file(pdf_file_name)

    # Write to csv file
    write_to_csv()


def extract_positions_from_file(pdf_file_name):
    print('Extracting from {}'.format(pdf_file_name))

    try:
        # Open the PDF file in read-binary mode
        with open(pdf_file_name, 'rb') as pdf_file:
            positions_in_file = []

            # Read the PDF file as a text file
            lines = read_lines_from_file(pdf_file)

            # Loop through each line and extract the position
            for line in lines:
                if len(line) == 0:
                    continue

                positions_in_file.append(extract_positions_from_line(line))

            print('Found {} positions'.format(len(positions_in_file)))

    except FileNotFoundError:
        print('File not found.')

    positions.extend(positions_in_file)


def read_lines_from_file(filename):
    lines = []

    # Create a PDF reader object
    pdf_reader = PyPDF2.PdfReader(filename)

    # Get the number of pages in the PDF file
    num_pages = len(pdf_reader.pages)

    # Loop through each page and extract the text
    for page_num in range(num_pages):
        # Get the page object
        page_obj = pdf_reader.pages[page_num]

        # Extract the text from the page
        page_text = page_obj.extract_text()

        # Print the text
        lines += page_text.split('\n')

    return lines


def extract_positions_from_line(line):
    print('Extracting from line {}'.format(line))
    matches = re.search(r'^([0-9]*) ([0-9]*) (.*(?= Stk)) Stk ([^ ]*)', line, re.MULTILINE)
    if not matches:
        return []

    position = matches.group(1)
    sku = matches.group(2)
    name_with_qty = matches.group(3).split(' ')
    qty = name_with_qty.pop()
    name = ' '.join(name_with_qty)
    price = matches.group(4)
    price = price.replace('’', '')

    print('Found position {} with sku {} and name {}'.format(
        position,
        sku,
        name
    ))

    return [position, sku, name, qty, price]


def write_to_csv():
    timestamp = datetime.now().strftime("%Y%m%d")
    with open('{}_heim_extraction.csv'.format(timestamp), 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        for position in positions:
            writer.writerow(position)


if __name__ == "__main__":
    main()
