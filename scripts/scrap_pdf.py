# With PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from tqdm import tqdm
import sys


# Usage: python scrap_pdf.py file_name
file_name = sys.argv[1]

reader = PdfReader(file_name)
i = 0

# Convert each page of a pdf file into an independent pdf file, possible text processing
for page in tqdm(reader.pages):
    writer = PdfWriter()
    text = page.extract_text()
    # Search text with regex or split, perform any modifications wanted
    # TODO
    name = 'file_page_' + i
    i += 1
    
    with open('output_pdf/' + name + '.pdf', 'wb') as f:
        writer.write(f)
