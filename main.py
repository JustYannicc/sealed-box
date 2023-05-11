import openpyxl
from qrgen import codegen
from pdfgen import pdfgen

#TODO: Add to the top of the first page Serial number list with the total number of pages and barcodes
#TODO: First row in spreadsheet has title and going down is all the serial numbers and the title has to be put on the first page and on the header for all subsequent pages pages
#TODO: Small sealed label that says box number and sealded and do not unseal
#TODO: Automatically print the pdfs


# Open the input spreadsheet and select the first worksheet
workbook = openpyxl.load_workbook('testinput.xlsx')
worksheet = workbook.active


# Loop over all columns in the worksheet
for column in worksheet.columns:
    # Set up the PDF canvas with A4 size
    name = column[0].value
    if name is not None:
        pdfgen(name, column)
        