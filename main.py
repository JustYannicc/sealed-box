import os
import tempfile

import openpyxl
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
import math
import qrgen
#TODO: Add to the top of the first page Serial number list with the total number of pages and barcodes
#TODO: Make a QR code that gives a list of all devices
#TODO: First row in spreadsheet has title and going down is all the serial numbers and the title has to be put on the first page and on the header for all subsequent pages pages
#TODO: Every colum creates its own pdf


# Define the size of the barcode image and the spacing between them
barcode_width = 40 * mm
barcode_height = 10 * mm
horizontal_spacing = 0.5 * mm
vertical_spacing = 3 * mm

# Open the input spreadsheet and select the first worksheet
workbook = openpyxl.load_workbook('input.xlsx')
worksheet = workbook.active
            
# Set up the PDF canvas with A4 size
pdf_canvas = canvas.Canvas('output.pdf', pagesize=A4)
pdf_canvas.setLineWidth(.3)
pdf_canvas.setFont('Helvetica', 8)

# Define the starting coordinates for the barcode and text
x, y = 3.75 * mm, 280 * mm

# Initialize variables to keep track of the total barcode count and the current page number
barcode_count = 1
current_pages = 1
# Loop through each row and column
count = 0
for row in worksheet.iter_rows():
    for cell in row:
        if cell.value is not None:
            count += 1
total_pages = math.ceil(count/100)

# Loop through each row in the worksheet and generate a barcode for the cell data
for row in worksheet.iter_rows(values_only=True):
    # Get the cell value
    cell_value = str(row[0]).split(".")[0]  # Remove the decimal ".0" from cell data
    
    # Generate the barcode image
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        options = {
            'module_width': 0.2,  # Make the barcode wider
            'module_height': 4.0,  # Make the barcode narrower
            'quiet_zone': 5.0,  # Leave less margin around the barcode
            'font_size': 0,
            'add_checksum': False,
        }
        EAN = Code128(cell_value, writer=ImageWriter())
        EAN.write(tmp, options)
        tmp.flush()
        barcode_image = tmp.name
    
    # Draw the barcode image and cell value text on the PDF canvas
    pdf_canvas.drawImage(barcode_image, x, y, width=barcode_width, height=barcode_height)
    text_width = pdf_canvas.stringWidth(cell_value, 'Helvetica', 8)
    pdf_canvas.drawString(x + (barcode_width - text_width) / 2, y + barcode_height + 0.5 * mm, cell_value)
    
    # Remove the temporary barcode image file
    os.unlink(barcode_image)
    
    # Move to the next barcode position
    x += barcode_width + horizontal_spacing
    if x > 180 * mm:
        x = 3.75 * mm
        y -= barcode_height + vertical_spacing
        
        # If we're at the bottom of the page, start a new page and add the total barcode count to the footer
        if y < 20 * mm:
            pdf_canvas.drawString(150*mm, 10*mm, f"Page {current_pages} of {total_pages}")
            pdf_canvas.drawString(10*mm, 10*mm, f"Total Barcodes: {barcode_count}")
            barcode_count = 0
            current_pages += 1
            pdf_canvas.showPage()
            x, y = 3.75 * mm, 280 * mm
            pdf_canvas.setFont('Helvetica', 8)
    
    # Increment the barcode count for each barcode generated
    barcode_count += 1

# Add the total number of barcodes to the last page footer
pdf_canvas.drawString(150*mm, 10*mm, f"Page {current_pages} of {total_pages}")
pdf_canvas.drawString(10*mm, 10*mm, f"Total Barcodes: {barcode_count}")
pdf_canvas.save()

