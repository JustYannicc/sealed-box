import os
import tempfile
import openpyxl
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
import math
from qrgen import codegen
#TODO: Add to the top of the first page Serial number list with the total number of pages and barcodes
#TODO: First row in spreadsheet has title and going down is all the serial numbers and the title has to be put on the first page and on the header for all subsequent pages pages
#TODO: Every colum creates its own pdf
#TODO: Small sealed label that says box number and sealded and do not unseal


# Define the size of the barcode image and the spacing between them
barcode_width = 40 * mm
barcode_height = 10 * mm
horizontal_spacing = 0.5 * mm
vertical_spacing = 3 * mm

# Open the input spreadsheet and select the first worksheet
workbook = openpyxl.load_workbook('testinput.xlsx')
worksheet = workbook.active


# Loop over all columns in the worksheet
for column in worksheet.columns:
    # Set up the PDF canvas with A4 size
    name = column[0].value
    pdf_canvas = canvas.Canvas(f'{name}.pdf', pagesize=A4)
    pdf_canvas.setLineWidth(.3)
    pdf_canvas.setFont('Helvetica', 8)

    # Define the starting coordinates for the barcode and text
    x, y = 3.75 * mm, 280 * mm

    # Initialize variables to keep track of the total barcode count and the current page number
    barcode_count = 1
    current_pages = 1
    # Loop through each row and column
    count = 0
    for cell in column[1:]:
        if cell.value is not None:
            count += 1
    total_pages = math.ceil(count/100)
    values = []
    
    # Loop through each row in the column and generate a barcode for the cell data
    for cell in column[1:]:
        if cell.value is not None:
            # Get the cell value
            cell_value = cell.value
            cell_value_parts = str(cell_value).split(".")[0] # Remove the decimal ".0" from cell data
            values.append(cell_value_parts)
            
            # Generate the barcode image
            with tempfile.NamedTemporaryFile(delete=False) as tmp:
                options = {
                    'module_width': 0.2,  # Make the barcode wider
                    'module_height': 4.0,  # Make the barcode narrower
                    'quiet_zone': 5.0,  # Leave less margin around the barcode
                    'font_size': 0,
                    'add_checksum': False,
                }
                EAN = Code128(cell_value_parts, writer=ImageWriter())
                EAN.write(tmp, options)
                tmp.flush()
                barcode_image = tmp.name
            
            # Draw the barcode image and cell value text on the PDF canvas
            pdf_canvas.drawImage(barcode_image, x, y, width=barcode_width, height=barcode_height)
            text_width = pdf_canvas.stringWidth(cell_value_parts, 'Helvetica', 8)
            pdf_canvas.drawString(x + (barcode_width - text_width) / 2, y + barcode_height + 0.5 * mm, cell_value_parts)
            
            # Remove the temporary barcode image file
            os.unlink(barcode_image)
            
            # Move to the next barcode position
            x += barcode_width + horizontal_spacing
            if x > 180 * mm:
                x = 3.75 * mm
                y -= barcode_height + vertical_spacing
                
                # If we're at the bottom of the page, start a new page and add the total barcode count to the footer
                if y < 20 * mm:
                    codegen(values)
                    pdf_canvas.drawString(150*mm, 10*mm, f"Page {current_pages} of {total_pages}")
                    pdf_canvas.drawString(10*mm, 10*mm, f"Total Barcodes: {barcode_count}")
                    barcode_count = 0
                    current_pages += 1
                    pdf_canvas.showPage()
                    x, y = 3.75 * mm, 280 * mm
                    pdf_canvas.setFont('Helvetica', 8)
                    values.clear()
                    #FIXME: Generates unlimted QR codes
            
            # Increment the barcode count for each barcode generated
            barcode_count += 1

    # Add the total number of barcodes to the last page footer
    pdf_canvas.drawString(150*mm, 10*mm, f"Page {current_pages} of {total_pages}")
    pdf_canvas.drawString(10*mm, 10*mm, f"Total Barcodes: {barcode_count}")
    codegen(values)
    pdf_canvas.save()

