from reportlab.lib.units import mm
import os
import math
from pathlib import Path
from datetime import datetime
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

# Define the size of the barcode image and the spacing between them
barcode_width = 40 * mm
barcode_height = 10 * mm
horizontal_spacing = 0.5 * mm
vertical_spacing = 3 * mm

def pdfgen(name, column):
    #Setting location of the file
    desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop') 
    datetoday = datetime.today().strftime('%Y-%m-%d')
    pathtosave = f"{desktop}/sealed/sealed_{datetoday}/seal_{name}"
    Path(pathtosave).mkdir(parents=True, exist_ok=True)
    file_path = f"{pathtosave}/{name}.pdf"
    
    pdf_canvas = canvas.Canvas(file_path, pagesize=A4)
    pdf_canvas.setLineWidth(.3)
    pdf_canvas.setFont('Helvetica', 8)

    # Define the starting coordinates for the barcode and text
    x, y = 3.75 * mm, 280 * mm

    # Initialize variables to keep track of the total barcode count and the current page number
    barcode_count = 1
    current_pages = 2
    # Loop through each row and column
    count = 0
    for cell in column[1:]:
        if cell.value is not None:
            count += 1
    total_pages = math.ceil(count/100) + 1
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
                    codegen(values, name, current_pages)
                    pdf_canvas.drawString(180*mm, 10*mm, f"Page {current_pages} of {total_pages}")
                    pdf_canvas.drawString(10*mm, 10*mm, f"Total Barcodes: {barcode_count}")
                    barcode_count = 0
                    current_pages += 1
                    pdf_canvas.showPage()
                    x, y = 3.75 * mm, 280 * mm
                    pdf_canvas.setFont('Helvetica', 8)
                    values.clear()
            
            # Increment the barcode count for each barcode generated
            barcode_count += 1
    
    # Add the total number of barcodes to the last page footer
    barcode_count -= 1
    pdf_canvas.drawString(180*mm, 10*mm, f"Page {current_pages} of {total_pages}")
    pdf_canvas.drawString(10*mm, 10*mm, f"Total Barcodes: {barcode_count}")
    codegen(values, name, current_pages)
    pdf_canvas.save()
    firstpage(name, total_pages, pathtosave)
    
def firstpage(name, numbertotalpages, output_path):
    filepath = f"{output_path}/test.pdf"
    # Initialize the PDF canvas
    c = canvas.Canvas(filepath, pagesize=A4)

    # Set the font size and style for the title
    c.setFont("Helvetica-Bold", 40)

    # Calculate the center position of the page
    page_width, page_height = A4
    title_width = c.stringWidth(f"Sealed: {name} - Date", "Helvetica-Bold", 40)
    title_x = (page_width - title_width) / 2
    title_y = page_height - 50  # Distance from the top of the page

    # Draw the title on the page
    c.drawString(title_x, title_y, f"Sealed: {name} - Date")

    # Set the font size and style for the additional information
    c.setFont("Helvetica", 20)

    # Calculate the position for the barcode count and page count
    count_x = title_x
    count_y = title_y - 50  # Distance below the title

    # Draw the total amount of barcodes
    c.drawString(count_x, count_y, "Total amount of barcodes: 1")

    # Draw the total number of pages
    page_count_y = count_y - 30  # Distance below the barcode count
    c.drawString(count_x, page_count_y, f"Total number of pages: {numbertotalpages}")
    
    c.setFont('Helvetica', 8)
    
    #Footer
    c.drawString(180*mm, 10*mm, f"Page 1 of {numbertotalpages}")
    c.drawString(10*mm, 10*mm, f"Total Barcodes: 1")

    # Save the canvas as a PDF file
    c.save()
    #TODO: Barcode count needs to be fixed
    #TODO: Add the QR codes with the page number on top
    #TODO: Merge PDFs into one
    #TODO: Fix formating