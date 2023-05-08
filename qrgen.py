import qrcode
import openpyxl

# Load the Excel file
wb = openpyxl.load_workbook('input2.xlsx')

# Select the active worksheet
ws = wb.active

# Create an empty list to store the values
values = []

# Loop through each row and append the value to the list
for row in ws.iter_rows():
    for cell in row:
        if cell.value is not None:
            # Remove the decimal ".0" from cell data
            value = str(cell.value).split(".")[0]
            values.append(value)

# Encode the values as a QR code
qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
qr.add_data('\r'.join(values))
qr.make(fit=True)


# Save the QR code image
img = qr.make_image(fill_color='black', back_color='white')
img.save('qr_code.png')
