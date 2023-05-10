import qrcode
import os

def codegen(values):
    # Encode the values as a QR code
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
    qr.add_data('\r'.join(values))
    qr.make(fit=True)

    # Set the file name
    file_name = 'qr_code_1.png'

    # Check if the file already exists
    if os.path.isfile(file_name):
        # Find the next available file name
        i = 2
        while True:
            new_file_name = f"qr_code_{i}.png"
            if not os.path.isfile(new_file_name):
                break
            i += 1
        file_name = new_file_name

    # Save the QR code image with the updated file name
    img = qr.make_image(fill_color='black', back_color='white')
    img.save(file_name)
