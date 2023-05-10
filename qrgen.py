import qrcode
import os
from pathlib import Path
from datetime import datetime

def codegen(values, name, i):
    # Encode the values as a QR code
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=10, border=4)
    qr.add_data('\r'.join(values))
    qr.make(fit=True)

    # Set the file name
    file_name = f"{name}_{i}.png"

    #Save the QR code
    desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop') 
    datetoday = datetime.today().strftime('%Y-%m-%d')
    pathtosave = f"{desktop}/sealed/sealed_{datetoday}/seal_{name}/qr"
    Path(pathtosave).mkdir(parents=True, exist_ok=True)
    file_path = f"{pathtosave}/qr_{file_name}"
    img = qr.make_image(fill_color='black', back_color='white')
    img.save(file_path)