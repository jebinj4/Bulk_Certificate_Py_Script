import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import os

# Load data from Excel
data = pd.read_excel('certificate_data.xlsx')

# Load your certificate template
template_path = 'certificate_template.png'  # Path to your template image

# Define font paths and sizes
name_font_path = "GreatVibes-Regular.ttf"  # Path to your cursive font file
standard_font_path = "Arial.ttf"  # Path to your standard font file

name_font_size = 160
details_font_size = 100

# Define positions
name_position = (1451, 1011)   # Adjust as needed
course_position = (1265, 1400) # Adjust as needed
qr_position = (164, 1760)  # Adjust as needed

def generate_qr_code(data):
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    return img

def create_certificate(name, course, reg_no, date):
    # Open the certificate template
    cert = Image.open(template_path).convert("RGB")
    draw = ImageDraw.Draw(cert)

    # Load fonts
    try:
        name_font = ImageFont.truetype(name_font_path, name_font_size)
    except IOError:
        print(f"Font file not found: {name_font_path}. Using default font.")
        name_font = ImageFont.load_default()
        
    try:
        details_font = ImageFont.truetype(standard_font_path, details_font_size)
    except IOError:
        print(f"Font file not found: {standard_font_path}. Using default font.")
        details_font = ImageFont.load_default()

    # Add text to the certificate
    draw.text(name_position, name, font=name_font, fill=(169, 116, 42))  # Gold color for name
    draw.text(course_position, course.upper(), font=details_font, fill="black")  # Upper case for course name

    # Generate QR code
    qr_data = f"Name: {name}\nCourse: {course}\nRegistration Number: {reg_no}\nDate: {date}"
    qr_img = generate_qr_code(qr_data)
    qr_img = qr_img.resize((310, 310))  # Resize QR code if necessary

    # Paste QR code onto the certificate
    cert.paste(qr_img, qr_position)

    # Save the certificate
    cert.save(f"certificates/{name}_{course}.png")

# Ensure directories exist
os.makedirs("certificates", exist_ok=True)

# Iterate over the rows in the Excel file and generate certificates
for index, row in data.iterrows():
    name = row['Name']
    course = row['Course']
    reg_no = row['Registration Number']
    date = row['Date']
    create_certificate(name, course, reg_no, date)

print("Certificates generated successfully.")
