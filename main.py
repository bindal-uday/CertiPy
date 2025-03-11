import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Load the template image
template_path = "./template.png"  # Path to your template image
template = Image.open(template_path)

# Font props
font_regular = "./fonts/Tomorrow-Regular.ttf"
font_semibold = "./fonts/Tomorrow-SemiBold.ttf"
font_bold = "./fonts/Tomorrow-Bold.ttf"
font_italic = "./fonts/Tomorrow-Italic.ttf"
font_black = "./fonts/Tomorrow-Black.ttf"

font_prefix = ImageFont.truetype(font_bold, 48)
font_name = ImageFont.truetype(font_black, 50)
font_department = ImageFont.truetype(font_regular, 44)
font_cert_number = ImageFont.truetype(font_semibold, 24)

# colors
Cyan="#12f1ff"
Coral="#ff4f12"
LightBlue="#12b3ff"
Turquoise="#12ffd4"
Violet="#a512ff"
LimeGreen="#a8ff12"
DarkGray="#333333"
LightGray="#f0f0f0"
Black="#000000"

# variables for prefix
prefix_code = "0x1-202502-"
prefix_name = "Mr./Ms. "

# Load the Excel file
excel_path = "./list.xlsx"  # Update with your Excel file path
df = pd.read_excel(excel_path)

# Function to create a certificate
def create_certificate(name, cert_number):

    # Create a copy of the template
    certificate = template.copy()
    draw = ImageDraw.Draw(certificate)

    common_y = 720
    template_width = certificate.width
    template_height = certificate.height

    prefix_bbox = draw.textbbox((0, 0), prefix_name, font=font_prefix)
    prefix_width = prefix_bbox[2] - prefix_bbox[0]

    name_bbox = draw.textbbox((0, 0), name, font=font_name)
    name_width = name_bbox[2] - name_bbox[0]

    cert_num_text = "Serial no: " + prefix_code + str(cert_number)
    cert_num_bbox = draw.textbbox((0, 0), cert_num_text, font=font_cert_number)
    cert_num_width = cert_num_bbox[2] - cert_num_bbox[0]
    cert_num_height = cert_num_bbox[3] - cert_num_bbox[1]

    # Calculate total width and center position
    total_width = prefix_width + name_width
    center_x = (template_width - total_width) // 2

    prefix_x = center_x 
    draw.text(
        (prefix_x, common_y),
        prefix_name,
        font=font_prefix,
        fill="DarkGray",
    )

    name_x = center_x + prefix_width
    draw.text(
        (name_x, common_y),
        name,
        font=font_name,
        fill="Cyan",
    )
    
    # Certificate no on right-bottom corner
    cert_number_position = (template_width - cert_num_width - 20, template_height - cert_num_height - 20)
    draw.text(
        cert_number_position,
        cert_num_text,
        font=font_cert_number,
        fill="LightGray",
    )

    # Save the certificate
    certificate.save(f'certificates/{name.replace(" ", "_")}_certificate.png')


# Create a directory for certificates if it doesn't exist
if not os.path.exists("certificates"):
    os.makedirs("certificates")

# Generate certificates for each student
for index, row in df.iterrows():
    student_name = row["Name"]
    cert_number = row["Certificate Number"]
    create_certificate(student_name, cert_number)

print("Certificates generated successfully!")
