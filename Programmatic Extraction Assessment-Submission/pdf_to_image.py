import fitz  # PyMuPDF
from PIL import Image
from PIL import Image
import pytesseract
import cv2
import re
import os



# - File Name, Full Name, Date of Birth, File No, Address, Address 2, City, State, ZIP
path = "C:\\Users\\nazid\\Downloads\\Programmatic Extraction Assessment\\Programmatic Extraction Assessment"
image_folder = 'C:\\Users\\nazid\\Downloads\\Programmatic Extraction Assessment\\Images'
file_no_pattern = r"(\d+)\s*\[On File"
zip_pattern = re.compile(r'\d{5}')
name_pattern = r"Page: \d+ (.+)"
address_pattern = r'(\d+\s+[A-Z]+\s+[A-Z]+\s*,\s*[A-Z]+\s*[A-Z]+)'

def convert_pdf_to_images(pdf_path, image_folder,img):
    pdf_document = fitz.open(pdf_path)
    # num_pages = pdf_document.page_count

    # for page_num in range(num_pages):
    page = pdf_document.load_page(1)
    image = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))  # Adjust the scaling as needed
    image.save(f"{image_folder}/page_{img}.png")

    pdf_document.close()

def rotate_image(image_path, output_path, angle):
    original_image = Image.open(image_path)
    
    # Calculate the new size of the image to fit the rotated version
    width, height = original_image.size
    rotated_image = original_image.rotate(angle, expand=True)
    new_width, new_height = rotated_image.size
    
    # Create a new image with the required size and paste the rotated image onto it
    new_image = Image.new('RGB', (new_width, new_height), (255, 255, 255))
    new_image.paste(rotated_image, (0, 0))
    
    new_image.save(output_path)

for name in os.listdir(path+ "\\Bucket 06"):
    # print(path + "\\Bucket 06\\" + name)
    file_path = path + "\\Bucket 06\\" + name
    img_name = str(name.split(".")[0])
    # print(str(name.split(".")[0]))
    convert_pdf_to_images(file_path,image_folder,img_name)

for name in os.listdir(image_folder):
    if name.endswith(".png"):
        image_path = image_folder + "\\" + name
        img_str=str(name.split(".")[0])
        output_path = image_folder + f"\\{img_str}.jpg"
        rotation_angle = -90  
        rotate_image(image_path, output_path, rotation_angle)
        







