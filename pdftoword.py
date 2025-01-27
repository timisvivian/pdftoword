from pdf2image import convert_from_path
from docx import Document
import pytesseract
import os

#PATHS can vary depending on user installation folder
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
os.environ['TESSDATA_PREFIX'] = r"C:/Program Files/Tesseract-OCR/tessdata"



def pdf_to_word_with_ocr(pdf_path, output_path):
    # Convert PDF pages to images
    images = convert_from_path(pdf_path)

    # Create a Word document
    doc = Document()

    # Perform OCR on each page
    for i, image in enumerate(images):
        text = pytesseract.image_to_string(image, lang='ron')  # Use Romanian language for OCR
        doc.add_paragraph(text)
        doc.add_page_break()

    # Save the Word document
    doc.save(output_path)

# Define input/output paths
pdf_files = [
    "Location/pdf_file_1.pdf",
    "Location/pdf_file_2.pdf",
    "Location/pdf_file_3.pdf",
]

#Choose where you want your converted doc to be available
output_dir = "Location"

for pdf_path in pdf_files:
    output_path = pdf_path.replace(".pdf", "_ocr.docx").replace("Location", output_dir)
    pdf_to_word_with_ocr(pdf_path, output_path)
    print(f"Converted with OCR: {output_path}")