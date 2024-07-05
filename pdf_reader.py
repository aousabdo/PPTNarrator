import PyPDF2
import logging

def read_pdf_slides(pdf_path):
    slides = []
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                slides.append(page.extract_text())
        logging.info(f"Successfully read {len(slides)} slides from {pdf_path}")
    except Exception as e:
        logging.error(f"Error reading PDF {pdf_path}: {str(e)}")
    return slides