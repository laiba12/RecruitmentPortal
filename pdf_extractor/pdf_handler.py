from PyPDF2 import PdfReader
import docx
import textract
import win32com.client
import subprocess
import pythoncom
import os
def extract_text_from_pdf(file_path):
    pdf_document = PdfReader(file_path)
    text = ""
    for page in pdf_document.pages:
        text += page.extract_text()
    return text

def extract_text_from_docx(file_path):
    doc = docx.Document(file_path)
    text = "\n".join([para.text for para in doc.paragraphs])
    print(text)
    return text

# def extract_text_from_doc(file_path):
#     try:
#         text = textract.process(file_path).decode('utf-8')
#         return text
#     except Exception as e:
#         print(f"An error occurred while extracting text from {file_path}: {str(e)}")
#         return ""

# def extract_text_from_doc(file_path):
#     try:
#         # Use antiword to extract text from the .doc file
#         result = subprocess.run(['antiword', file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
#         if result.returncode != 0:
#             raise Exception(f"Antiword failed: {result.stderr}")
        
#         return result.stdout
#     except Exception as e:
#         raise Exception(f"Error extracting text from .doc file using antiword: {str(e)}")



# def extract_text_from_doc(file_path):
#     if not os.path.exists(file_path):
#         raise FileNotFoundError(f"File not found: {file_path}")

#     pythoncom.CoInitialize()
    
#     try:
#         word = win32com.client.Dispatch("Word.Application")
#         word.Visible = False

#         try:
#             doc = word.Documents.Open(file_path)
#             text = doc.Content.Text
#             doc.Close()
#             return text
#         except Exception as e:
#             raise Exception(f"Failed to open document: {e}")
#     finally:
#         word.Quit()
#         pythoncom.CoUninitialize()
       
 
