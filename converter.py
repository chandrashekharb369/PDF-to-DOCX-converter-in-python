import fitz  
from docx import Document
import tkinter as tk
from tkinter import filedialog

def pdf_to_doc(pdf_path, docx_path):
    pdf_document = fitz.open(pdf_path)
    docx_document = Document()
    for page_num in range(pdf_document.page_count):
        # Get the text content of the page
        page = pdf_document[page_num]
        text = page.get_text()
        docx_document.add_paragraph(text)
    docx_document.save(docx_path)
    pdf_document.close()

def convert_pdf_to_doc():
    pdf_file_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF files", "*.pdf")])
    if pdf_file_path:
        docx_file_path = filedialog.asksaveasfilename(title="Save As", defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if docx_file_path:
            pdf_to_doc(pdf_file_path, docx_file_path)
            result_label.config(text="Conversion successful!")

root = tk.Tk()
root.title("PDF to DOCX Converter")

convert_button = tk.Button(root, text="Convert PDF to DOCX", command=convert_pdf_to_doc)
convert_button.pack(pady=20)

result_label = tk.Label(root, text="")
result_label.pack()
root.mainloop()
