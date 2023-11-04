import tkinter as tk
import ttkbootstrap as ttk
from tkinter import filedialog
import os
from docx2pdf import convert

def convert_docx_to_pdf(docx_file_paths):
    for docx_file_path in docx_file_paths:
        try:
            convert(docx_file_path)
            result_label.config(text=f"Conversion successful: {os.path.splitext(docx_file_path)[0]}.pdf")
        except Exception as e:
            result_label.config(text=f"Error: {str(e)}")

def browse_and_convert():
    docx_file_paths = filedialog.askopenfilenames(filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")])

    if docx_file_paths:
        convert_docx_to_pdf(docx_file_paths)

window = ttk.Window()
window.title("DEMO")
window.geometry('400x150')

window.title("Batch DOCX to PDF Converter")

convert_button = ttk.Button(window, text="Select and Convert DOCX to PDF", command=browse_and_convert)
convert_button.pack(pady=20)

result_label = ttk.Label(window, text="")
result_label.pack()

window.mainloop()
