import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import fitz 
from pdf2docx import Converter

def convert_pdf_to_docx(pdf_file, docx_file):
    try:
       
        pdf_document = fitz.open(pdf_file)
        doc = Converter(pdf_document)
        doc.convert(docx_file, start=0, end=None)
        doc.close()
        messagebox.showinfo("Success", f"Conversion successful: {pdf_file} -> {docx_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Conversion failed: {e}")

def browse_file():
    input_file_path = filedialog.askopenfilename(
        title="Select a PDF Document",
        filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
    )
    if input_file_path:
        output_file_path = input_file_path.replace(".pdf", ".docx")
        convert_pdf_to_docx(input_file_path, output_file_path)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw() 
                                                                                            


    main_frame = tk.Tk()
    main_frame.title("PDF to Word Converter")

    browse_button = tk.Button(main_frame, text="Browse", command=browse_file)
    browse_button.pack(pady=20)

    main_frame.mainloop()
