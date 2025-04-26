import tkinter as tk
from tkinter import filedialog
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

def highlight_paragraphs(filepath):
    doc = Document(filepath)
    for para in doc.paragraphs:
        if para.text.strip().startswith('<'):
            for run in para.runs:
                run.font.highlight_color = WD_COLOR_INDEX.GRAY_50  # Grey highlight
    new_filepath = filepath.replace('.docx', '_highlighted.docx')
    doc.save(new_filepath)
    print(f"Processed file saved as: {new_filepath}")

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        title="Select a DOCX file",
        filetypes=[("Word Documents", "*.docx")]
    )
    if file_path:
        highlight_paragraphs(file_path)
    else:
        print("No file selected.")

if __name__ == "__main__":
    main()
