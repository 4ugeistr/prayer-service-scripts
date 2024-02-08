import os
from docx import Document

def process_paragraph(paragraph):
    # Extract text from the paragraph
    text = []
    for run in paragraph.runs:
        text.append(run.text)
    text = ''.join(text)

    # Check if the paragraph contains Bible excerpt
    if any(word.isdigit() for word in text.split()):
        return f'<p>{text}</p>'
    else:
        return None

def process_docx(file_path):
    document = Document(file_path)
    html_content = []

    for paragraph in document.paragraphs:
        processed_paragraph = process_paragraph(paragraph)
        if processed_paragraph:
            html_content.append(processed_paragraph)

    return '\n'.join(html_content)

def save_as_html(html_content, output_file_path):
    with open(output_file_path, 'w', encoding='utf-8') as html_file:
        html_file.write(html_content)

def main():
    docx_folder = r'c:\Work\GitHub\prayer-service-scripts\flow\20240204_zachala'
    output_folder = docx_folder+r'\output'

    # Process each docx file in the folder
    for docx_file in os.listdir(docx_folder):
        if docx_file.endswith('.docx'):
            file_path = os.path.join(docx_folder, docx_file)
            html_content = process_docx(file_path)
            output_file_path = os.path.join(output_folder, f"{docx_file.split('.')[0]}.html")
            save_as_html(html_content, output_file_path)
            print(f"{docx_file} converted to HTML.")

if __name__ == "__main__":
    main()
